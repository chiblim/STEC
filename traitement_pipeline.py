# === Imports ===
from os import makedirs, path
from pandas import read_excel, concat, merge, to_datetime, to_timedelta, Timedelta, DataFrame
from datetime import datetime, timedelta
from numpy import ceil
import customtkinter as ctk
import tkinter.messagebox
from openpyxl import Workbook
import matplotlib.pyplot as plt

# === Dossiers et fichiers ===
INPUT_FOLDER = "Input"
OUTPUT_FOLDER = "OUTPUT"
makedirs(OUTPUT_FOLDER, exist_ok=True)
chemin_input_STEC = path.join(INPUT_FOLDER, "Input_STEC.xlsx")

# === Fonctions ===
def lire_and_concat(chemin_excel: str,nom_feuil: str, nom_col: str, cols_a_garder: list, supp_doublons: bool):
    df_sheet = read_excel(chemin_excel, sheet_name=nom_feuil)

    if len(df_sheet[nom_col]) < 1:
        raise ValueError("Aucun url fourni dans la feuille de calcul")
    else:
        df_url_concated = DataFrame()
        for i in range(len(df_sheet[nom_col])):
            df_url_suivant = read_excel(df_sheet[nom_col].iloc[i], usecols=cols_a_garder)
            df_url_concated = concat([df_url_concated, df_url_suivant], ignore_index=True)

        if supp_doublons:
            df_url_concated = df_url_concated.drop_duplicates()
        return df_url_concated

def is_weekend(date_str):
    date_obj = datetime.strptime(date_str, "%Y-%m-%d")
    return date_obj.weekday() in [5, 6]

def jour_ouvre(d1, d2):
    d1 = datetime.strptime(d1, "%d/%m/%Y")
    d2 = datetime.strptime(d2, "%d/%m/%Y")

    s = 0
    n = (d2 - d1).days
    current_day = d1 + timedelta(days=1)
    for _ in range(n + 1):
        current_day = d1 + timedelta(days=1)
        if not is_weekend(current_day.strftime("%Y-%m-%d")):
            s += 1
    return s

def format_date_courte(date_obj):
    if isinstance(date_obj, str):
        date_obj = datetime.strptime(date_obj, "%Y-%m-%d")
    return date_obj.strftime("%d/%m/%Y")

def process_transitions(df):
    lignes = []
    df_sorted = df.sort_values(by=["id", "date_debut"])
    for id_val, group in df_sorted.groupby("id"):
        group = group.reset_index(drop=True)
        for i in range(len(group) - 1):
            ligne1 = group.loc[i]
            ligne2 = group.loc[i + 1]
            d1 = format_date_courte(ligne1["date_debut"])
            d2 = format_date_courte(ligne2["date_debut"])
            lignes.append({
                "id": id_val,
                "gare_depart": ligne1["emplacement wms entrée"],
                "gare_arrivee": ligne2["emplacement wms entrée"],
                "date_debut": d1,
                "duree": jour_ouvre(d1, d2),
                "engin": ligne1["Engin"],
                "caisse": ligne1["Caisse"],
                "kit": ligne1["Kit"]
            })
    result_df = DataFrame(lignes)
    result_df["date_debut"] = to_datetime(result_df["date_debut"], format="%d/%m/%Y")
    result_df = result_df.sort_values(by=["date_debut", "id"]).reset_index(drop=True)
    result_df["ordre_changement"] = result_df.groupby("id").cumcount() + 1
    return result_df

def if_stec_dynamique(df_transitions, df_seuils):
    df_seuils.columns = [col.strip().lower().replace(" ", "_") for col in df_seuils.columns]
    df = df_transitions.copy()
    df_merged = df.merge(
        df_seuils,
        how="left",
        left_on="gare_depart",
        right_on="gare",
        suffixes=("", "_seuil")
    )
    missing_gare_mask = df_merged["duree_min_pour_stec"].isna()
    if missing_gare_mask.any():
        default_seuils = df_seuils[df_seuils["gare"] == "Autre"].squeeze()
        if default_seuils.empty:
            raise ValueError("La gare par défaut 'Autre' n'est pas présente dans le répertoire des seuils.")
        for col in ["duree_min_pour_stec", "duree_min_pour_stec_si_mm_gare"]:
            df_merged.loc[missing_gare_mask, col] = default_seuils[col]
    df_merged["seuil"] = df_merged.apply(
        lambda row: row["duree_min_pour_stec_si_mm_gare"]
        if row["gare_depart"] == row["gare_arrivee"]
        else row["duree_min_pour_stec"],
        axis=1
    )
    df_merged["STEC ?"] = df_merged["duree"] > df_merged["seuil"]
    return df_merged.drop(columns=["gare", "duree_min_pour_stec", "duree_min_pour_stec_si_mm_gare", "seuil"])

def generer_periodes_semaine(df, date_col="date_debut"):
    df[date_col] = to_datetime(df[date_col])
    date_min = df[date_col].min()
    date_max = df[date_col].max()
    date_min = date_min - Timedelta(days=date_min.weekday())
    date_max = date_max + Timedelta(days=(6 - date_max.weekday()))
    periodes = []
    current_start = date_min
    periode_num = 1
    while current_start <= date_max:
        current_end = current_start + Timedelta(days=6)
        periodes.append({
            "periode": periode_num,
            "date_debut": current_start,
            "date_fin": current_end
        })
        periode_num += 1
        current_start += Timedelta(days=7)
    return DataFrame(periodes)

def explode_par_periodes(df_transitions, df_periodes):
    lignes = []
    for _, row in df_transitions.iterrows():
        date_debut = row["date_debut"]
        date_fin = date_debut + to_timedelta(row["duree"], unit="D")
        periodes_concernees = df_periodes[
            (df_periodes["date_fin"] >= date_debut) &
            (df_periodes["date_debut"] <= date_fin)
        ]
        for _, p in periodes_concernees.iterrows():
            ligne = row.copy()
            ligne["periode"] = p["periode"]
            lignes.append(ligne)
    return DataFrame(lignes)

def resumer_par_periode_et_support(df):
    import numpy as np
    df_stec = df[df["STEC ?"] == True].copy()
    grouped = df_stec.groupby(["periode", "Code contenant", "Gerbage", "Surface"], as_index=False).agg({
        "id": "count"
    })
    grouped.rename(columns={
        "periode": "Periode (sem)",
        "Code contenant": "Type support",
        "id": "nb supports en STEC",
        "Gerbage": "gerbage"
    }, inplace=True)
    grouped["surface totale"] = grouped["Surface"] * grouped["nb supports en STEC"]
    grouped["surface après gerbage"] = np.where(
        (grouped["gerbage"] == 0) | (grouped["gerbage"].isna()),
        grouped["surface totale"],
        ceil(grouped["nb supports en STEC"] / grouped["gerbage"]) * grouped["Surface"]
    )
    grouped["nb supports après gerbage"] = np.where(
        (grouped["gerbage"] == 0) | (grouped["gerbage"].isna()),
        grouped["nb supports en STEC"],
        ceil(grouped["nb supports en STEC"] / grouped["gerbage"])
    )
    return grouped[
        ["Periode (sem)", "Type support", "nb supports en STEC", "Surface", "surface totale",
         "gerbage", "surface après gerbage", "nb supports après gerbage"]
    ]
def est_dans_zone_interdite(x, y, largeur, hauteur, zones_interdites):
    for zone in zones_interdites:
        xi = zone["x"]
        yi = zone["y"]
        wi = zone["width"]
        hi = zone["height"]
        if not (x + largeur <= xi or x >= xi + wi or y + hauteur <= yi or y >= yi + hi):
            return True
    return False


def chevauche_autre_rectangle(x, y, largeur, hauteur, positions, marge_passage):
    for x2, y2, w2, h2, _ in positions:
        if not (x + largeur + marge_passage <= x2 or x >= x2 + w2 + marge_passage or
                y + hauteur + marge_passage <= y2 or y >= y2 + h2 + marge_passage):
            return True
    return False


def placer_rectangles_optimise(rectangles, largeur_zone, hauteur_zone, zone_interdite, marge=0.1, marge_passage=0.5):
    positions = []
    y_courant = 0

    rectangles = sorted(rectangles, key=lambda r: r["hauteur"], reverse=True)

    for rect in rectangles:
        largeur = rect["largeur"]
        hauteur = rect["hauteur"]
        nom = rect["nom"]

        place = False
        y = y_courant
        while y + hauteur <= hauteur_zone:
            x = 0
            while x + largeur <= largeur_zone:
                if est_dans_zone_interdite(x, y, largeur, hauteur, zone_interdite):
                    x += 0.5
                    continue
                if chevauche_autre_rectangle(x, y, largeur, hauteur, positions, marge_passage):
                    x += 0.5
                    continue
                positions.append((x, y, largeur, hauteur, nom))
                place = True
                break
            if place:
                break
            y += 0.5
        if not place:
            print(f" Le rectangle {nom} n'a pas pu être placé.")
    return positions


def dessiner_agencement(positions, zone_largeur, zone_hauteur, zones_interdites, save_path=None):
    fig, ax = plt.subplots()
    ax.set_xlim(0, zone_largeur)
    ax.set_ylim(0, zone_hauteur)

    for zone in zones_interdites:
        ax.add_patch(plt.Rectangle((zone["x"], zone["y"]), zone["width"], zone["height"], color='red', alpha=0.3))

    for x, y, w, h, nom in positions:
        ax.add_patch(plt.Rectangle((x, y), w, h, edgecolor='black', facecolor='blue', alpha=0.6))
        ax.text(x + w / 2, y + h / 2, nom, ha='center', va='center', fontsize=8)

    ax.set_aspect('equal')
    plt.gca().invert_yaxis()
    plt.title("Proposition d'agencement de la STEC", fontsize=14, fontweight='bold')

    if save_path:
        plt.savefig(save_path)
        plt.close(fig)
    else:
        plt.show()





def lancer_pipeline_agencement(chemin_excel,event=None):
    

    # === Traitements ===
    colonnes_extraction = [
        "Compétence", "Engin", "Caisse", "Kit", "Code tâche",
        "date_debut", "emplacement wms entrée"
    ]

    df_input_extraction = lire_and_concat(chemin_excel,"Extraction", "URL extraction", colonnes_extraction, False)

    df_input_competences = lire_and_concat(chemin_excel,"Compétences", "Compétences STEC", ["Compétences STEC"], True)
    df_input_extraction = df_input_extraction[df_input_extraction["Compétence"].isin(df_input_competences["Compétences STEC"])]
    df_input_extraction["id"] = (
        df_input_extraction["Engin"].astype(str) + "_" +
        df_input_extraction["Caisse"].astype(str) + "_" +
        df_input_extraction["Kit"].astype(str)
    )

    df_input_prat = lire_and_concat(chemin_excel,"Prat", "URL prat", ["Nom du Praticable", "Surface", "Gerbage"], True)
    df_input_npp = lire_and_concat(chemin_excel,"NPP", "URL NPP", ["Code article", "Code contenant"], True)

    df_input_ctlg_ctnt = df_input_npp.merge(
        df_input_prat, how="left", left_on="Code contenant", right_on="Nom du Praticable"
    )
    df_input_ctlg_ctnt = df_input_ctlg_ctnt.drop(columns=["Nom du Praticable"])

    df_extraction = process_transitions(df_input_extraction)

    df_gare_criteres = lire_and_concat(chemin_excel,
        "Gares_criteres", "URL Gares et critères STEC",
        ["gare", "duree min pour STEC", "duree min pour STEC si mm gare"], True
    )

    df_extraction_stec = if_stec_dynamique(df_extraction, df_gare_criteres)
    df_extraction_stec = df_extraction_stec[df_extraction_stec["STEC ?"] == True]

    df_extraction_stec = df_extraction_stec.merge(
        df_input_ctlg_ctnt, how="left", left_on="kit", right_on="Code article"
    ).drop(columns=["Code article"])

    df_periodes = generer_periodes_semaine(df_extraction_stec)
    df_extraction_stec_planif = explode_par_periodes(df_extraction_stec, df_periodes)
    df_extraction_stec_resum = resumer_par_periode_et_support(df_extraction_stec_planif)

    # === Export des résultats ===
    df_extraction_stec_planif.to_excel(path.join(OUTPUT_FOLDER, "resultat_avec_planning.xlsx"))
    df_extraction_stec_resum.to_excel(path.join(OUTPUT_FOLDER, "resultat_resume.xlsx"))

    df_max_supports = df_extraction_stec_resum.groupby("Type support", as_index=False).agg({
        "nb supports après gerbage": "max",
        "surface après gerbage": "max"
    })
    df_max_supports["nb supports après gerbage"] = df_max_supports["nb supports après gerbage"].astype(int)


    liste_contenants = []
    for _, row in df_max_supports.iterrows():
        nom = row["Type support"]
        surface = row["surface après gerbage"]
        largeur = hauteur = surface ** 0.5  # Hypothèse : carré
        liste_contenants.append({
            "nom": nom,
            "largeur": largeur,
            "hauteur": hauteur,
            "surface": surface
        })

    zone_interdite = [
        {"x": 7, "y": 10, "width": 3, "height": 40},     # barre gauche
        {"x": 30, "y": 10, "width": 3, "height": 40},    # barre droite
        {"x": 7, "y": 10, "width": 26, "height": 3}      # barre horizontale
    ]

    positions = placer_rectangles_optimise(
        rectangles=liste_contenants,
        largeur_zone=40,
        hauteur_zone=50,
        zone_interdite=zone_interdite,
        marge=0.1,
        marge_passage=0.5  # Espace de sécurité entre rectangles
    )

    image_path = path.join(OUTPUT_FOLDER, "agencement.png")
    dessiner_agencement(positions, 40, 50, zone_interdite, save_path=image_path)

    return positions, df_max_supports
