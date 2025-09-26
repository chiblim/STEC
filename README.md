# STEC

Ce projet propose un pipeline complet d’analyse, de planification et d’agencement pour la zone STEC (zone de stockage temporaire des pièces déposées).
Il combine traitement de données, génération d’indicateurs, optimisation spatiale et interface utilisateur.

- Structure générale

Entrées : un fichier Excel (Input_STEC.xlsx) structuré en 5 feuilles (Extraction, Prat, NPP, Gares_criteres, Compétences).

Sorties : résultats sous forme de fichiers Excel (planning détaillé et résumé par période) et visualisation graphique de l’agencement.

Interface : une application graphique (CustomTkinter) permettant de collecter les URLs, générer le fichier d’entrée et lancer le pipeline.

Les dossiers d’entrée et de sortie (Input/, Output/) sont créés automatiquement.

- Pipeline de traitement
  
1. Lecture et préparation des données

lire_and_concat : lit les URLs dans le fichier principal (Input_STEC.xlsx), charge les fichiers associés, concatène les données et supprime les doublons si nécessaire.

Gestion des dates :

is_weekend : vérifie si une date tombe un week-end.

jour_ouvre : calcule le nombre de jours ouvrés entre deux dates.

format_date_courte : normalise les formats de date.

2. Analyse des transitions

process_transitions : détecte les déplacements des pièces entre emplacements.

Trie par identifiant d’engin et date.

Génère une ligne par transition avec durée en jours ouvrés.

if_stec_dynamique : compare la durée de chaque transition aux seuils définis par gare.

Indique si une pièce est en STEC (oui/non).

3. Analyse temporelle

generer_periodes_semaine : génère des intervalles hebdomadaires couvrant toute la période étudiée.

explode_par_periodes : associe chaque transition aux semaines concernées.

resumer_par_periode_et_support : calcule, par période et type de support :

le nombre de supports en STEC,

la surface totale occupée,

les besoins après prise en compte du gerbage (superposition des supports).

4. Optimisation de l’agencement physique

Contraintes spatiales :

est_dans_zone_interdite : évite les zones non utilisables (piliers, allées techniques).

chevauche_autre_rectangle : empêche le chevauchement entre supports.

Placement :

placer_rectangles_optimise : place les supports dans la zone disponible avec un algorithme heuristique bin packing 2D glouton (méthode « shelf »).

Visualisation :

dessiner_agencement : génère une image de l’agencement avec repères visuels.

5. Export des résultats

Excel :

resultat_avec_planning.xlsx : transitions détaillées semaine par semaine.

resultat_resume.xlsx : synthèse des surfaces et supports.

PNG : image de la proposition d’agencement.

 Interface graphique

Le module interface_stec permet d’interagir facilement avec le pipeline :

Générer un template Excel vierge.

Charger un fichier d’entrée existant.

Lancer le pipeline d’analyse.

Exporter un rapport PDF avec planning et agencement.

L’interface gère également les cas d’erreurs (fichier manquant, agencement non généré).

- Points clés

Algorithmes optimisés en O(n·log n) → scalables pour de grands volumes de données.

Résultats déterministes et donc reproductibles.

Conçu pour un usage industriel : lisible, robuste et adapté aux contraintes du terrain.
