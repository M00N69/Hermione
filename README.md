# Analyse de la Représentativité des Pesées sur Ligne - Conformité WELMEC

## Introduction

Ce code Python est conçu pour analyser la représentativité des pesées sur ligne de production et déterminer si elles sont conformes aux exigences de la norme WELMEC. Il utilise deux fichiers Excel :

* **Suivi Perte Matière.xlsx**: Contient des informations sur les quantités de matières premières utilisées pour chaque ordre de fabrication ("Of").
* **Suivi Acquisitions.xlsx**: Contient des informations sur les pesées effectuées sur la ligne de production, y compris la date, l'heure et le poids net du produit.

## Fonctionnement du code

1. **Chargement des fichiers Excel**: Le code commence par télécharger les deux fichiers Excel.
2. **Traitement des données**: Les données des deux fichiers sont combinées pour créer un DataFrame unique (`df_merged`) qui contient toutes les informations pertinentes.
3. **Calcul des pesées par heure**: Le code calcule le nombre de pesées effectuées pour chaque ordre de fabrication ("Of") par heure.
4. **Calcul des pesées totales**: Le code calcule le nombre total de pesées effectuées pour chaque ordre de fabrication.
5. **Calcul des statistiques de production**: Le code calcule les statistiques de la production, telles que le poids moyen, l'écart type et la durée de production pour chaque ordre.
6. **Calcul des pesées attendues**: Le code calcule le nombre de pesées attendu pour chaque ordre en fonction de la durée de production et de la moyenne des pesées par heure.
7. **Calcul des limites de tolérance WELMEC**: Le code calcule les limites de tolérance WELMEC (TNE, TU1, TU2) en fonction du poids moyen et de l'écart type des pesées.
8. **Affichage des statistiques**: Le code affiche les statistiques de production, le nombre de pesées attendues et les limites de tolérance WELMEC.
9. **Visualisations**: Le code crée des visualisations pour aider à identifier la représentativité des pesées :
    * **Histogramme des pesées par heure**: Affiche la distribution du nombre de pesées effectuées par heure.
    * **Graphique de dispersion des pesées attendues vs. réelles**: Compare le nombre de pesées attendu à celui réellement effectué pour chaque ordre.
    * **Boîte à moustaches de la distribution des pesées par OF**: Affiche la distribution des poids nets pour chaque ordre de fabrication, ce qui permet d'identifier les valeurs aberrantes et la variabilité.

## Interprétation des résultats

* **Histogramme des pesées par heure**: Un histogramme uniformément distribué indique une fréquence de pesées cohérente. Un histogramme avec des pics ou des creux importants peut suggérer une fréquence de pesées irrégulière.
* **Graphique de dispersion des pesées attendues vs. réelles**: Des points de données regroupés autour de la ligne de référence indiquent une bonne corrélation entre les pesées attendues et réelles. Les points de données éloignés de la ligne suggèrent des ordres où la fréquence de pesées peut ne pas être représentative.
* **Boîte à moustaches de la distribution des pesées par OF**: Les valeurs aberrantes en dehors des moustaches peuvent indiquer des pesées anormales. Une boîte à moustaches large ou des moustaches longues indiquent une grande variabilité dans les pesées, ce qui pourrait indiquer une fréquence de pesées incohérente ou une gamme plus large de poids réels pour l'ordre.

## Exigences WELMEC

Le code vérifie si les pesées sont conformes aux exigences suivantes de la norme WELMEC :

* **WELMEC 6.3 :**
    * **1.2**: Le nombre de pesées doit être suffisamment représentatif de la production.
    * **1.3**: Le prélèvement des échantillons doit être aléatoire.
    * **1.5**: L'efficacité du système de contrôle de la masse nette doit être comparable à la méthode de référence.
* **WELMEC 6.7 :**
    * **Appendix B1**: La taille de l'échantillon doit être proportionnelle à la taille du lot.
* **WELMEC 6.8 :**
    * **Section 3**: La méthode de détermination de la masse nette doit être précise et reproductible.

## Utilisation du code

1. Télécharger les fichiers Excel "Suivi Perte Matière.xlsx" et "Suivi Acquisitions.xlsx".
2. Exécuter le code Python.
3. Analyser les résultats pour déterminer si les pesées sont représentatives et conformes aux exigences WELMEC.

## Remarques

* **Adaptation du code**: Le code peut être adapté pour répondre aux besoins spécifiques de chaque production, par exemple en modifiant la sélection des ordres dans la boîte à moustaches ou en ajoutant d'autres visualisations.
* **Importance de la validation**: Il est important de valider les résultats du code avec des données réelles et de les comparer aux exigences spécifiques de chaque industrie ou réglementation.
* **Amélioration du processus**: L'utilisation de ce code et l'analyse des résultats peuvent aider à améliorer le processus de contrôle de la masse nette et à garantir la conformité à la norme WELMEC.
