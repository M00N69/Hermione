import streamlit as st
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np

st.title("Analyse de la Représentativité des Pesées sur Ligne - Conformité WELMEC")

# Upload des fichiers Excel
suivi_perte_matiere = st.file_uploader("Sélectionnez le fichier 'Suivi Perte Matière.xlsx'", type=['xlsx'])
suivi_acquisitions = st.file_uploader("Sélectionnez le fichier 'Suivi Acquisitions.xlsx'", type=['xlsx'])

# Vérifiez si les fichiers sont uploadés
if suivi_perte_matiere and suivi_acquisitions:
    # Chargement des fichiers si uploadés
    try:
        df_perte_matiere = pd.read_excel(suivi_perte_matiere, engine='openpyxl')
        df_acquisitions = pd.read_excel(suivi_acquisitions, engine='openpyxl')

        # Séparation de la colonne d'acquisition en date et heure
        df_acquisitions['Heure acquisition'] = pd.to_datetime(df_acquisitions['Heure acquisition'])
        df_acquisitions['Date'] = df_acquisitions['Heure acquisition'].dt.date
        df_acquisitions['Heure'] = df_acquisitions['Heure acquisition'].dt.time

        # Fusion des deux DataFrames
        df_merged = pd.merge(df_acquisitions, df_perte_matiere, on='Of')

        # Affichage des données fusionnées
        st.dataframe(df_merged)

        # Calcul du nombre de pesées par heure et par OF
        df_pesées_par_heure = df_merged.groupby(['Of', 'Date', 'Heure']).size().reset_index(name='Nombre de pesées')

        # Affichage des pesées par heure
        st.subheader("Nombre de pesées par heure")
        st.dataframe(df_pesées_par_heure)

        # Calcul du nombre de pesées par OF
        df_pesées_par_of = df_pesées_par_heure.groupby('Of').agg({'Nombre de pesées': 'sum'}).reset_index()

        # Affichage des pesées par OF
        st.subheader("Nombre total de pesées par OF")
        st.dataframe(df_pesées_par_of)

        # Calcul des statistiques de la production
        df_production_stats = df_merged.groupby('Of').agg({
            'Qté réelle': 'sum',
            'Qté lanc.': 'sum',
            'Valeur': ['mean', 'std']
        }).reset_index()
        df_production_stats.columns = ['Of', 'Qté réelle', 'Qté lanc.', 'Poids moyen (g)', 'Ecart type (g)']

        # Calcul de la durée de production (en heures)
        df_merged['Date de fin d\'OF'] = pd.to_datetime(df_merged['Date de fin d\'OF'])
        df_production_stats['Durée production (heures)'] = (
            df_merged.groupby('Of')['Date de fin d\'OF'].max() - 
            df_merged.groupby('Of')['Date de fin d\'OF'].min()
        ).dt.total_seconds() / 3600

        # Affichage des statistiques de production
        st.subheader("Statistiques de Production")
        st.dataframe(df_production_stats)

        # Calcul du nombre de pesées attendu
        df_production_stats['Pesées attendues'] = (
            df_pesées_par_heure.groupby('Of')['Nombre de pesées'].mean() * 
            df_production_stats['Durée production (heures)']
        ).values

        # Affichage du nombre de pesées attendu
        st.subheader("Pesées attendues vs. réelles")
        st.dataframe(df_production_stats)

        # Calcul des limites de tolérance WELMEC
        df_production_stats['TNE (g)'] = df_production_stats['Poids moyen (g)'] * 0.09
        df_production_stats['TU1 (g)'] = df_production_stats['Poids moyen (g)'] - df_production_stats['TNE (g)']
        df_production_stats['TU2 (g)'] = df_production_stats['Poids moyen (g)'] - 2 * df_production_stats['TNE (g)']

        # Affichage des limites de tolérance
        st.subheader("Limites de Tolérance WELMEC")
        st.dataframe(df_production_stats[['Of', 'TNE (g)', 'TU1 (g)', 'TU2 (g)']])

        # Histogramme des pesées
        st.subheader("Histogramme des Pesées")
        st.bar_chart(df_merged['Valeur'].value_counts().sort_index())

        # Affichage des justifications WELMEC
        st.subheader("Justifications WELMEC")

        st.markdown("**WELMEC 6.3 :**")
        st.markdown("- **1.2 :** Le nombre de pesées doit être suffisamment représentatif de la production. Cela signifie que le nombre de pesées doit être proportionnel à la durée de la production et à la taille du lot.")
        st.markdown("- **1.3 :** Le prélèvement des échantillons doit être aléatoire pour garantir la représentativité de la production.")
        st.markdown("- **1.5 :** L'efficacité du système de contrôle de la masse nette doit être comparable à la méthode de référence définie dans WELMEC 6.7.")

        st.markdown("**WELMEC 6.7 :**")
        st.markdown("- **Appendix B1 :** La taille de l'échantillon doit être proportionnelle à la taille du lot. Pour des lots de moins de 100 piluliers, un test de dépistage standard peut être utilisé.")

        st.markdown("**WELMEC 6.8 :**")
        st.markdown("- **Section 3 :** La méthode de détermination de la masse nette doit être précise et reproductible.")

        st.markdown("**Conclusions :**")
        st.markdown("- **Représentativité des pesées:** Comparez le nombre de pesées attendues avec le nombre de pesées réelles. Si la différence est importante, cela peut indiquer que la fréquence des pesées est insuffisante.")
        st.markdown("- **Conformité WELMEC:**  Le code Python calcule les limites de tolérance WELMEC (TNE, TU1, TU2) et permet de vérifier si les pesées sont conformes aux exigences WELMEC.")
        st.write("Utilisez ce code pour analyser vos données et justifier la conformité de vos contrôles de masse nette.")
        
    except ImportError as e:
        st.error(f"An error occurred while loading the Excel files: {e}. Please ensure that the 'openpyxl' library is installed.")
