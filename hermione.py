import streamlit as st
import pandas as pd
from datetime import datetime

st.title("Analyse de la Représentativité des Pesées sur Ligne")

# Upload des fichiers Excel
suivi_perte_matiere = st.file_uploader("Sélectionnez le fichier 'Suivi Perte Matière.xlsx'", type=['xlsx'])
suivi_acquisitions = st.file_uploader("Sélectionnez le fichier 'Suivi Acquisitions.xlsx'", type=['xlsx'])

# Chargement des fichiers si uploadés
if suivi_perte_matiere and suivi_acquisitions:
    df_perte_matiere = pd.read_excel(suivi_perte_matiere)
    df_acquisitions = pd.read_excel(suivi_acquisitions)

    # Séparation de la colonne d'acquisition en date et heure
    df_acquisitions['Date'] = pd.to_datetime(df_acquisitions['Heure acquisition']).dt.date
    df_acquisitions['Heure'] = pd.to_datetime(df_acquisitions['Heure acquisition']).dt.time

    # Fusion des deux DataFrames
    df_merged = pd.merge(df_acquisitions, df_perte_matiere, left_on='Of', right_on='Of')

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
    df_production_stats = df_merged.groupby('Of').agg({'Qté réelle': 'sum', 'Qté lanc.': 'sum'}).reset_index()
    df_production_stats['Durée production (heures)'] = (pd.to_datetime(df_production_stats['Date de fin d'OF']) - pd.to_datetime(df_production_stats['Date de fin d'OF']).dt.date).dt.total_seconds() / 3600

    # Affichage des statistiques de production
    st.subheader("Statistiques de Production")
    st.dataframe(df_production_stats)

    # Calcul du nombre de pesées attendu
    df_production_stats['Pesées attendues'] = df_pesées_par_heure.groupby('Of').agg({'Nombre de pesées': 'mean'})['Nombre de pesées'] * df_production_stats['Durée production (heures)']

    # Affichage du nombre de pesées attendu
    st.subheader("Pesées attendues vs. réelles")
    st.dataframe(df_production_stats)

    # Analyse et affichage des conclusions
    st.subheader("Conclusions")
    st.write("**Analyse de la représentativité des pesées:**")
    st.write("Pour chaque OF, comparez le nombre de pesées attendues avec le nombre de pesées réelles.  ")
    st.write("Si la différence est importante, cela peut indiquer que la fréquence des pesées est insuffisante. ")
    st.write("Vous pouvez ajuster la consigne de prélèvement et contrôle en fonction de ces données.")
