import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt

st.title("Analyse de la Représentativité des Pesées sur Ligne - Conformité WELMEC")

# Upload des fichiers Excel
suivi_perte_matiere = st.file_uploader("Sélectionnez le fichier 'Suivi Perte Matière.xlsx'", type=['xlsx'])
suivi_acquisitions = st.file_uploader("Sélectionnez le fichier 'Suivi Acquisitions.xlsx'", type=['xlsx'])

# Vérifiez si les fichiers sont uploadés
if suivi_perte_matiere and suivi_acquisitions:
    try:
        # Chargement des fichiers si uploadés
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

        # Calcul du nombre de pesées par OF (Correction)
        df_pesées_par_of = df_merged.groupby('Of').size().reset_index(name='Nombre de pesées')  # Regroupement uniquement par OF

        # Affichage des pesées par OF
        st.subheader("Nombre total de pesées par OF")
        st.dataframe(df_pesées_par_of)

        # Calcul des statistiques de la production
        df_production_stats = df_merged.groupby('Of').agg({
            'Qté réelle': 'sum',
            'Qté lanc.': 'sum',
            'Mesure valeur': ['mean', 'std']  # Using 'Mesure valeur' as the column for calculations
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

        #  ---  Key Metrics for Assessing Sampling Representativeness --- 
        
        # 1. Mean Number of Weighings Per Hour:
        mean_weighings_per_hour = df_pesées_par_heure['Nombre de pesées'].mean()
        st.markdown("**Moyenne des pesées par heure:** {}".format(mean_weighings_per_hour))

        # 2. Standard Deviation of Weighings Per Hour:
        std_weighings_per_hour = df_pesées_par_heure['Nombre de pesées'].std()
        st.markdown("**Écart type des pesées par heure:** {}".format(std_weighings_per_hour))

        # 3.  Average Production Time Per Order:
        average_production_time = df_production_stats['Durée production (heures)'].mean()
        st.markdown("**Durée moyenne de production par OF:** {:.2f} heures".format(average_production_time))

        # 4.  Percent of Orders With 'Expected Weighings' Close to Actual Weighings:
        close_to_expected = df_production_stats[
            abs(df_production_stats['Pesées attendues'] - df_production_stats['Nombre de pesées']) < 0.1 * df_production_stats['Pesées attendues']
        ]
        percentage_close = len(close_to_expected) / len(df_production_stats) * 100
        st.markdown("**Pourcentage des OF avec un nombre de pesées proche du nombre attendu:** {:.1f}%".format(percentage_close))

        # ---  Visualizations ---
        
        #  1.  Histogram of Weighings per Hour:
        st.subheader("Histogramme des pesées par heure")
        plt.figure(figsize=(10, 6))  # Set figure size
        plt.hist(df_pesées_par_heure['Nombre de pesées'], bins=10, edgecolor='black')
        plt.xlabel("Nombre de pesées")
        plt.ylabel("Fréquence")
        plt.title("Distribution des pesées par heure")
        st.pyplot(plt)

        #  2.  Scatter Plot of Expected vs. Actual Weighings:
        st.subheader("Pesées attendues vs. réelles")
        plt.figure(figsize=(10, 6))  # Set figure size
        plt.scatter(df_production_stats['Pesées attendues'], df_production_stats['Nombre de pesées'])
        plt.xlabel("Pesées attendues")
        plt.ylabel("Nombre de pesées réelles")
        plt.title("Comparaison des pesées attendues et réelles")
        st.pyplot(plt)

        #  3.  Box Plot of Weighing Distribution for Each Order (with error handling):
        try:
            st.subheader("Distribution des Pesées par OF")
            plt.figure(figsize=(10, 6))

            # Select a limited number of orders for the box plot (e.g., the first 5)
            selected_orders = df_merged['Of'].unique()[:5]  

            # Filter the data to include only selected orders
            filtered_data = df_merged[df_merged['Of'].isin(selected_orders)]

            # Create the box plot directly using the 'Mesure valeur' column
            plt.boxplot(filtered_data['Mesure valeur'], labels=filtered_data['Of'].unique())

            plt.xlabel("OF")
            plt.ylabel("Poids net (g)")
            plt.title("Distribution des pesées par ordre de fabrication (5 premiers OF)")
            st.pyplot(plt)
        except Exception as e:
            st.warning(f"Erreur lors de la création du graphique de boîte à moustaches: {e}.")
            st.write("Veuillez vérifier vos données et réessayer.")

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
        st.markdown("- **Conformité WELMEC:** Le code Python calcule les limites de tolérance WELMEC (TNE, TU1, TU2) et permet de vérifier si les pesées sont conformes aux exigences WELMEC.")
        st.write("Utilisez ce code pour analyser vos données et justifier la conformité de vos contrôles de masse nette.")
        
    except ImportError as e:
        st.error(f"An error occurred while loading the Excel files: {e}. Please ensure that the 'openpyxl' library is installed.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
