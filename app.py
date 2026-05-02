import streamlit as st
import pandas as pd
import pdfplumber
import xlsxwriter
from io import BytesIO

# Configuration de la page
st.set_page_config(page_title="Cabinet Oueslati - ComptaAuto", layout="wide")

def find_client_account(client_name, plan):
    """
    Recherche le compte comptable d'un client dans le plan de comptes.
    """
    # Ici, nous supposons que vous avez une logique de scoring (ex: fuzzy matching)
    # Pour l'exemple, nous simulons un calcul de scores simple
    # Remplacez cette partie par votre logique exacte si elle diffère
    scores = plan["Nom"].apply(lambda x: 100 if str(client_name).lower() in str(x).lower() else 0)

    # CORRECTION DE L'INDENTATION ET DE LA LOGIQUE
    if not scores.empty and scores.max() > 0:
        return plan.iloc[scores.idxmax()]["Compte"]
    else:
        # Retourne un compte par défaut si aucun match n'est trouvé
        return "41100000"

def main():
    st.title("🚀 ComptaAuto Studio")
    st.subheader("Automatisation des écritures comptables")

    # Barre latérale pour les paramètres
    st.sidebar.header("Paramètres")
    plan_file = st.sidebar.file_uploader("Charger le Plan de Comptes (Excel)", type=['xlsx'])

    # Zone principale
    uploaded_file = st.file_uploader("Glissez vos factures (PDF)", type=['pdf'])

    if uploaded_file and plan_file:
        # Lecture du plan de comptes
        plan_df = pd.read_excel(plan_file)
        
        st.success("Fichiers chargés avec succès !")
        
        # Simulation de traitement (Remplacez par votre logique complète)
        with pdfplumber.open(uploaded_file) as pdf:
            # Exemple : extraction du nom du client (à adapter selon vos factures)
            first_page = pdf.pages[0]
            text = first_page.extract_text()
            # On imagine extraire un nom de client ici
            client_name_extracted = "Client Exemple" 

        # Appel de la fonction corrigée
        compte = find_client_account(client_name_extracted, plan_df)
        
        st.write(f"### Résultat de l'analyse")
        st.info(f"Client détecté : **{client_name_extracted}** ⮕ Compte : **{compte}**")

        # Affichage d'un tableau éditable pour correction manuelle
        data = {
            "Date": ["2026-05-01"],
            "Libellé": [f"Facture {client_name_extracted}"],
            "Compte": [compte],
            "Débit": [120.0],
            "Crédit": [0.0]
        }
        df_final = pd.DataFrame(data)
        
        st.write("Vérifiez et modifiez les écritures si nécessaire :")
        edited_df = st.data_editor(df_final, num_rows="dynamic")

        # Bouton de téléchargement Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            edited_df.to_excel(writer, index=False, sheet_name='Journal')
        
        st.download_button(
            label="📥 Télécharger le Journal Excel",
            data=output.getvalue(),
            file_name="ecritures_comptables.xlsx",
            mime="application/vnd.ms-excel"
        )

if __name__ == "__main__":
    main()
