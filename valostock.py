import streamlit as st
import pandas as pd
import os
from datetime import datetime
import matplotlib.pyplot as plt
from pathlib import Path

st.set_page_config(page_title="Valorisation des Stocks", layout="wide")
st.title("📊 Dashboard - Valorisation des stocks par magasin et fournisseur")

# Chemins dynamiques vers Dropbox partagé (compatible multi-postes)
dropbox_root = Path.home() / "SGDF Dropbox" / "Dossier de l'équipe SGDF" / "La Tribu" / "05_Systeme Information" / "Keyneo"
PRODUCT_PATH = dropbox_root / "Import produits" / "Database_WEB.xlsx"
HISTO_FILE = dropbox_root / "Valorisation Stock" / "historique_valorisation.csv"

# Création ou chargement de l'historique
if os.path.exists(HISTO_FILE):
    historique_df = pd.read_csv(HISTO_FILE)
else:
    historique_df = pd.DataFrame(columns=["date", "organisationId", "brand", "valorisation"])

st.sidebar.header("📂 Importer les fichiers")
stock_files = st.sidebar.file_uploader("Fichiers d'export Keyneo (un par magasin)", type=["csv"], accept_multiple_files=True)

# Choix de la vue du rapport
view_mode = st.sidebar.radio("Vue du rapport", ["Par fournisseur puis magasin", "Par magasin puis fournisseur"])

if stock_files:
    # Charger tous les fichiers stock et les concaténer
    stock_list = []
    for file in stock_files:
        df = pd.read_csv(file, sep=';')
        stock_list.append(df)
    stocks_df = pd.concat(stock_list, ignore_index=True)

    # Chargement de la base produit : automatique ou manuelle si introuvable
    if PRODUCT_PATH.exists():
        products_df = pd.read_excel(PRODUCT_PATH)
        st.success("Base produit chargée automatiquement depuis Dropbox.")
    else:
        st.warning("Fichier base produit introuvable dans Dropbox. Veuillez l'importer manuellement ci-dessous :")
        product_file = st.file_uploader("Base produit (Excel)", type=["xls", "xlsx"])
        if product_file:
            products_df = pd.read_excel(product_file)
            st.success("Base produit chargée manuellement.")
        else:
            st.stop()

    # Nettoyage base produit
    products_df = products_df.rename(columns={"SKU": "sku", "PurchasingPrice": "purchasing_price", "Brand": "brand"})
    products_df = products_df[["sku", "purchasing_price", "brand"]]
    stocks_df["sku"] = stocks_df["sku"].astype(str)
    products_df["sku"] = products_df["sku"].astype(str)

    # Fusion des données
    merged_df = pd.merge(stocks_df, products_df, on="sku", how="left")
    merged_df["valorisation"] = merged_df["quantity"] * merged_df["purchasing_price"]

    # Date d'import = aujourd'hui
    date_import = datetime.today().strftime('%Y-%m-%d')

    # Rapport principal
    report_df = merged_df.groupby(["organisationId", "brand"], as_index=False)["valorisation"].sum()
    report_df.insert(0, "date", date_import)
    report_df["valorisation"] = report_df["valorisation"].map(lambda x: round(x, 2))

    # Supprimer les lignes avec valorisation ≤ 0
    report_df = report_df[report_df["valorisation"] > 0]

    # Supprimer les doublons exacts dans le rapport du jour
    report_df = report_df.drop_duplicates()

    # Supprimer les anciennes valeurs du jour dans l'historique pour ne pas dupliquer
    historique_df = historique_df[~(
        (historique_df["date"] == date_import) &
        (historique_df["organisationId"].isin(report_df["organisationId"])) &
        (historique_df["brand"].isin(report_df["brand"]))
    )]

    # Créer le dossier final uniquement si nécessaire
    if not HISTO_FILE.parent.exists():
        try:
            HISTO_FILE.parent.mkdir(parents=False, exist_ok=True)
        except Exception as e:
            st.error(f"Erreur lors de la création du dossier de destination : {e}")
            st.stop()

    # Mise à jour de l'historique (Dropbox)
    historique_df = pd.concat([historique_df, report_df], ignore_index=True)
    historique_df.to_csv(HISTO_FILE, index=False)

    # Filtres dynamiques
    st.sidebar.header("🔍 Filtres")
    dates = historique_df["date"].unique().tolist()
    magasins = historique_df["organisationId"].unique().tolist()
    fournisseurs = historique_df["brand"].unique().tolist()

    selected_date = st.sidebar.selectbox("Date", sorted(dates, reverse=True))
    selected_magasin = st.sidebar.multiselect("Magasins", magasins, default=magasins)
    selected_brand = st.sidebar.multiselect("Fournisseurs", fournisseurs, default=fournisseurs)

    filtered_df = historique_df[
        (historique_df["date"] == selected_date) &
        (historique_df["organisationId"].isin(selected_magasin)) &
        (historique_df["brand"].isin(selected_brand))
    ]

    # Supprimer les lignes avec valorisation ≤ 0 dans le rapport filtré
    filtered_df = filtered_df[filtered_df["valorisation"] > 0]

    # Tri selon le mode choisi + tri par valorisation décroissante
    if view_mode == "Par fournisseur puis magasin":
        filtered_df = filtered_df.sort_values(by=["brand", "organisationId", "valorisation"], ascending=[True, True, False])
    else:
        filtered_df = filtered_df.sort_values(by=["organisationId", "brand", "valorisation"], ascending=[True, True, False])

    # Affichage tableau principal
    st.subheader(f"Valorisation du {selected_date}")
    st.dataframe(filtered_df, use_container_width=True)

    # Téléchargement du rapport filtré en Excel
    excel_file = f"valorisation_{selected_date}.xlsx"
    filtered_df.to_excel(excel_file, index=False)
    with open(excel_file, "rb") as f:
        st.download_button(
            label="📂 Télécharger le rapport filtré (Excel)",
            data=f,
            file_name=excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Historique complet
    st.subheader("🗓️ Historique complet")
    st.dataframe(historique_df.sort_values("date", ascending=False), use_container_width=True)

else:
    st.info("Merci d'importer les fichiers de stock (un par magasin) pour générer le rapport. La base produit est chargée automatiquement depuis Dropbox ou importable manuellement.")
