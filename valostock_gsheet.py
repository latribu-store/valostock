import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
import gspread
from google.oauth2 import service_account
import smtplib
from email.message import EmailMessage

st.set_page_config(page_title="Valorisation vers Google Sheets", layout="wide")
st.title("📊 App - Valorisation vers Google Sheets (Looker Ready)")

# Config fichiers
HISTO_FILE = "historique_valorisation.csv"
SPREADSHEET_ID = "1lOtH16m_xs1-EzQ7D_tp8wZz3fZu2eTbLQFU099MSNw"
SHEET_NAME = "Données"
LOOKER_URL = "https://lookerstudio.google.com/s/i1sjkqxFJro"

# Email config depuis secrets.toml ou Streamlit Cloud
SMTP_SERVER = st.secrets["email"]["smtp_server"]
SMTP_PORT = st.secrets["email"]["smtp_port"]
SMTP_USER = st.secrets["email"]["smtp_user"]
SMTP_PASSWORD = st.secrets["email"]["smtp_password"]
DEFAULT_RECEIVER = st.secrets["email"]["receiver"]

st.sidebar.header("📂 Importer les fichiers")
stock_files = st.sidebar.file_uploader("Fichiers de stock (un par magasin)", type=["csv"], accept_multiple_files=True)
product_file = st.sidebar.file_uploader("Base produit (Excel)", type=["xls", "xlsx"])
emails_supp = st.sidebar.text_input("📧 Autres destinataires (séparés par des virgules)")

if stock_files and product_file:
    stock_list = [pd.read_csv(f, sep=';') for f in stock_files]
    stocks_df = pd.concat(stock_list, ignore_index=True)
    products_df = pd.read_excel(product_file)

    products_df = products_df.rename(columns={"SKU": "sku", "PurchasingPrice": "purchasing_price", "Brand": "brand"})
    products_df = products_df[["sku", "purchasing_price", "brand"]]

    stocks_df["sku"] = stocks_df["sku"].astype(str)
    products_df["sku"] = products_df["sku"].astype(str)

    merged_df = pd.merge(stocks_df, products_df, on="sku", how="left")
    merged_df["valorisation"] = merged_df["quantity"] * merged_df["purchasing_price"]

    date_import = datetime.today().strftime('%d-%m-%Y')
    report_df = merged_df.groupby(["organisationId", "brand"], as_index=False)["valorisation"].sum()
    report_df.insert(0, "date", datetime.today().strftime('%Y-%m-%d'))
    report_df = report_df[report_df["valorisation"] > 0].drop_duplicates()
    report_df["valorisation"] = report_df["valorisation"].round(2)

    if os.path.exists(HISTO_FILE):
        historique_df = pd.read_csv(HISTO_FILE)
    else:
        historique_df = pd.DataFrame(columns=["date", "organisationId", "brand", "valorisation"])

    historique_df = pd.concat([
        historique_df[~(
            (historique_df["date"] == datetime.today().strftime('%Y-%m-%d')) &
            (historique_df["organisationId"].isin(report_df["organisationId"])) &
            (historique_df["brand"].isin(report_df["brand"]))
        )],
        report_df
    ])

    historique_df["date"] = pd.to_datetime(historique_df["date"])
    latest_date = historique_df["date"].max()
    historique_df["est_derniere_date"] = historique_df["date"] == latest_date
    historique_df["date"] = historique_df["date"].dt.strftime("%Y-%m-%d")

    historique_df.to_csv(HISTO_FILE, index=False)

    st.success(f"✅ Données ajoutées à l'historique ({len(report_df)} lignes)")

    if st.button("📤 Mettre à jour Google Sheets + envoyer par e-mail"):
        try:
            # MAJ Google Sheets
            scopes = ["https://www.googleapis.com/auth/spreadsheets"]
            gcp_service_account_info = json.loads(st.secrets["gcp_service_account"])
            creds = service_account.Credentials.from_service_account_info(gcp_service_account_info, scopes=scopes)
            client = gspread.authorize(creds)
            sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)

            sheet.clear()
            data = [historique_df.columns.tolist()] + historique_df.values.tolist()
            sheet.update("A1", data)

            # Envoi email
            default_extra_recipients = [
                "alexandre.audinot@latribu.fr",
                "jm.lelann@latribu.fr",
                "philippe.risso@firea.com"
            ]
            all_recipients = [DEFAULT_RECEIVER] + default_extra_recipients + [e.strip() for e in emails_supp.split(",") if e.strip() != ""]

            msg = EmailMessage()
            msg["Subject"] = f"📊 Rapport de valorisation des stocks au {date_import}"
            msg["From"] = SMTP_USER
            msg["To"] = ", ".join(all_recipients)
            msg.set_content(
                f"Bonjour,\n\nVoici le lien vers le tableau de bord dynamique de valorisation des stocks fournisseurs :\n👉 {LOOKER_URL}\n"
            )

            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.send_message(msg)

            st.success("📈 Google Sheets mis à jour et lien Looker envoyé par e-mail !")

        except Exception as e:
            st.error("❌ Erreur pendant la mise à jour ou l'envoi de l'e-mail.")
            st.exception(e)

    st.subheader("🗂️ Historique actuel")
    st.dataframe(historique_df, use_container_width=True)
else:
    st.info("Veuillez importer les fichiers de stock et de base produit pour démarrer.")
