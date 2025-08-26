import streamlit as st
import pandas as pd
import os
import json
import requests
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

# Email config depuis secrets
SMTP_SERVER = st.secrets["email"]["smtp_server"]
SMTP_PORT = st.secrets["email"]["smtp_port"]
SMTP_USER = st.secrets["email"]["smtp_user"]
SMTP_PASSWORD = st.secrets["email"]["smtp_password"]
DEFAULT_RECEIVER = st.secrets["email"]["receiver"]

# 🔥 Lecture dynamique du fichier Service Account depuis Google Drive
file_id = "12O9eFGFmwTu1n6kF4AIDIm0KXKMIgOvg"
url = f"https://drive.google.com/uc?id={file_id}"
response = requests.get(url)
response.raise_for_status()
gcp_service_account_info = json.loads(response.content)

# Authentification Google Sheets
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = service_account.Credentials.from_service_account_info(gcp_service_account_info, scopes=scopes)
client = gspread.authorize(creds)

# ---- Helpers: Google Sheets safe upsert (append without overwriting) ----
def _ensure_date_series(s):
    ds = pd.to_datetime(s, errors="coerce")
    return ds.dt.strftime("%Y-%m-%d")

def _gsheet_read_as_df(sheet_id, tab_name):
    try:
        ws = client.open_by_key(sheet_id).worksheet(tab_name)
    except Exception:
        sh = client.open_by_key(sheet_id)
        ws = sh.add_worksheet(title=tab_name, rows=2, cols=20)
    rows = ws.get_all_values()
    if not rows:
        return pd.DataFrame(), ws
    header = rows[0]
    data = rows[1:]
    if not data:
        return pd.DataFrame(columns=header), ws
    df = pd.DataFrame(data, columns=header)
    return df, ws

def _gsheet_upsert_dataframe(sheet_id, tab_name, df_new):
    # Read existing
    df_old, ws = _gsheet_read_as_df(sheet_id, tab_name)
    # Align columns
    if not df_old.empty:
        for c in df_new.columns:
            if c not in df_old.columns:
                df_old[c] = pd.NA
        for c in df_old.columns:
            if c not in df_new.columns:
                df_new[c] = pd.NA
        df_old = df_old[df_new.columns]
    # Normalize date
    if "date" in df_new.columns:
        df_new["date"] = _ensure_date_series(df_new["date"])
    if not df_old.empty and "date" in df_old.columns:
        df_old["date"] = _ensure_date_series(df_old["date"])
    # Concatenate
    df_all = pd.concat([df_old, df_new], ignore_index=True) if not df_old.empty else df_new.copy()
    # Deduplicate on composite key (date|organisationId|brand)
    key_cols = [c for c in ["date", "organisationId", "brand"] if c in df_all.columns]
    if key_cols:
        key = df_all[key_cols].astype(str).agg("|".join, axis=1)
        df_all = df_all.loc[~key.duplicated(keep="last")].copy()
    # Recompute est_derniere_date
    if "date" in df_all.columns:
        dmax = pd.to_datetime(df_all["date"], errors="coerce").max()
        df_all["est_derniere_date"] = (pd.to_datetime(df_all["date"], errors="coerce") == dmax)
    # Sort for readability
    sort_cols = [c for c in ["date", "organisationId", "brand"] if c in df_all.columns]
    if sort_cols:
        df_all = df_all.sort_values(sort_cols)
    # Write back
    ws.clear()
    values = [list(df_all.columns)] + df_all.astype(object).where(pd.notnull(df_all), "").values.tolist()
    ws.update("A1", values)
    return df_all
# ---- end helpers ----


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
            # Upsert: merge existing sheet content with historique_df
            df_all = _gsheet_upsert_dataframe(SPREADSHEET_ID, SHEET_NAME, historique_df)

            # Préparer et envoyer email
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
