import gspread
from google.oauth2.service_account import Credentials
import traceback

SERVICE_ACCOUNT_FILE = "valostock-gsheet-a3fb9b0b374a.json"
SPREADSHEET_ID = "1lOtH16m_xs1-EzQ7D_tp8wZz3fZu2eTbLQFU099MSNw"

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
client = gspread.authorize(creds)

try:
    sheet = client.open_by_key(SPREADSHEET_ID)
    print("✅ Connexion réussie au Google Sheets !")
    print("Onglets disponibles :", [ws.title for ws in sheet.worksheets()])
except Exception as e:
    print("❌ Erreur :", e)
    traceback.print_exc()
