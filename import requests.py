import requests
import pandas as pd
from msal import ConfidentialClientApplication

# 🔐 App-configuratie
client_id = "5ff111e9-3264-4cbd-8d54-d9dffd5ae1e3"
client_secret = "G7l8Q~2.XgJF2O8OpKhxaTqaP3aPq9FtSM7HIcaf"
tenant_id = "483c6905-9f7e-4ed4-b4db-ab08ec904d4e"
site_name = "demotestsofie"  # Enkel de sitenaam, niet het volledige domein
site_domain = "argosvzw.sharepoint.com"
lijst_naam = "Registratietaken"  # Naam van de SharePoint-lijst

# 🌐 AUTHENTICATIE
authority = f"https://login.microsoftonline.com/{tenant_id}"
scope = ["https://graph.microsoft.com/.default"]

app = ConfidentialClientApplication(client_id, authority=authority, client_credential=client_secret)
token_result = app.acquire_token_for_client(scopes=scope)

if "access_token" not in token_result:
    raise Exception("❌ Authenticatie mislukt:", token_result.get("error_description"))

headers = {
    "Authorization": f"Bearer {token_result['access_token']}",
    "Accept": "application/json"
}

# 🔍 1. Haal site ID op
site_url = f"https://graph.microsoft.com/v1.0/sites/{site_domain}:/sites/{site_name}"
site_res = requests.get(site_url, headers=headers).json()
site_id = site_res.get("id")

# 🔍 2. Haal lijst ID op
list_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{lijst_naam}"
list_res = requests.get(list_url, headers=headers).json()
list_id = list_res.get("id")

# 📥 3. Haal lijstitems op
items_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_id}/items?expand=fields&top=999"
items_res = requests.get(items_url, headers=headers).json()

records = [item['fields'] for item in items_res.get('value', [])]
df = pd.DataFrame(records)

# ✅ Laat eerste 5 rijen zien
print(df.head())
