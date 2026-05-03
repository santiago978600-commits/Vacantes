import os
import requests
import msal

# Configuración desde variables de entorno (GitHub Secrets)
CLIENT_ID = os.environ.get("MS_GRAPH_CLIENT_ID")
TENANT_ID = os.environ.get("MS_GRAPH_TENANT_ID")
CLIENT_SECRET = os.environ.get("MS_GRAPH_CLIENT_SECRET")
REFRESH_TOKEN = os.environ.get("MS_GRAPH_REFRESH_TOKEN")
SHAREPOINT_SITE_ID = os.environ.get("SHAREPOINT_SITE_ID")
FILE_PATH_IN_SHAREPOINT = os.environ.get("FILE_PATH_IN_SHAREPOINT", "General/Informe vacantes Auto.xlsx")
LOCAL_FILE_NAME = "Informe vacantes Auto.xlsx"

def download_file():
    if not all([CLIENT_ID, TENANT_ID, REFRESH_TOKEN]):
        print("Error: Faltan credenciales de Microsoft Graph en las variables de entorno.")
        return False

    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)

    print("Obteniendo token de acceso...")
    result = app.acquire_token_by_refresh_token(REFRESH_TOKEN, scopes=["https://graph.microsoft.com/.default"])

    if "access_token" not in result:
        print(f"Error al obtener token: {result.get('error_description')}")
        return False

    access_token = result['access_token']
    headers = {'Authorization': f'Bearer {access_token}'}

    # 1. Obtener el Drive ID del sitio de SharePoint
    print(f"Buscando el sitio de SharePoint: {SHAREPOINT_SITE_ID}")
    site_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_SITE_ID}/drive"
    
    response = requests.get(site_url, headers=headers)
    if response.status_code != 200:
        print(f"Error al obtener el drive del sitio: {response.text}")
        return False
    
    drive_id = response.json()['id']

    # 2. Descargar el archivo
    print(f"Descargando archivo: {FILE_PATH_IN_SHAREPOINT}")
    download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{FILE_PATH_IN_SHAREPOINT}:/content"
    
    response = requests.get(download_url, headers=headers)
    if response.status_code == 200:
        with open(LOCAL_FILE_NAME, 'wb') as f:
            f.write(response.content)
        print(f"Archivo descargado exitosamente como {LOCAL_FILE_NAME}")
        return True
    else:
        print(f"Error al descargar el archivo: {response.status_code} - {response.text}")
        return False

if __name__ == "__main__":
    download_file()
