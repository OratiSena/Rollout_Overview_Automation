import requests
from msal import ConfidentialClientApplication
from io import BytesIO

class GraphError(Exception):
    pass

def download_excel_bytes(drive_id, item_id, share_url):
    from core.config import Settings

    # Autenticação MSAL
    app = ConfidentialClientApplication(
        Settings.CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{Settings.TENANT_ID}",
        client_credential=Settings.CLIENT_SECRET,
    )
    token_result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in token_result:
        raise GraphError("Erro ao obter token: " + str(token_result.get("error_description")))

    headers = {"Authorization": "Bearer " + token_result["access_token"]}

    # Monta URL do arquivo
    if drive_id and item_id:
        url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    elif share_url:
        import base64
        share_id = base64.urlsafe_b64encode(share_url.encode()).decode().rstrip("=")
        url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content"
    else:
        raise GraphError("drive_id/item_id ou share_url devem ser informados.")

    resp = requests.get(url, headers=headers)
    if resp.status_code != 200:
        raise GraphError(f"Erro ao baixar arquivo: {resp.status_code} {resp.text}")

    return BytesIO(resp.content)