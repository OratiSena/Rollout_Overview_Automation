import os

# suporta rodar com/sem Streamlit
def _get_secret(path, default=""):
    try:
        import streamlit as st
        return st.secrets[path.split(".")[0]][path.split(".")[1]]
    except Exception:
        # fallback para variáveis de ambiente
        key = path.upper().replace(".", "_")
        return os.getenv(key, default)

class Settings:
    TENANT_ID     = _get_secret("graph.tenant_id")
    CLIENT_ID     = _get_secret("graph.client_id")
    CLIENT_SECRET = _get_secret("graph.client_secret")
    DRIVE_ID      = _get_secret("graph.drive_id")
    ITEM_ID       = _get_secret("graph.item_id")
    SHARE_URL     = _get_secret("graph.share_url")
    SHEET_NAME    = _get_secret("data.sheet_name", "")
