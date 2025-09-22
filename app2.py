"""Pagina de automacao para a Integracao."""

from __future__ import annotations

from io import BytesIO
import os
from typing import Optional

import pandas as pd
import requests
import streamlit as st


SHEET_URL_ENV = "INTEGRACAO_SHEET_URL"
SHEET_URL_SECRET_KEY = "integracao_sheet_url"


def _resolve_sheet_url() -> Optional[str]:
    """Resolve a URL da planilha a partir de secrets ou variavel de ambiente."""
    try:
        secret_val = st.secrets.get(SHEET_URL_SECRET_KEY)  # type: ignore[attr-defined]
    except Exception:
        secret_val = None
    if secret_val:
        return str(secret_val)
    env_val = os.getenv(SHEET_URL_ENV)
    if env_val:
        return env_val
    return None


def _fetch_sheet(timeout: int = 30) -> Optional[pd.DataFrame]:
    """Baixa a planilha de Integracao do Google Sheets como DataFrame."""
    sheet_url = _resolve_sheet_url()
    if not sheet_url:
        st.error(
            "URL da planilha nao configurada. Defina INTEGRACAO_SHEET_URL ou st.secrets['integracao_sheet_url']."
        )
        return None
    try:
        resp = requests.get(sheet_url, timeout=timeout)
        resp.raise_for_status()
    except Exception as exc:  # pragma: no cover - apenas logging
        st.error(f"Falha ao baixar a planilha de integracao: {exc}")
        return None
    try:
        return pd.read_excel(BytesIO(resp.content))
    except Exception as exc:  # pragma: no cover
        st.error(f"Falha ao ler a planilha baixada: {exc}")
        return None


def page_integracao() -> None:
    """Renderiza a pagina principal da automacao de Integracao."""
    st.markdown(
        """
        <h2 style='margin: 6px 0 12px 0; font-size: 28px;'>Integra\u00e7\u00e3o</h2>
        """,
        unsafe_allow_html=True,
    )
    st.caption("Fonte: Google Sheets (Integra\u00e7\u00e3o)")

    with st.spinner("Carregando planilha online..."):
        df = _fetch_sheet()

    if df is None:
        st.stop()

    st.success(f"Planilha carregada com {len(df):,} linhas.")
    st.dataframe(df.head(50), use_container_width=True)
