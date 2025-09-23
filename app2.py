"""Pagina de automacao para a Integracao."""

from __future__ import annotations

from pathlib import Path
from io import BytesIO
import os
from typing import Optional

import pandas as pd
import requests
import streamlit as st
import plotly.express as px
from core.etl_integracao import process_integration_data, summarize_status


SHEET_URL_ENV = "INTEGRACAO_SHEET_URL"
SHEET_URL_SECRET_KEY = "integracao_sheet_url"



def _resolve_sheet_url() -> Optional[str]:
    """Resolve a URL da planilha a partir de secrets ou variavel de ambiente."""
    secret_val: Optional[str] = None
    try:
        secrets_obj = getattr(st, "secrets", None)
        if secrets_obj:
            if isinstance(secrets_obj, dict):
                secret_val = secrets_obj.get(SHEET_URL_SECRET_KEY)
            else:
                try:
                    secret_val = secrets_obj[SHEET_URL_SECRET_KEY]  # type: ignore[index]
                except Exception:
                    secret_val = getattr(secrets_obj, SHEET_URL_SECRET_KEY, None)
    except Exception:
        secret_val = None
    if secret_val:
        return str(secret_val)

    env_val = os.getenv(SHEET_URL_ENV)
    if env_val:
        return env_val

    try:
        secrets_path = Path('.streamlit/secrets.toml')
        if secrets_path.exists():
            import tomllib
            data = tomllib.loads(secrets_path.read_text(encoding='utf-8'))
            sheet_val = data.get(SHEET_URL_SECRET_KEY)
            if sheet_val:
                return str(sheet_val)
    except Exception:
        pass

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
        <h1 style='margin: 6px 0 12px 0; font-size: 36px;'>Integração</h1>
        <h3 style='margin: 0 0 24px 0; font-size: 18px; color: grey;'>
        Suba o Excel (.xlsb, .xlsx) e acompanhe KPIs e detalhamento por status
        </h3>
        """,
        unsafe_allow_html=True,
    )
    st.caption("Fonte: Google Sheets (CONTROLE_CLARO_RAN_INTEGRAÇÃO)")

    with st.spinner("Carregando planilha online..."):
        df = _fetch_sheet()

    if df is None:
        st.stop()

    try:
        df = process_integration_data(df)
    except ValueError as e:
        st.error(f"Erro ao processar os dados: {e}")
        st.stop()

    # Remover números das datas
    date_columns = ["Integration date"]
    for col in date_columns:
        if col in df.columns:
            df[col] = df[col].dt.date

    # Adjusting the graph colors and x-axis label
    # Filter out rows with empty 'Site Name'
    df = df[df["Site Name"].notna()]

    # Update the message to count sites instead of rows
    site_count = df["Site Name"].nunique() if "Site Name" in df.columns else 0
    st.success(f"Planilha carregada com {site_count:,} sites identificados.")

    # Filtros na página principal
    st.markdown(
        """
        <h2 style='margin: 12px 0; font-size: 24px;'>Filtros</h2>
        """,
        unsafe_allow_html=True,
    )
    status_filter = st.multiselect(
        "Filtrar por Status 4G:", options=df["4G Status"].unique(), default=[]
    )

    if status_filter:
        df = df[df["4G Status"].isin(status_filter)]

    # Status de Integração
    st.markdown(
        """
        <h2 style='margin: 12px 0; font-size: 24px;'>Status de Integração</h2>
        """,
        unsafe_allow_html=True,
    )

    # Alternar entre gráficos
    graph_option = st.radio(
        "Escolha o gráfico:",
        options=["Status por Tecnologia", "General Status"],
        index=0
    )

    if graph_option == "Status por Tecnologia":
        # Gráfico de Status por Tecnologia com 9 colunas específicas
        integration_columns = [
            "4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 4G", "OT 2G"
        ]

        status_counts = pd.concat([
            df[col].value_counts().rename_axis("Status").reset_index(name="Count").assign(Type=col)
            for col in integration_columns if col in df.columns
        ])

        # Filter and map statuses to the first set
        status_counts = status_counts[status_counts['Status'].isin(["Finished", "Need update", "Waiting", "Pending"])]
