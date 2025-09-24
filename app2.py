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


def _fetch_sheet_local(file_path: str) -> Optional[pd.DataFrame]:
    """Lê a planilha de integração local como DataFrame."""
    try:
        return pd.read_excel(file_path)
    except Exception as exc:
        st.error(f"Falha ao ler a planilha local: {exc}")
        return None


def page_integracao() -> None:
    """Renderiza a página principal da automação de Integração."""
    st.markdown(
        """
        <h1 style='margin: 6px 0 12px 0; font-size: 36px;'>Integração</h1>
        <h3 style='margin: 0 0 24px 0; font-size: 18px; color: grey;'>
        Suba o Excel (.xlsb, .xlsx) e acompanhe KPIs e detalhamento por status
        </h3>
        """,
        unsafe_allow_html=True,
    )
    st.caption("Fonte: Planilha local ou Google Sheets")

    with st.spinner("Carregando planilha..."):
        df = None
        if os.getenv("LOCAL_TEST", "false").lower() == "true":
            # Leitura local para testes
            df = _fetch_sheet_local("C:\\Users\\Vitor Sena\\Desktop\\Automacoes\\Rollout_sites\\CONTROLE_CLARO_RAN_INTEGRAÇÃO.xlsx")
        else:
            # Tentar leitura online
            df = _fetch_sheet()
            if df is None:
                # Fallback para leitura local caso a URL não esteja configurada
                df = _fetch_sheet_local("C:\\Users\\Vitor Sena\\Desktop\\Automacoes\\Rollout_sites\\CONTROLE_CLARO_RAN_INTEGRAÇÃO.xlsx")

    if df is None:
        st.error("Falha ao carregar a planilha. Verifique se o arquivo local ou a URL estão configurados corretamente.")
        st.stop()

    try:
        df = process_integration_data(df)
    except ValueError as e:
        st.error(f"Erro ao processar os dados: {e}")
        st.stop()

    # Remover horários: converter qualquer coluna datetime para date (aplica às tabelas)
    for col in df.select_dtypes(include=["datetime64[ns]", "datetime64"]).columns:
        df[col] = df[col].dt.date

    # Ajustando as cores dos gráficos e rótulo do eixo x
    # Filtrar linhas com 'Site Name' vazio
    df = df[df["Site Name"].notna()]

    # Filtrar apenas sites com valores válidos em 'General Status'
    valid_general_status = {"on going", "finished"}
    df = df[df["General Status"].str.lower().isin(valid_general_status)]

    # Atualizar a mensagem para contar apenas os sites válidos
    site_count = df["Site Name"].nunique() if "Site Name" in df.columns else 0
    st.success(f"Planilha carregada com {site_count:,} sites válidos identificados.")

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
        options=["Integração Concluído x Faltando", "General Status"],
        index=0
    )

    if graph_option == "Integração Concluído x Faltando":
        # Gráfico de Integração Concluído x Faltando
        integration_columns = [
            "4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status"
        ]

        status_counts = pd.concat([
            df[col].value_counts().rename_axis("Status").reset_index(name="Count").assign(Type=col)
            for col in integration_columns if col in df.columns
        ])

        # Map raw status values into three groups: Concluido, Faltando
        def _map_to_group(s):
            if pd.isna(s):
                return None  # Ignorar valores nulos
            v = str(s).strip().lower()
            # Concluido
            if v in {"finished"}:
                return "Concluido"
            # Faltando
            if v in {"pending", "kpi rejected", "pendência", "pendência kpi", "upload to iw", "waiting approval", "waiting", "aguardando aprovação"}:
                return "Faltando"
            # Unknown ou outros valores não devem ser contados
            return None

        status_counts["Status"] = status_counts["Status"].map(_map_to_group)

        # Excluir valores não mapeados (None)
        status_counts = status_counts[status_counts["Status"].notna()]

        # Aggregate counts after mapping
        status_counts = (
            status_counts.groupby(["Type", "Status"])["Count"].sum().reset_index()
        )

        # Garantir ordem das categorias no eixo x conforme integration_columns
        status_counts["Type"] = pd.Categorical(status_counts["Type"], categories=integration_columns, ordered=True)

        fig = px.bar(
            status_counts,
            x="Type",
            y="Count",
            color="Status",
            text="Count",
            title="Resumo do Status por Categoria",
            labels={"Type": "Categoria", "Count": "Quantidade", "Status": "Status"},
            category_orders={"Type": integration_columns, "Status": ["Concluido", "Faltando"]},
            color_discrete_map={
                "Concluido": "#1f77b4",  # Azul similar ao rollout
                "Faltando": "#ff7f0e",   # Laranja similar ao rollout
            }
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    elif graph_option == "General Status":
        # Gráfico de General Status
        general_status_counts = df["General Status"].value_counts().rename_axis("Status").reset_index(name="Count")

        fig = px.bar(
            general_status_counts,
            x="Status",
            y="Count",
            text="Count",
            title="Resumo do General Status",
            labels={"Status": "Status", "Count": "Quantidade"},
            color="Status",
            color_discrete_map={
                "Finished": "#28a745",  # Verde vibrante
                "On going": "#007bff",  # Azul vibrante
                "Waiting": "#ffc107"  # Amarelo vibrante
            }
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    # Tabela de resumo: selecionar apenas colunas existentes
    desired_columns = [
        "Site Name", "Integration date", "MOS", "General Status", "4G Status", "2G Status",
        "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status"
    ]
    existing_cols = [c for c in desired_columns if c in df.columns]
    status_summary = df[existing_cols]

    # ---- Converter status para a PRIMEIRA tabela (apenas) para Concluido / Faltando ----
    # Mapear valores para Concluido / Faltando (aplica-se somente na tabela de resumo)
    def map_to_two(s):
        if pd.isna(s):
            return None
        v = str(s).strip().lower()
        if v == "finished":
            return "Concluido"
        if v in {"pending", "kpi rejected", "pendência", "pendência kpi", "pendencia", "upload to iw", "waiting approval", "waiting", "aguardando aprovação"}:
            return "Faltando"
        return None

    # Aplicar a transformação somente nas colunas de teste/status presentes
    status_cols = [c for c in ["4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status"] if c in status_summary.columns]
    summary_for_display = status_summary.copy()
    for c in status_cols:
        summary_for_display[c] = summary_for_display[c].map(map_to_two)

    # Estilização da tabela de resumo: Concluido (azul), Faltando (laranja)
    def style_two(val):
        if pd.isna(val):
            return ""
        v = str(val).strip().lower()
        if v == "concluido":
            return "color: #1f77b4; font-weight: 600"
        if v == "faltando":
            return "color: #ff7f0e; font-weight: 600"
        return ""

    if not summary_for_display.empty:
        styled_summary = summary_for_display.style.applymap(style_two, subset=status_cols)
        st.dataframe(styled_summary, use_container_width=True)
    else:
        st.write("Nenhum registro para exibir na tabela de resumo.")

    # ---- Tabela Fiel: manter valores originais (fiel), mas aplicar cores por valor ----
    # Cores por valor usadas na planilha (aproximação):
    faithful_colors = {
        "finished": "background-color: #d4edda; color: #155724",  # greenish
        "pending": "background-color: #fff3cd; color: #856404",   # orange/yellow
        "kpi rejected": "background-color: #f8d7da; color: #721c24",  # red-ish
        "waiting approval": "background-color: #cfe2ff; color: #084298",  # blue-ish
        "upload to iw": "background-color: #d1ecf1; color: #0c5460",  # light-blue
        "on going": "background-color: #cfe2ff; color: #084298",  # blue-ish (On going)
        "ongoing": "background-color: #cfe2ff; color: #084298"  # variant
    }

    def style_faithful(val):
        if pd.isna(val):
            return ""
        v = str(val).strip().lower()
        return faithful_colors.get(v, "")

    st.markdown(
        """
        <h2 style='margin: 12px 0; font-size: 24px;'>Tabela Fiel</h2>
        """,
        unsafe_allow_html=True,
    )
    with st.expander("Tabela Fiel", expanded=False):
        # aplicar largura do Comment e coloração nas colunas de status existentes
        fiel_status_cols = [c for c in ["General Status", "4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status"] if c in df.columns]
        base_style = df.style.set_properties(subset=["Comment"] if "Comment" in df.columns else [], **{"width": "300px"})
        if fiel_status_cols:
            base_style = base_style.applymap(style_faithful, subset=fiel_status_cols)
        st.dataframe(base_style, use_container_width=True)
