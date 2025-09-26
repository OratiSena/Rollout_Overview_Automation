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

    # OBS: 'Integration date' e 'MOS' podem não existir mais na planilha online;
    # não forçamos conversão aqui — os filtros e tabelas lidam com colunas opcionais.

    # Ajustando as cores dos gráficos e rótulo do eixo x
    # Filtrar linhas com 'Site Name' vazio
    df = df[df["Site Name"].notna()]

    # Mostrar quantos sites a planilha tem (antes de aplicar filtros)
    site_count = df["Site Name"].nunique() if "Site Name" in df.columns else 0
    # Criar o placeholder antes do cabeçalho para que tanto a mensagem inicial
    # quanto a atualizada (após filtros) apareçam acima do título 'Filtros'
    msg_ph = st.empty()
    msg_ph.success(f"Planilha carregada com {site_count:,} sites válidos identificados.")

    # Cabeçalho de filtros — os widgets reais ficam dentro do expander abaixo
    if st.session_state.get("int_show_filters", True):
        st.markdown(
            """
            <h2 style='margin: 12px 0; font-size: 24px;'>Filtros</h2>
            """,
            unsafe_allow_html=True,
        )

        # --- FILTROS (agora em expander, similar ao rollout) ---
        with st.expander("Filtros", expanded=True):
            # Preparar apenas as colunas que o Reset precisa (sem criar widgets ainda)
            status_columns = [c for c in ["4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV"] if c in df.columns]
            ot_columns = [c for c in ["OT 2G", "OT 4G"] if c in df.columns]

            # Escolha do gráfico no topo do expander (como antes) e Reset no canto direito
            # Forçar layout horizontal das opções do radio via CSS local (seletor mais amplo)
            st.markdown(
                """
                <style>
                /* Force radio options inline by making labels inline-flex */
                div[role="radiogroup"] { display: flex !important; flex-direction: row !important; gap: 10px !important; align-items: center !important; }
                div[role="radiogroup"] label { display: inline-flex !important; align-items: center !important; margin-right: 18px !important; }
                /* fallback: target common streamlit radio container classes */
                .stRadio div[role="radiogroup"] { display: flex !important; }
                </style>
                """,
                unsafe_allow_html=True,
            )
            # aumentar a proporcao para empurrar o botão Reset mais para a direita
            top_cols = st.columns([8, 1])
            with top_cols[0]:
                # Make 'Integração Concluído x Faltando' the first/default option,
                # then 'Sites por status (status atual)', then 'General Status'.
                graph_option = st.radio(
                    "Escolha o gráfico:",
                    options=["Integração Concluído x Faltando", "Sites por status (status atual)", "General Status"],
                    index=0,
                    key="f_graph_option",
                )
                # Global Status radio: Geral, Concluidos, Faltando, Unknow
                status_choice = st.radio("Status", options=["Geral", "Concluidos", "Faltando", "Unknow"], index=0, key="f_status_choice")
            with top_cols[1]:
                if st.button("Resetar", key="f_reset_filters"):
                    keys_to_clear = ["f_gen_status", "f_txt_search", "f_region", "f_arq", "f_graph_option"]
                    # include the new status_atual filter key
                    keys_to_clear.append("f_status_atual")
                    for col in status_columns:
                        keys_to_clear.append(f"map_{col}")
                    for col in ot_columns:
                        keys_to_clear.append(f"f_ot_{col}")
                    for k in keys_to_clear:
                        if k in st.session_state:
                            del st.session_state[k]
                    # Restore explicit defaults (ensure widgets show their initial values)
                    st.session_state["f_gen_status"] = []
                    st.session_state["f_txt_search"] = ""
                    st.session_state["f_region"] = []
                    st.session_state["f_arq"] = []
                    st.session_state["f_graph_option"] = "Sites por status (status atual)"
                    # restore status_choice default
                    st.session_state["f_status_choice"] = "Geral"
                    # For status mapping and OT filters, set empty selections
                    for col in status_columns:
                        st.session_state.setdefault(f"map_{col}", [])
                    for col in ot_columns:
                        st.session_state.setdefault(f"f_ot_{col}", [])

                    # clear the status_atual filter default
                    st.session_state.setdefault("f_status_atual", [])

                    # Try to rerun the app; Streamlit API differs across versions.
                    try:
                        st.experimental_rerun()
                    except Exception:
                        try:
                            st.rerun()
                        except Exception:
                            st.stop()


            # (Integration date / MOS removed from filters - sheet no longer provides them)
            int_start = int_end = None
            mos_start = mos_end = None

            # Preparar valores e colunas utilizáveis (após os controles do topo)
            search_columns = [c for c in ["Site Name", "General Status", "4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "Region", "Comment", "ARQ Number", "OT Status", "OT 2G", "OT 4G"] if c in df.columns]
            # New filter: Selecionar por status_atual (computed)
            # We compute the available options from the raw dataframe (before applying other filters)
            # and display them in a fixed order with numeric prefixes.
            def _compute_status_atual_series(src_df: pd.DataFrame) -> pd.Series:
                steps = ["4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status"]

                def _row_status(row):
                    gen = row.get("General Status") if "General Status" in row.index else None
                    if pd.isna(gen) or str(gen).strip() == "":
                        return None

                    any_nonnull = False
                    for col in steps:
                        if col not in row.index:
                            continue
                        val = row[col]
                        if pd.isna(val) or str(val).strip() == "":
                            continue
                        any_nonnull = True
                        v = str(val).strip().lower()
                        if v == "finished":
                            continue
                        return col

                    if not any_nonnull:
                        return "Sem informação"
                    return "Concluído"

                return src_df.apply(_row_status, axis=1)

            # Build ordered display options (numbered) and a mapping to actual values
            steps_order = ["4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status"]
            try:
                temp_status = _compute_status_atual_series(df)
                unique_vals = [v for v in pd.unique(temp_status) if pd.notna(v)]
            except Exception:
                unique_vals = []

            display_map = {}
            ordered_display = []
            # Number and include only values that exist in the data
            for i, step in enumerate(steps_order, start=1):
                # show the step in the selector if the column exists in the sheet
                if step in df.columns:
                    label = f"{i} - {step}"
                    display_map[label] = step
                    ordered_display.append(label)

            # Include Sem informação and Concluído at the end if present
            if "Sem informação" in unique_vals:
                display_map["Sem informação"] = "Sem informação"
                ordered_display.append("Sem informação")
            if "Concluído" in unique_vals:
                display_map["Concluído"] = "Concluído"
                ordered_display.append("Concluído")

            sel_status_atual = st.multiselect("Selecione o Status:", options=ordered_display, default=[], key="f_status_atual")
            # Top row: General Status first, then search, then region, then ARQ
            r1c1, r1c2, r1c3, r1c4 = st.columns([2,3,2,2])
            with r1c1:
                # Do not expose 'Unknown' option here; Unknown is handled via the Status radio
                gen_opts = ["Finished", "On going"]
                sel_general = st.multiselect("General Status:", options=gen_opts, default=[], help="Filtra por General Status", key="f_gen_status")
            with r1c2:
                txt_search = st.text_input("Pesquisar (Site, status, Region, Comment, ARQ):", key="f_txt_search")
            with r1c3:
                region_opts = sorted(df["Region"].dropna().unique().tolist()) if "Region" in df.columns else []
                sel_region = st.multiselect("Region:", options=region_opts, default=[], key="f_region")
            with r1c4:
                arq_opts = sorted(df["ARQ Number"].dropna().unique().tolist()) if "ARQ Number" in df.columns else []
                sel_arq = st.multiselect("ARQ Number:", options=arq_opts, default=[], key="f_arq")

            # Third area: (status radio moved above)
            if status_columns:
                pass

            # OT filters on a single row
            if ot_columns:
                ot_cols = st.columns(len(ot_columns))
                sel_ot = {}
                for i, col in enumerate(ot_columns):
                    # Build options excluding Unknown-like values
                    raw_opts = sorted(df[col].dropna().unique().tolist())
                    opts = [v for v in raw_opts if str(v).strip().lower() not in {"unknown", "", "nan"}]
                    # no default selection so OT filters are not applied implicitly
                    key_name = f"f_ot_{col}"
                    sel = ot_cols[i].multiselect(f"{col}:", options=opts, default=[], key=key_name)
                    sel_ot[col] = sel

            # end of filters (no divider)

    # Aplicar os filtros sobre uma cópia (widgets existiram dentro do expander)
    df_filtered = df.copy()

    # Texto livre
    if txt_search:
        txt = txt_search.strip().lower()
        mask = pd.Series(False, index=df_filtered.index)
        for c in search_columns:
            mask = mask | df_filtered[c].astype(str).str.lower().str.contains(txt, na=False)
        df_filtered = df_filtered[mask]

    # Region
    if sel_region:
        df_filtered = df_filtered[df_filtered["Region"].isin(sel_region)]

    # Integration date / MOS filtering removed (columns may no longer exist in sheet)

    # General Status
    if sel_general:
        def general_match(v):
            if pd.isna(v):
                return "Unknown"
            s = str(v).strip()
            if s.lower() == "finished":
                return "Finished"
            if s.lower() in {"on going", "ongoing"}:
                return "On going"
            return "Unknown"
        df_filtered = df_filtered[df_filtered["General Status"].apply(general_match).isin(sel_general)]

    # ARQ
    if sel_arq:
        df_filtered = df_filtered[df_filtered["ARQ Number"].isin(sel_arq)]

    # Status map filters (Concluido/Faltando) for each status column
    def map_to_two_local(s):
        if pd.isna(s):
            return None
        v = str(s).strip().lower()
        if v == "finished":
            return "Concluido"
        if v in {"pending", "kpi rejected", "pendência", "pendência kpi", "upload to iw", "waiting approval", "waiting", "aguardando aprovação"}:
            return "Faltando"
        return None

    # Apply global Status radio filter
    status_choice = st.session_state.get("f_status_choice", "Geral")
    # Interpret General Status 'in integration' as Finished/On going/Unknown mapping already used elsewhere
    def _is_integration_row(row):
        v = row.get("General Status") if "General Status" in row.index else None
        if pd.isna(v):
            return False
        s = str(v).strip().lower()
        return s != "unknown" and s != ""

    if status_choice == "Concluidos":
        # Use OT Status finished as the source of truth for Concluidos (OT Status columns)
        ot_mask = pd.Series(False, index=df_filtered.index)
        for col in ot_columns:
            if col in df_filtered.columns:
                ot_mask = ot_mask | (df_filtered[col].astype(str).str.strip().str.lower() == "finished")
        df_filtered = df_filtered[ot_mask]
    elif status_choice == "Faltando":
        # Faltando: any status column mapped to Faltando
        mask = pd.Series(False, index=df_filtered.index)
        for col in status_columns:
            mapped = df_filtered[col].apply(map_to_two_local)
            mask = mask | mapped.isin(["Faltando"])
        df_filtered = df_filtered[mask]
    elif status_choice == "Unknow":
        # Unknow: rows not in integration (General Status empty or Unknown)
        mask = pd.Series(False, index=df_filtered.index)
        if "General Status" in df_filtered.columns:
            mask = df_filtered["General Status"].astype(str).str.strip().str.lower().isin(["", "nan", "unknown"])
        else:
            # if General Status absent, nothing is Unknown
            mask = pd.Series(False, index=df_filtered.index)
        df_filtered = df_filtered[mask]
    else:
        # Geral: show only rows that are in integration (exclude Unknown/empty General Status)
        mask = pd.Series(False, index=df_filtered.index)
        if "General Status" in df_filtered.columns:
            mask = df_filtered["General Status"].astype(str).str.strip().str.lower().apply(lambda s: s not in {"", "nan", "unknown"})
        df_filtered = df_filtered[mask]

    # OT raw filters
    for col, sel in (sel_ot.items() if ot_columns else []):
        if sel:
            df_filtered = df_filtered[df_filtered[col].isin(sel)]

    # Usar df_filtered em vez de df daqui para frente (gráfico e tabelas)
    df = df_filtered

    # --- Compute status_atual for the filtered dataframe and allow filtering by it ---
    def _compute_status_atual_for_df(src_df: pd.DataFrame) -> pd.Series:
        steps = ["4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status"]

        def _row_status(row):
            gen = row.get("General Status") if "General Status" in row.index else None
            if pd.isna(gen) or str(gen).strip() == "":
                return None

            any_nonnull = False
            for col in steps:
                if col not in row.index:
                    continue
                val = row[col]
                if pd.isna(val) or str(val).strip() == "":
                    continue
                any_nonnull = True
                v = str(val).strip().lower()
                if v == "finished":
                    continue
                return col

            if not any_nonnull:
                return "Sem informação"
            return "Concluído"

        return src_df.apply(_row_status, axis=1)

    try:
        # Prefer OT Status finished to mark Concluído explicitly: when OT Status == Finished -> status_atual = 'Concluído'
        df_temp = df.copy()
        if "OT Status" in df_temp.columns:
            ot_finished = df_temp["OT Status"].astype(str).str.strip().str.lower() == "finished"
            df_temp.loc[ot_finished, "status_atual"] = "Concluído"
            # For the rest, compute normally
            remaining = df_temp["status_atual"].isna()
            if remaining.any():
                df_temp.loc[remaining, "status_atual"] = _compute_status_atual_for_df(df_temp[remaining])
        else:
            df_temp["status_atual"] = _compute_status_atual_for_df(df_temp)
        df["status_atual"] = df_temp["status_atual"]
    except Exception:
        df["status_atual"] = None

    # Map display labels back to actual values and apply the new status_atual filter if present
    if sel_status_atual:
        mapped = []
        for sel in sel_status_atual:
            mapped.append(display_map.get(sel, sel))
        df = df[df["status_atual"].isin(mapped)]

    # Atualizar a mensagem placeholder para contar apenas os sites filtrados
    site_count = df["Site Name"].nunique() if "Site Name" in df.columns else 0
    msg_ph.success(f"Registros mostrados: {site_count:,} sites após filtros aplicados.")

    # Status de Integração
    if st.session_state.get("int_show_status", True):
        st.markdown(
            """
            <h2 style='margin: 12px 0; font-size: 24px;'>Status de Integração</h2>
            """,
            unsafe_allow_html=True,
        )

        # Utiliza a escolha feita nos filtros (se presente)
        graph_option = st.session_state.get("f_graph_option", "Integração Concluído x Faltando")

        # compute number of sites considered 'in integration' (General Status not Unknown/empty)
        in_integration_count = 0
        if "General Status" in df.columns:
            def _is_integration(v):
                if pd.isna(v):
                    return False
                s = str(v).strip().lower()
                return s != "unknown" and s != ""
            in_integration_count = int(df["General Status"].apply(_is_integration).sum())

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
                title=f"Resumo do Status por Categoria | {in_integration_count:,} sites em Integração",
                labels={"Type": "Status", "Count": "Quantidade", "Status": "Status"},
                category_orders={"Type": integration_columns, "Status": ["Concluido", "Faltando"]},
                color_discrete_map={
                    "Concluido": "#1f77b4",  # Azul similar ao rollout
                    "Faltando": "#ff7f0e",   # Laranja similar ao rollout
                }
            )
            # Let Plotly decide best text position (inside/outside) and give room at the top
            fig.update_traces(textposition="auto")
            fig.update_layout(margin=dict(t=80), legend_title_text="Status", uniformtext_minsize=8)
            st.plotly_chart(fig, use_container_width=True)

        elif graph_option == "Sites por status (status atual)":
            # Count sites by computed status_atual (exclude None) and enforce order
            sac = df["status_atual"].value_counts(dropna=True).rename_axis("Status").reset_index(name="Count")
            # Desired ordering: follow steps order, then Sem informação, then Concluído last
            desired_order = ["4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status", "Sem informação", "Concluído"]
            present = [s for s in desired_order if s in sac["Status"].values]
            if present:
                sac["Status"] = pd.Categorical(sac["Status"], categories=present, ordered=True)
                sac = sac.sort_values("Status")
            if sac.empty:
                st.write("Nenhum dado disponível para 'Sites por status (status atual)'.")
            else:
                fig2 = px.bar(
                    sac,
                    x="Status",
                    y="Count",
                    text="Count",
                    title=f"Sites por status (status atual) | {in_integration_count:,} sites em Integração",
                    labels={"Status": "Status atual", "Count": "Quantidade"},
                    color="Status",
                    category_orders={"Status": present}
                )
                fig2.update_traces(textposition="outside")
                st.plotly_chart(fig2, use_container_width=True)

        elif graph_option == "General Status":
            # Gráfico de General Status
            # If the Status radio is set to 'Unknow', show only the Unknown bar and a different title
            if st.session_state.get("f_status_choice", "Geral") == "Unknow":
                # Count unknown/not-filled rows
                if "General Status" in df.columns:
                    mask = df["General Status"].astype(str).str.strip().str.lower().isin(["", "nan", "unknown"])
                    unknown_count = int(mask.sum())
                else:
                    unknown_count = 0
                general_status_counts = pd.DataFrame({"Status": ["Unknown"], "Count": [unknown_count]})
                title = f"Resumo do General Status | {unknown_count:,} Sites não integrados/não preenchidos"
                fig = px.bar(
                    general_status_counts,
                    x="Status",
                    y="Count",
                    text="Count",
                    title=title,
                    labels={"Status": "Status", "Count": "Quantidade"},
                    color="Status",
                )
            else:
                general_status_counts = df["General Status"].value_counts().rename_axis("Status").reset_index(name="Count")

                fig = px.bar(
                    general_status_counts,
                    x="Status",
                    y="Count",
                    text="Count",
                    title=f"Resumo do General Status | {in_integration_count:,} sites em Integração",
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

    # Tabela de resumo: selecionar apenas colunas existentes (pertence à seção 'Status de Integração')
    if st.session_state.get("int_show_status", True):
        desired_columns = [
            "Site Name", "General Status", "4G Status", "2G Status",
            "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status"
        ]
        # Ensure status_atual is positioned immediately to the right of General Status
        existing_cols = [c for c in desired_columns if c in df.columns]
        cols = []
        for c in existing_cols:
            cols.append(c)
            if c == "General Status":
                # insert status_atual right after General Status (if it exists)
                if "status_atual" in df.columns:
                    cols.append("status_atual")

        # Append Region and ARQ Number as the last columns if present
        for endc in ["Region", "ARQ Number"]:
            if endc in df.columns and endc not in cols:
                cols.append(endc)

        status_summary = df[cols]

    # ---- Converter status para a PRIMEIRA tabela (apenas) para Concluido / Faltando ----
    # Esta lógica pertence exclusivamente à seção 'Status de Integração' e
    # deve ser executada somente quando essa seção estiver visível.
    if st.session_state.get("int_show_status", True):
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
            # ensure we don't try to map our injected helper column
            if c == "status_atual":
                continue
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

    # ---- Tabela Consolidada: manter valores originais (fiel), mas aplicar cores por valor ----
    if st.session_state.get("int_show_consolidada", True):
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
            <h2 style='margin: 12px 0; font-size: 24px;'>Tabela Consolidada</h2>
            """,
            unsafe_allow_html=True,
        )
        with st.expander("Tabela Consolidada", expanded=False):
            # aplicar largura do Comment e coloração nas colunas de status existentes
            fiel_status_cols = [c for c in ["General Status", "4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 2G", "OT 4G", "OT Status"] if c in df.columns]
            base_style = df.style.set_properties(subset=["Comment"] if "Comment" in df.columns else [], **{"width": "300px"})
            if fiel_status_cols:
                base_style = base_style.applymap(style_faithful, subset=fiel_status_cols)
            st.dataframe(base_style, use_container_width=True)
