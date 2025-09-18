# -*- coding: utf-8 -*-

# ?? Modulos padrao da biblioteca Python
from pathlib import Path
from datetime import datetime
import logging

# ?? Manipulacao de dados e visualizacao
import pandas as pd

try:
    import plotly.express as px
    import plotly.graph_objects as go
except ModuleNotFoundError:
    import importlib.util as _importlib_util
    spec = _importlib_util.find_spec('plotly.express')
    if spec is None:
        import subprocess, sys
        pkg_dir = Path(__file__).resolve().parent / '_pkg_cache'
        pkg_dir.mkdir(exist_ok=True)
        cmd = [sys.executable, '-m', 'pip', 'install', 'plotly==5.24.0', '--target', str(pkg_dir)]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, check=False)
        output = result.stdout.decode('utf-8', errors='ignore') if isinstance(result.stdout, (bytes, bytearray)) else str(result.stdout)
        if result.returncode != 0:
            logging.warning('Fallback pip install plotly failed (%s): %s', result.returncode, output[:500])
        if str(pkg_dir) not in sys.path:
            sys.path.insert(0, str(pkg_dir))
    import plotly.express as px
    import plotly.graph_objects as go


# ?? Interface web com Streamlit
import streamlit as st

# ?? Modulo local (certifique-se de que o diretorio 'core' esta no mesmo nivel do script)
import core.etl_rollout as etl

# ?? Funcoes especificas do modulo etl_rollout
from core.etl_rollout import (
    clean_rollout_dataframe,
    kpi_from_explicit_cells,
    get_explicit_phase_map,
    sites_for_phase_explicit,
    last_status_snapshot,
    last_delay_days,
)


# Page setup and theme accents
st.set_page_config(page_title="Centro de Automacao", layout="wide")
ACCENT = "#F74949"

# Global CSS for sidebar minor tweaks (indent child items)
st.markdown(
    """
    <style>
    section[data-testid="stSidebar"] [data-testid="stCheckbox"],
    div[data-testid="stSidebar"] [data-testid="stCheckbox"] {
        margin: 2px 0 2px 14px;
    }
    section[data-testid="stSidebar"] [data-testid="stCheckbox"] label p,
    div[data-testid="stSidebar"] [data-testid="stCheckbox"] label p {
        font-size: 13px;
        margin: 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Small inline SVG icon for bar chart (no emojis)
ICON_BAR_SVG = (
    "<svg width=\"16\" height=\"16\" viewBox=\"0 0 24 24\" fill=\"#d0d0d0\" xmlns=\"http://www.w3.org/2000/svg\">"
    "<rect x=\"3\" y=\"12\" width=\"4\" height=\"9\" rx=\"1\"/>"
    "<rect x=\"10\" y=\"8\" width=\"4\" height=\"13\" rx=\"1\"/>"
    "<rect x=\"17\" y=\"4\" width=\"4\" height=\"17\" rx=\"1\"/>"
    "</svg>"
)

# Local path to persist uploaded Excel
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
SAVED_FILE = DATA_DIR / "rollout.xlsb"  # fallback legacy path
META_FILE = DATA_DIR / "rollout_meta.json"
try:
    import streamlit_antd_components as sac
    HAS_SAC = True
except Exception:
    HAS_SAC = False


# ---------------- Sidebar / Navegacao ----------------
if "route" not in st.session_state:
    st.session_state.route = "rollout"


def _nav_item_removed(*_args, **_kwargs):
    # Mantido como stub para compatibilidade com antigos trechos (desativados)
    # de navegacao. Nao e usado na execucao atual.
    pass

with st.sidebar:
    # Logo ZTE centralizada (tamanho medio)
    try:
        if Path("zte-logo.png").exists():
            _c1, _c2, _c3 = st.columns([1, 2, 1])
            with _c2:
                st.image("zte-logo.png", width=140)
    except Exception:
        pass

    # Sidebar: Rollout (expansivel) + checkbox simples
    st.session_state.setdefault("show_status", True)
    st.session_state.setdefault("show_lead", True)
    st.markdown("<div style='color:#9aa0a6; font-weight:600; font-size:13px; margin:6px 0 8px; display:flex; align-items:center;'>Automacoes<div style='flex:1; border-top:1px solid #3a3f44; margin-left:8px;'></div></div>", unsafe_allow_html=True)
    with st.expander("Rollout", expanded=True):
        st.checkbox("Visualizacao por Status", key="show_status")

        st.checkbox("Analise por Site (lead time)", key="show_lead")   
      
        st.session_state.setdefault("show_fiel", True)
        st.checkbox("Tabela Fiel/Real", key="show_fiel")

    # Secao e menu no estilo Ant Design
    st.session_state.setdefault("show_status", True)
    st.session_state.setdefault("route", "rollout")
    if False:
        try:
            try:
                sac.divider("Automacoes")
            except Exception:
                st.markdown(
                    "<div style='color:#9aa0a6; font-weight:600; font-size:13px; margin:6px 0 6px;'>Automacoes</div>",
                    unsafe_allow_html=True,
                )

            selected = sac.menu(
                items=[
                    sac.MenuItem(
                        "Rollout",
                        key="rollout",
                        children=[
                            sac.MenuItem("Visualizacao por Status", icon="bar-chart", key="rollout_status"),
                        ],
                    ),
                ],
                open_all=True,
                index=1,  # seleciona o primeiro filho
            )
            # Normaliza retorno
            sel_key = selected if isinstance(selected, str) else getattr(selected, "key", str(selected))
            if sel_key in ("rollout", "rollout_status"):
                st.session_state.route = "rollout"
                st.session_state.show_status = True
        except Exception:

            _nav_item_removed("Rollout", "rollout", indent=0)
            if "show_status" not in st.session_state:
                st.session_state["show_status"] = True

            st.checkbox("Visualizacao por Status", key="show_status")
    if False:
        # Fallback simples com o mesmo visual hierarquico
        st.markdown(
            "<div style='color:#9aa0a6; font-weight:600; font-size:13px; margin:6px 0 6px;'>Automacoes</div>",
            unsafe_allow_html=True,
        )
        _nav_item_removed("Rollout", "rollout", indent=0)
        if "show_status" not in st.session_state:
            st.session_state["show_status"] = True

        st.checkbox("Visualizacao por Status", key="show_status")


    st.markdown("---")
    st.markdown("<div style='text-align:center; color:#9aa0a6; font-size:12px;'>Centro de Automacao - Claro</div>", unsafe_allow_html=True)
@st.cache_data(show_spinner=False)
def read_excel_no_header(path: Path) -> pd.DataFrame:
    # header=None preserva as 7 linhas do topo (KPIs na linha 6)
    ext = path.suffix.lower()
    engine = None
    if ext == ".xlsb":
        engine = "pyxlsb"
    elif ext in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        engine = "openpyxl"
    elif ext == ".xls":
        engine = "xlrd"
    try:
        return pd.read_excel(path, sheet_name=0, header=None, engine=engine)
    except Exception as e:
        st.error(f"Falha ao ler {path.name} com engine {engine or 'auto'}: {e}")
        raise

def _load_saved_meta():
    try:
        import json
        if META_FILE.exists():
            return json.loads(META_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return None

def _save_meta(saved_path: Path, original_name: str):
    try:
        import json
        META_FILE.write_text(
            json.dumps(
                {
                    "saved_path": str(saved_path),
                    "original_name": original_name,
                    "uploaded_at": datetime.now().isoformat(timespec="seconds"),
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
    except Exception:
        pass

def _cleanup_saved_excels():
    try:
        patterns = ("*.xlsb", "*.xlsx", "*.xlsm", "*.xls")
        for pat in patterns:
            for p in DATA_DIR.glob(pat):
                try:
                    p.unlink()
                except Exception:
                    pass
    except Exception:
        pass


def dark(fig: go.Figure) -> go.Figure:
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_size=13,
        hoverlabel=dict(bgcolor="rgba(0,0,0,0.85)"),
        legend=dict(title=None),
        margin=dict(t=60, r=30, b=28, l=55),
    )
    return fig


# --- RESET GERAL (filtros + escopo + cliques) ---
def reset_all():
    # Filtros
    for k in (
        "sel_phase_full", "q_search", "f_uf", "f_reg", "f_subcon", "f_type",
        "f_model", "f_po", "f_year", "slider_lt",
    ):
        st.session_state.pop(k, None)

    # Visualizacao / Situacao (defaults)
    st.session_state["viz_type"] = "Barras"
    st.session_state["escopo"] = "Ambos"
    st.session_state["sit_radio"] = "Ambos"
    st.session_state["viz_type_radio"] = "Barras"

    # Se salvarmos selecoes por clique em session_state, limpe aqui tambem
    for k in ("clicked_phase_short", "clicked_serie"):
        st.session_state.pop(k, None)

    # Evita a mensagem "st.rerun() within a callback is a no-op"; a interface sera rerenderizada automaticamente
    # st.rerun()


# ---------------- Pagina Rollout ----------------
def _is_all_label(s) -> bool:
    return isinstance(s, str) and ("todas" in s.lower())


def request_reset():
    """Marca reset pendente e forca novo ciclo de execucao."""
    st.session_state["__do_reset__"] = True
    st.rerun()


def render_lead_analysis(df_raw: pd.DataFrame):
    """Renderiza a secao 'Analise por Site (lead time)'. Independente dos filtros."""
    phase_map = get_explicit_phase_map(df_raw)
     
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.markdown("<h2 style='margin: 6px 0 12px 0; font-size: 24px;'>Analise por Site (lead time)</h2>", unsafe_allow_html=True)
    with st.expander("Abrir analise por site", expanded=False):
        mode = st.radio("Modo", ["Media (todos os sites)", "Site especifico"], horizontal=True, key="site_analysis_mode")

        stay = etl.stage_stay_days(df_raw)
        phase_order = [s for (_f, s, _c) in phase_map]

        if mode == "Media (todos os sites)":
            means = {}
            for s in phase_order:
                col = f"stay_{s}"
                if col in stay.columns:
                    v = pd.to_numeric(stay[col], errors="coerce")
                    means[s] = float(v.mean(skipna=True)) if hasattr(v, "mean") else 0.0
            avg_df = pd.DataFrame({"fase_curta": list(means.keys()), "dias": list(means.values())})
            figm = px.bar(avg_df, x="fase_curta", y="dias", title="Tempo medio parado por status (dias)", text="dias")
            figm.update_traces(texttemplate="%{text:.1f}")
            figm.update_yaxes(title="Dias (media)")
            figm.update_xaxes(title="Status")
            figm = dark(figm)
            st.plotly_chart(figm, use_container_width=True, key="lead_avg_chart")
        else:
            uniq_sites = sorted(stay["SITE"].dropna().astype(str).unique().tolist())
            q_site = st.text_input("Pesquisar SITE", placeholder="Digite parte do SITE...")
            matches = [s for s in uniq_sites if q_site.strip().lower() in s.lower()] if q_site else []
            site_sel = st.selectbox("Selecionar SITE", matches, index=0) if matches else None
            if not matches and q_site:
                st.info("Nenhum SITE encontrado para a pesquisa.")

            if site_sel:
                row = stay[stay["SITE"].astype(str) == str(site_sel)].head(1)
                data = []
                total = 0
                for s in phase_order:
                    col = f"stay_{s}"
                    if col in row.columns:
                        val = float(pd.to_numeric(row[col], errors="coerce").fillna(0).iloc[0])
                        data.append({"fase_curta": s, "dias": val})
                        total += max(val, 0)
                site_df = pd.DataFrame(data)
                st.caption(f"Total decorrido (soma dos status) para {site_sel}: {int(total)} dias")
                figs = px.bar(site_df, x="fase_curta", y="dias", title=f"Tempo parado por status (dias) - {site_sel}", text="dias")
                figs.update_traces(texttemplate="%{text:.0f}")
                figs.update_yaxes(title="Dias")
                figs.update_xaxes(title="Status")
                figs = dark(figs)
                st.plotly_chart(figs, use_container_width=True, key=f"lead_site_chart_{site_sel}")

def _get_site_col_idx_from_raw(df_raw: pd.DataFrame) -> int:
    """Descobre a coluna SITE no df_raw (linha 7 do header)."""
    try:
        _site_col, _wide = etl._actuals_wide(df_raw)
        return int(_site_col)
    except Exception:
        header = pd.Series(df_raw.iloc[6]).astype(str).str.strip().str.upper()
        site_hit = header[header.isin(["SITE", "SITE NAME"])]
        return int(site_hit.index[0]) if not site_hit.empty else 4



def render_fiel_real(df_raw: pd.DataFrame, sites_f: pd.DataFrame):
    """
    Tabela Fiel/Real (Streamlit):
      - Le linhas 4..7 do Excel e monta L4 (status), L5 (Plan/RePlan/Actual/...),
        L6 (quantidade) e L7 (titulo da coluna).
      - Exibicao em 2 niveis (limite do Streamlit):
          Top  = L4 (status)
          Bot  = (L5 + '/ L6' se houver) + '  ' + L7
      - Respeita filtros (sites_f)
      - Datas sem horario (nunca em 'Qty')
      - Expander Opcoes da tabela (blocos por L4 e densidade)
      - Largura das colunas calculada a partir do tamanho do cabecalho
      - Download com colunas achatadas
    """
    import io

    # ---- Constantes (0-based) ----
    HEADER_TOP = 3      # linha 4
    HEADER_BOT = 6      # linha 7 (inclusive)
    BODY_START = 7      # dados a partir da linha 8

    # ---- Helpers ----
    def _norm(s):
        if s is None:
            return ""
        s = str(s).strip()
        return "" if s.lower() == "nan" else s

    def _up(s): return _norm(s).upper()

    def _is_numberlike(x):
        try:
            float(str(x).replace(",", "."))
            return True
        except Exception:
            return False

    def _find_start_idx_for_host(head_df, fallback=4):
        for j in range(head_df.shape[1]):
            for i in range(min(7, head_df.shape[0])):
                if _up(head_df.iat[i, j]) in {"HOST NAME", "HOSTNAME"}:
                    return j
        return fallback  # coluna E


    # ---- Aplica filtros por SITE ----
    # ---- Alinha as LINHAS com a mesma selecao da tabela "Visualizacao por Status" ----
    site_idx = _get_site_col_idx_from_raw(df_raw)
    body = df_raw.iloc[BODY_START:].copy()

    # 1) Comeca pelos filtros globais ja aplicados em sites_f
    keep_sites = set(sites_f["SITE"].astype(str))

    # 2) Se um status especifico estiver selecionado, aplica tambem Concluidos/Faltando
    final_sites = keep_sites.copy()
    try:
        try:
            import core.etl_rollout as etl
        except Exception:
            import core.etl_rollout as etl

        sel_label = st.session_state.get("sel_phase_full", "")
        is_all = (not sel_label) or str(sel_label).strip().lower().startswith(("todas", "all"))
        if not is_all:
            chosen_full = sel_label.split(" (")[0].strip()

            phase_map   = etl.get_explicit_phase_map(df_raw)
            full2short  = {f: s for (f, s, _c) in phase_map}
            order_short = [s for (f, s, _c) in phase_map]
            short2idx   = {s: i for i, s in enumerate(order_short)}
            chosen_short = full2short.get(chosen_full)

            concl_sites = set()
            if chosen_short:
                _site_col, wide_f = etl._actuals_wide(df_raw)
                if chosen_short in wide_f.columns:
                    concl_mask  = pd.to_datetime(wide_f[chosen_short], errors="coerce").notna()
                    concl_sites = set(wide_f.loc[concl_mask, "SITE"].astype(str))

            snap = etl.last_status_snapshot(df_raw)[["SITE", "last_phase_short"]].copy()
            def _next_short(x):
                i = short2idx.get(str(x), -1)
                i = min(i + 1, len(order_short) - 1)
                return order_short[i]
            snap["fase_curta"] = snap["last_phase_short"].map(_next_short)
            pend_sites = set(snap.loc[snap["fase_curta"] == chosen_short, "SITE"].astype(str))

            esc = st.session_state.get("escopo", "Ambos")
            if esc == "Concluidos":
                final_sites = keep_sites & concl_sites
            elif esc == "Faltando":
                final_sites = keep_sites & pend_sites
            else:  # Ambos
                final_sites = keep_sites & (concl_sites | pend_sites)
    except Exception:
        final_sites = keep_sites

    # 3) aplica o recorte final nas linhas da Tabela Fiel
    if final_sites:
        body = body[body.iloc[:, site_idx].astype(str).isin(final_sites)]




    # ---- Detecta inicio (Host Name) e captura cabecalhos 4..7 ----
    head0 = df_raw.iloc[:7]
    start_idx = _find_start_idx_for_host(head0, fallback=4)

    head = df_raw.iloc[HEADER_TOP:HEADER_BOT + 1, start_idx:].copy()  # 4..7
    body = body.iloc[:, start_idx:].reset_index(drop=True)

    # L4..L7 crus
    l4_raw = head.iloc[0].tolist()  # Status
    l5_raw = head.iloc[1].tolist()  # Plan/RePlan/Actual/...
    l6_raw = head.iloc[2].tolist()  # Quantidades
    l7_raw = head.iloc[3].tolist()  # Rotulo final (PPWS-PL, WHP-AC...)

    # L4: ffill para repetir o status por todas as colunas do bloco
    l4, last = [], ""
    for x in l4_raw:
        val = _norm(x)
        if val:
            last = val
        l4.append(last)

    # L5: so mantemos tokens de etapa; metadados ficam vazios aqui
    L5_KEEP = {
        "PLAN", "REPLAN", "RE-PLAN", "RPLAN",
        "ACTUAL", "AC", "ACT",
        "TIME", "GOAL", "ISSUE", "REASON", "WHO",
    }
    l5 = [(_norm(x) if _up(x) in L5_KEEP else "") for x in l5_raw]

    # L6: so mostra quando L5 existir e o valor parecer numerico
    l6 = [(_norm(x) if (l5[i] and _is_numberlike(x)) else "") for i, x in enumerate(l6_raw)]

    # L7: rotulo final
    l7 = [_norm(x) for x in l7_raw]

    # MultiIndex 4 niveis (para logica interna)
    cols4 = pd.MultiIndex.from_arrays([l4, l5, l6, l7], names=["L4", "L5", "L6", "L7"])
    df_all = body.copy()
    df_all.columns = cols4

    # ---- Opcoes da tabela ----
    st.markdown("<h3 style='margin: 18px 0 6px;'>Tabela Fiel</h3>", unsafe_allow_html=True)
    with st.expander("Opcoes da tabela", expanded=False):
        l4_all = [x for x in list(dict.fromkeys(df_all.columns.get_level_values(0))) if _norm(x)]
        show_blks = st.multiselect(
            "Blocos (linha 4 do Excel)", options=l4_all, default=l4_all,
            help="Selecione quais blocos/status deseja visualizar."
        )
        dens = st.radio(
            "Densidade por bloco",
            ["AC", "AC + Plan", "Completo"],
            horizontal=True,
            help="AC = so 'Actual'; AC+Plan = 'Plan' + 'Actual'; Completo = todas as colunas."
        )

    # ---- Selecao de colunas (blocos + densidade) ----
    l4_lv = df_all.columns.get_level_values(0).astype(str)
    l5_lv = df_all.columns.get_level_values(1).astype(str)

    essentials_lvl7 = {
        "HOST NAME", "HOSTNAME", "SITE NAME", "SITENAME", "STATE", "UF",
        "CURRENT STATUS", "GROUP", "SUBCON", "TYPE", "QTY", "MODEL",
        "SOW", "SOW TYPE", "SOW STATUS", "RECORD DATE",
    }
    keep_cols = [c for c in df_all.columns if _up(c[-1]) in essentials_lvl7]

    def _mask_dens():
        u = l5_lv.str.upper()
        if dens == "AC":
            return u.isin(["ACTUAL", "AC", "ACT"])
        if dens == "AC + Plan":
            return u.isin(["PLAN", "REPLAN", "RE-PLAN", "RPLAN", "ACTUAL", "AC", "ACT"])
        return pd.Series([True] * len(u))  # Completo

    mask_blk = l4_lv.isin(show_blks) if show_blks else pd.Series([True] * len(l4_lv))
    mask_den = _mask_dens()
    chosen_cols = df_all.columns[mask_blk & mask_den]

    # Ordem: essenciais + escolhidos (sem duplicar)
    seen, ordered = set(), []
    for c in list(keep_cols) + list(chosen_cols):
        if c not in seen:
            ordered.append(c)
            seen.add(c)
    df_sel = df_all.loc[:, ordered]

    # ---- Datas sem horario (NUNCA em 'Qty') ----
    def _looks_date_col(col_tuple):
        name_l7 = _up(col_tuple[-1])
        if name_l7 == "QTY":
            return False
        return (
            name_l7.endswith("-AC") or
            name_l7.endswith("-PL") or
            name_l7.endswith("-RPL") or
            name_l7 in {"RECORD DATE"}
        )

    for col in df_sel.columns:
        if _looks_date_col(col):
            ser = pd.to_datetime(df_sel[col], errors="coerce")
            if ser.notna().sum() > 0:
                df_sel[col] = ser.dt.strftime("%d-%b-%y").where(ser.notna(), df_sel[col])

    # ---- COMPACTACAO para 2 niveis de exibicao ----
    lvl0, lvl1 = [], []
    for (L4, L5, L6, L7) in df_sel.columns:
        top = _norm(L4)  # sempre o status aqui
        bot_parts = []
        if _norm(L5):
            bot_parts.append(L5)
        if _norm(L6):
            bot_parts.append(f"{L6}")
        if _norm(L7):
            bot_parts.append(f" {L7}")
        bot = " / ".join(bot_parts) if bot_parts else (_norm(L7) or _norm(L4))
        lvl0.append(top)
        lvl1.append(bot)

    cols2 = pd.MultiIndex.from_arrays([lvl0, lvl1])
    df_view = df_sel.copy()
    df_view.columns = cols2

    # ---- Largura automatica por tamanho de cabecalho ----
    colcfg = {}
    def _auto_w(s: str):
        # largura minima 120, maxima 360, proporcional ao texto
        n = max(len(s), 8)
        return max(120, min(360, int(n * 7.2)))

    # default: auto em todas
    for col in df_view.columns:
        header_len = len(str(col[0])) + len(str(col[1]))
        try:
            colcfg[col] = st.column_config.Column(width=_auto_w("".join([str(col[0]), str(col[1])])))
        except Exception:
            pass

    # especificos menores
    for name, width in [("Host Name", 160), ("Site Name", 160), ("State", 90)]:
        for col in df_view.columns:
            if _norm(col[1]).lower().endswith(name.lower()) or _norm(col[1]).lower() == name.lower():
                try:
                    colcfg[col] = st.column_config.Column(width=width)
                except Exception:
                    pass

    # ---- Render ----
    try:
        st.dataframe(df_view, use_container_width=True, height=520, column_config=colcfg)
    except Exception:
        st.dataframe(df_view, use_container_width=True, height=520)

    # ---- Download (achatado) ----
    df_xlsx = df_sel.copy()
    flat_cols = []
    for col in df_xlsx.columns:
        parts = [p for p in col if _norm(p)]
        flat_cols.append(" / ".join(parts) if parts else "")
    df_xlsx.columns = flat_cols

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_xlsx.to_excel(writer, index=False, sheet_name="FIEL_REAL")
    bio.seek(0)
    st.download_button(
        "Baixar Fiel/Real (recorte).xlsx",
        data=bio.getvalue(),
        file_name="FIEL_REAL_filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_fiel_real_streamlit",
    )




def page_rollout():
    # Aplica reset pendente ANTES de criar widgets (evita conflito de keys)
    if st.session_state.get("__do_reset__", False):
        for _k in (
            "sel_phase_full", "q_search", "f_uf", "f_reg", "f_subcon", "f_type",
            "f_model", "f_po", "f_year", "slider_lt",
            "viz_type_radio", "sit_radio", "sel_phase_box",
            "clicked_phase_short", "clicked_serie",
        ):
            st.session_state.pop(_k, None)
        # Defaults de estado (nao sao widgets)
        st.session_state["viz_type"] = "Barras"
        st.session_state["escopo"] = "Ambos"
        st.session_state["sel_phase_full"] = "Todas"
        st.session_state["__do_reset__"] = False
    st.title("Rollout Claro RAN - Overview")
    st.caption("Suba o Excel (.xlsb, .xlsx) e acompanhe KPIs e detalhamento por status.")

    # ========== 1) Upload + Carregar ==========
    uploaded = st.file_uploader(
        "Upload do arquivo Excel (.xlsb, .xlsx)",
        type=["xlsb", "xlsx", "xlsm"],
        accept_multiple_files=False,
    )
    if uploaded is not None:
        ext = (Path(uploaded.name).suffix or "").lower()
        if ext not in {".xlsb", ".xlsx", ".xlsm"}:
            st.error(f"Extensao nao suportada: {ext}")
        else:
            safe_name = Path(uploaded.name).name
            saved_path = DATA_DIR / safe_name
            _cleanup_saved_excels()
            with open(saved_path, "wb") as f:
                f.write(uploaded.getbuffer())
            _save_meta(saved_path, uploaded.name)
            if uploaded.name == saved_path.name:
                st.success(f"Arquivo salvo: {saved_path.name}")
            else:
                st.success(f"Arquivo salvo: {uploaded.name} ? {saved_path.name}")
            try:
                ts = datetime.now().strftime("%d/%m/%Y %H:%M")
                st.caption(f"Enviado agora ({ts})")
            except Exception:
                pass

    meta = _load_saved_meta()
    # Mostrar info do arquivo salvo (se existir), mesmo sem upload novo
    shown_info = False
    if meta and Path(meta.get("saved_path", "")).exists():
        sp = Path(meta["saved_path"]).name
        on = meta.get("original_name", sp)
        dt = meta.get("uploaded_at", "")
        dt_disp = dt.replace("T", " | ") if isinstance(dt, str) else dt
        if on == sp:
            st.caption(f"Arquivo atual: {on}  enviado em {dt_disp}")
        else:
            st.caption(f"Arquivo atual: {on} (salvo como {sp})  enviado em {dt_disp}")
        shown_info = True
    else:
        # Fallback: procurar qualquer Excel salvo no diretorio de dados
        cands = []
        for pat in ("*.xlsb", "*.xlsx", "*.xlsm"):
            cands.extend(DATA_DIR.glob(pat))
        if cands:
            cand = sorted(cands, key=lambda p: p.stat().st_mtime)[-1]
            try:
                ts = datetime.fromtimestamp(cand.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
            except Exception:
                ts = ""
            st.caption(f"Arquivo atual: {cand.name}  salvo em {ts}")
            shown_info = True

    if st.button("Carregar planilha", key="btn_load"):
        target_path = None
        if meta and Path(meta.get("saved_path", "")).exists():
            target_path = Path(meta["saved_path"])    
        elif SAVED_FILE.exists():  # fallback legado .xlsb
            target_path = SAVED_FILE
        else:
            cands = []
            for pat in ("*.xlsb", "*.xlsx", "*.xlsm"):
                cands.extend(DATA_DIR.glob(pat))
            if cands:
                target_path = sorted(cands, key=lambda p: p.stat().st_mtime)[-1]
        if not target_path:
            st.error("Nenhum arquivo salvo. Faca o upload primeiro.")
            st.stop()
        with st.spinner("Lendo e tratando..."):
            df_raw = read_excel_no_header(target_path)
            df_clean, df_header = clean_rollout_dataframe(df_raw)
        st.session_state.rollout_df_raw = df_raw
        st.session_state.rollout_df_clean = df_clean
        st.session_state.rollout_df_header = df_header
        try:
            st.session_state.rollout_file_path = str(target_path)
        except Exception:
            pass
        st.success("Planilha carregada!")

    if "rollout_df_raw" not in st.session_state:
        st.info("Carregue a planilha para continuar.")
        return

    df_raw = st.session_state.rollout_df_raw
    df_clean = st.session_state.rollout_df_clean

    # Pequeno espaco entre upload e a primeira tabela
    st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)

    # ========== 2) KPIs por fase (linha 6) ==========
    kpi = kpi_from_explicit_cells(df_raw, method="header", sites_col_letter="E").copy()
    kpi = kpi.rename(columns={"qtd": "Concluidos"})
    # Computa total a partir de Concluidos + faltam
    kpi["total"] = kpi["Concluidos"] + kpi["faltam"]

    # Visualizacao por Status (unica secao atual)
    if not st.session_state.get("show_status", True):
        if st.session_state.get("show_lead", True):
            render_lead_analysis(df_raw)
        return
    
    # Titulo grande para a secao
    st.markdown(
        """
        <h2 style='margin: 6px 0 12px 0; font-size: 28px;'>Visualizacao por Status</h2>
        """,
        unsafe_allow_html=True,
    )

    # Tabela 'Status em geral (arquivo)'
    with st.expander("Status em geral (Overview)", expanded=False):
        st.dataframe(
            kpi[["fase_curta", "Concluidos", "faltam", "total"]].set_index("fase_curta"),
            use_container_width=True,
        )
    
    
    

    # Mapeamentos: agora incluem as fases iniciais (1.x, 2.x, 3.1, 4.1, ...)
    phase_map = get_explicit_phase_map(df_raw)  # [(full, short, col_idx)]
    full_to_short = {full: short for (full, short, _c) in phase_map}
    short_to_full = {short: full for (full, short, _c) in phase_map}
    phase_list_full = [full for (full, _s, _c) in phase_map]

    # Estados padrao
    st.session_state.setdefault("viz_type", "Barras")
    st.session_state.setdefault("escopo", "Ambos")
    st.session_state.setdefault("sel_phase_full", "Todas")
    st.session_state.setdefault("q_search", "")
    for _k in ("f_uf", "f_reg", "f_subcon", "f_type", "f_model", "f_po", "f_year", "f_carimbo"):
        st.session_state.setdefault(_k, [])

    # ========== 4) Tipo de grafico e Escopo ==========
    col_viz, col_sit, col_reset = st.columns([1, 1, 0.25])
    viz_type = col_viz.radio(
        "Visualizacao",
        ["Barras", "Pizza"],
        horizontal=True,
        index=0 if st.session_state.get("viz_type", "Barras") == "Barras" else 1,
        key="viz_type_radio",
    )
    st.session_state["viz_type"] = viz_type

    sit_opts = ["Concluidos", "Faltando"] + (["Ambos"] if viz_type == "Barras" else [])
    default_sit = st.session_state.get("escopo", "Ambos")
    if default_sit not in sit_opts:
        default_sit = "Concluidos" if viz_type == "Pizza" else "Ambos"
    Situacao = col_sit.radio("Situacao", sit_opts, horizontal=True, index=sit_opts.index(default_sit), key="sit_radio")
    st.session_state["escopo"] = Situacao

    if col_reset.button("Resetar", use_container_width=True, key="btn_reset_all"):
        request_reset()



    # ========== 5) e 6) Filtros ==========
    with st.expander("Filtros", expanded=True):
        # Layout superior: status + pesquisa
        top1, top2 = st.columns([1.1, 1.9])

        # Filtro de status
        #Construcao da variavel sites
        order_full = [full for (full, short, _c) in phase_map]
        order_short = [short for (full, short, _c) in phase_map]
        short2idx = {s: i for i, s in enumerate(order_short)}
        full2short = {full: short for (full, short, _c) in phase_map}
        
        
        # Se nao vier pronto do ETL, derive o 'full' pelo mapa
        short2full = {s: f for (f, s, _c) in phase_map}
        # Snapshot (sempre com last_phase_full)
        snap = last_status_snapshot(df_raw)[
            ["SITE", "last_phase_short", "last_phase_full", "last_date"]
        ].copy()

        # Normalizacoes e campos derivados
        snap["last_phase_short"] = (
            snap["last_phase_short"].astype(str).where(snap["last_phase_short"].isin(order_short), None)
        )
        snap["last_idx"] = snap["last_phase_short"].map(short2idx)
        snap["current_idx"] = snap["last_idx"].fillna(-1).astype(int) + 1
        snap["current_idx"] = snap["current_idx"].clip(0, len(order_short) - 1)
        snap["current_short"] = snap["current_idx"].map(lambda i: order_short[i])
        snap["current_full"]  = snap["current_short"].map(short2full)


        cols_keep = [c for c in [
            "SITE", "state", "Group", "Subcon", "Type", "Qty", "Model", "Infra PO", "current_status", "year"
        ] if c in df_clean.columns]
        static = df_clean[cols_keep].drop_duplicates(subset=["SITE"]).copy()
        static = static.rename(columns={"Infra PO": "PO", "Group": "Regional"})
        if "UF" not in static.columns and "state" in static.columns:
            static["UF"] = static["state"]

        delay_df = last_delay_days(df_raw)[["SITE", "delay_days"]].copy()
        delay_df["delay_days"] = pd.to_numeric(delay_df["delay_days"], errors="coerce").fillna(0).astype(int)

        # Final merge para construir sites
        sites = snap.merge(static, on="SITE", how="left").merge(delay_df, on="SITE", how="left")

        # Enriquecer com ano vindo do base_all (quando existir)
        try:
            _site_year = base_all[["SITE", "year"]].dropna().drop_duplicates(subset=["SITE"])
            sites = sites.merge(_site_year, on="SITE", how="left", suffixes=("", "_from_base"))
            if "year_from_base" in sites.columns:
                sites["year"] = sites["year_from_base"].where(sites["year_from_base"].notna(), sites.get("year"))
                sites = sites.drop(columns=["year_from_base"])
        except Exception:
            pass


        status_labels = ["Todas"] + [f"{full} ({full_to_short[full]})" for full in phase_list_full]
        sel_status_label = top1.selectbox(
            "Selecione o status", status_labels,
            index=status_labels.index(st.session_state.sel_phase_full)
            if st.session_state.sel_phase_full in status_labels else 0,
            key="sel_phase_box",
        )
        st.session_state.sel_phase_full = sel_status_label

        # Filtro de pesquisa por termos
        st.session_state.setdefault("q_terms", [])
        def _add_q_term():
            val = st.session_state.get("q_search_new", "").strip()
            if val:
                parts = [p.strip() for p in val.replace(";", "\n").replace(",", "\n").splitlines() if p.strip()]
                cur = list(st.session_state.get("q_terms", []))
                for p in parts:
                    if not any(p.lower() == c.lower() for c in cur):
                        cur.append(p)
                st.session_state["q_terms"] = cur
            st.session_state["q_search_new"] = ""

        top2.text_input(
            "Pesquisar (SITE, status, UF/Regional, Subcon, Type, Model, PO)",
            placeholder="Digite e pressione Enter para adicionar",
            key="q_search_new",
            on_change=_add_q_term,
        )

        # Mostrar termos adicionados
        try:
            if st.session_state.get("q_terms"):
                chips = st.session_state.get("q_terms", [])
                _cchips = st.container()
                with _cchips:
                    st.write("Pesquisas:", ", ".join([f"'{t}'" for t in chips]))
                    if st.button("Limpar pesquisas", key="btn_clear_terms"):
                        st.session_state["q_terms"] = []
        except Exception:
            pass

        # Layout dos filtros visuais
        r1c1, r1c2, r1c3, r1c4 = st.columns(4)
        r2c1, r2c2, r2c3, r2c4 = st.columns(4)
        r3c1, r3c2, r3c3, r3c4 = st.columns(4)

        # Construcao do base_all com fases
        frames = []
        for full, short, _c in phase_map:
            tmp = sites_for_phase_explicit(df_raw, df_clean, full)
            tmp["fase_label"] = full
            tmp["fase_curta"] = short
            frames.append(tmp)
        base_all = pd.concat(frames, ignore_index=True)

        # Normalizacoes
        if "Infra PO" in base_all.columns and "PO" not in base_all.columns:
            base_all = base_all.rename(columns={"Infra PO": "PO"})
        if "Group" in base_all.columns and "Regional" not in base_all.columns:
            base_all = base_all.rename(columns={"Group": "Regional"})
        if "UF" not in base_all.columns and "state" in base_all.columns:
            base_all["UF"] = base_all["state"]

        # Snapshot e atraso
        snap_all = last_status_snapshot(df_raw)[["SITE", "last_phase_short", "last_date"]]
        delay_all = last_delay_days(df_raw)[["SITE", "delay_days"]]
        base_all = base_all.merge(snap_all, on="SITE", how="left").merge(delay_all, on="SITE", how="left")

        # Ano
        if "year" not in base_all.columns or base_all["year"].isna().all():
            year_guess = pd.to_datetime(base_all.get("actual_date"), errors="coerce").dt.year
            year_guess = year_guess.fillna(pd.to_datetime(base_all.get("last_date"), errors="coerce").dt.year)
            base_all["year"] = year_guess

        # Filtros visuais com base no df_clean
        uf_opts = sorted(df_clean.get("state", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        reg_opts = sorted(df_clean.get("Group", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        subcon_opts = sorted(df_clean.get("Subcon", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        type_opts = sorted(df_clean.get("Type", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        model_opts = sorted(df_clean.get("Model", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        po_opts = sorted(df_clean.get("Infra PO", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())

        sel_uf = r1c1.multiselect("UF", uf_opts, default=[], key="f_uf")
        sel_reg = r1c2.multiselect("Regional", reg_opts, default=[], key="f_reg")
        sel_subcon = r1c3.multiselect("Subcon", subcon_opts, default=[], key="f_subcon")
        sel_type = r1c4.multiselect("Type", type_opts, default=[], key="f_type")
        sel_model = r2c1.multiselect("Model", model_opts, default=[], key="f_model")
        sel_po = r2c2.multiselect("PO", po_opts, default=[], key="f_po")

        # Filtro de ano
        year_opts = sorted([int(y) for y in pd.to_numeric(base_all["year"], errors="coerce").dropna().unique().tolist()])
        sel_year = r2c3.multiselect("Ano", year_opts, default=[], key="f_year")

        # Filtro de carimbo (se existir)
        sel_carimbo = []
        try:
            carimbo_opts = sorted(base_all["Carimbo"].dropna().astype(str).unique().tolist())
            st.session_state.setdefault("f_carimbo", [])
            sel_carimbo = r3c1.multiselect("Carimbo", carimbo_opts, default=[], key="f_carimbo")
        except KeyError:
            pass

        # Filtro de lead time
        lt_series = pd.to_numeric(base_all.get("delay_days", pd.Series(index=base_all.index)), errors="coerce").fillna(0)
        lt_min = 0
        lt_max = int(max(lt_series.max(), 0))
        if lt_max == lt_min:
            lt_max = lt_min + 1
        if "slider_lt" not in st.session_state:
            st.session_state["slider_lt"] = (int(lt_min), int(lt_max))
        lt_sel = r2c4.slider("Lead time (dias)", min_value=int(lt_min), max_value=int(lt_max),
                            value=st.session_state.get("slider_lt", (int(lt_min), int(lt_max))),
                            step=1, key="slider_lt")



    # garantir que 'Carimbo' exista em `sites`
    if "Carimbo" not in sites.columns and "Carimbo" in base_all.columns:
        tmp = base_all[["SITE", "Carimbo"]].copy()
        tmp["SITE"] = tmp["SITE"].astype(str)
        sites["SITE"] = sites["SITE"].astype(str)
        sites = sites.merge(tmp, on="SITE", how="left")


    # 6.2) Aplicar filtros da UI ao snapshot
    mask_sites = pd.Series(True, index=sites.index)

    # Nao restringe por "Selecione o status" aqui; deixa a filtragem fina para a etapa da tabela

    if sel_uf:
        mask_sites &= sites["UF"].astype(str).isin(sel_uf)
    if sel_reg:
        mask_sites &= sites["Regional"].astype(str).isin(sel_reg)
    if sel_subcon:
        mask_sites &= sites["Subcon"].astype(str).isin(sel_subcon)
    if sel_type:
        mask_sites &= sites["Type"].astype(str).isin(sel_type)
    if sel_model:
        mask_sites &= sites["Model"].astype(str).isin(sel_model)
    if sel_po:
        mask_sites &= sites["PO"].astype(str).isin(sel_po)
    if sel_year:
        mask_sites &= pd.to_numeric(sites["year"], errors="coerce").astype("Int64").isin(sel_year)
    if 'sel_carimbo' in locals() and sel_carimbo:
        if "Carimbo" in sites.columns:
            mask_sites &= sites["Carimbo"].astype(str).isin(sel_carimbo)

    lt_low, lt_high = (lt_sel if "lt_sel" in locals() else (0, 10**9))
    sites["delay_days"] = pd.to_numeric(sites["delay_days"], errors="coerce").fillna(0)
    mask_sites &= sites["delay_days"].between(lt_low, lt_high)

    # Pesquisa por multiplos termos
    terms = [t.strip() for t in st.session_state.get("q_terms", []) if str(t).strip()]
    if terms:
        cols_search = [c for c in [
            "SITE", "current_full", "current_short", "last_phase_short", "UF", "Regional", "Subcon", "Type", "Model", "PO"
        ] if c in sites.columns]
        mm_any = pd.Series(False, index=sites.index)
        for term in terms:
            ql = str(term).strip().lower()
            if not ql:
                continue
            mm = pd.Series(False, index=sites.index)
            for c in cols_search:
                mm |= sites[c].astype(str).str.lower().str.contains(ql, na=False)
            mm_any |= mm
        mask_sites &= mm_any

    sites_f = sites.loc[mask_sites].drop_duplicates(subset=["SITE"]).reset_index(drop=True).copy()
    total_sites = int(sites_f["SITE"].nunique())

    # 6.3) Contagens por fase (Concluidos x Faltando) - respeita filtros (por coluna Actual)
    try:
        _sites_col, wide = etl._actuals_wide(df_raw)
    except Exception:
        header = pd.Series(df_raw.iloc[6]).astype(str).str.strip().str.upper()
        site_hit = header[header.isin(["SITE", "SITE NAME"])]
        site_idx = int(site_hit.index[0]) if not site_hit.empty else 4
        wide = pd.DataFrame({"SITE": pd.Series(df_raw.iloc[7:, site_idx])})
        for _f, s, idx in phase_map:
            wide[str(s)] = pd.to_datetime(pd.Series(df_raw.iloc[7:, idx]), errors="coerce")
        wide = wide.dropna(subset=["SITE"]).reset_index(drop=True)
    filt_set = set(sites_f["SITE"].astype(str))
    wide_f = wide[wide["SITE"].astype(str).isin(filt_set)].drop_duplicates(subset=["SITE"]).copy()
    concluded_counts = {s: int(pd.to_datetime(wide_f.get(s), errors="coerce").notna().sum()) for s in order_short}
    bars = pd.DataFrame({
        "fase_curta": order_short,
        "Concluidos": [concluded_counts[s] for s in order_short],
        "Faltando": [max(0, total_sites - concluded_counts[s]) for s in order_short],
    })
    # Se um status especifico foi selecionado, limita barras
    if not _is_all_label(st.session_state.sel_phase_full):
        _cf = st.session_state.sel_phase_full.split(" (")[0]
        _cs = full2short.get(_cf, None)
        if _cs in set(order_short):
            bars = bars[bars["fase_curta"] == _cs]
    bars["total"] = bars["Concluidos"] + bars["Faltando"]
    # Sufixo de data/hora para titulos (usa meta; fallback: mtime do arquivo)
    def _current_data_suffix():
        try:
            meta2 = _load_saved_meta()
            if meta2 and meta2.get("uploaded_at"):
                dt = str(meta2.get("uploaded_at"))
                return dt.replace("T", " | ")
        except Exception:
            pass
        try:
            p = st.session_state.get("rollout_file_path")
            if p:
                from datetime import datetime as _dt
                ts = _dt.fromtimestamp(Path(p).stat().st_mtime)
                return ts.strftime("%Y-%m-%d | %H:%M:%S")
        except Exception:
            pass
        return None
    _ts_suffix = _current_data_suffix()
    if viz_type == "Barras":
        if Situacao == "Ambos":
            long = bars.melt(
                id_vars=["fase_curta"],
                value_vars=["Concluidos", "Faltando"],
                var_name="tipo",
                value_name="valor",
            )
        else:
            keep = Situacao
            long = bars.rename(columns={keep: "valor"})[["fase_curta", "valor"]].assign(tipo=keep)
        ymax = max(int(bars["total"].max()), 1)
        fig = px.bar(
            long, x="fase_curta", y="valor", color="tipo",
            color_discrete_map={"Concluidos": "#1f77b4", "Faltando": "#ff7f0e"},
            category_orders={
                "tipo": ["Concluidos", "Faltando"],
                "fase_curta": order_short,
            },
            text="valor",
            barmode="stack" if Situacao == "Ambos" else "relative",
            title=("Sites por status (concluidos x faltando)" + (f" | {_ts_suffix}" if _ts_suffix else "")),
        )
        
        
        fig.update_traces(
            texttemplate="%{text}",
            hovertemplate="<b>%{x}</b><br>%{data.name}: <b>%{y}</b><extra></extra>",
        )
        fig.update_yaxes(range=[0, ymax * 1.18], title="Quantidade de sites")
        fig.update_xaxes(title="Status")
        fig.for_each_trace(lambda t: t.update(textposition="outside") if t.name == "Faltando" else t.update(textposition="inside"))
        fig = dark(fig)
    else:
        col = "Concluidos" if Situacao == "Concluidos" else "Faltando"
        pie_df = bars.rename(columns={col: "valor"})[["fase_curta", "valor"]]
        fig = px.pie(
            pie_df,
            names="fase_curta",
            values="valor",
            title=(f"Distribuicao por status - {Situacao}" + (f" | {_ts_suffix}" if _ts_suffix else "")),
            hole=0.35,
        )
        fig.update_traces(textinfo="value+percent", hovertemplate="<b>%{label}</b><br>" + Situacao + ": <b>%{value}</b> (%{percent})<extra></extra>")
        fig = dark(fig)

    # Renderiza sem eventos de clique
    st.plotly_chart(fig, use_container_width=True, key="status_main_chart")



    table_df = sites_f.copy()

    # --- 1) fase_label/fase_curta = PROXIMO status apos o ultimo concluido (pendente) ---
    # Usamos a ordem vinda do phase_map ja montado acima
    # order_short, short2idx, short2full ja existem no trecho anterior
    def _next_short(last_short: str) -> str:
        i = short2idx.get(str(last_short), -1)
        i_next = min(i + 1, len(order_short) - 1)
        return order_short[i_next]

    table_df["fase_curta"] = table_df.get("last_phase_short").map(_next_short)
    table_df["fase_label"] = table_df["fase_curta"].map(short2full)

    # --- 2) Se um status foi escolhido, filtrar conforme Situacao ---
    if not _is_all_label(st.session_state.sel_phase_full):
        chosen_full = st.session_state.sel_phase_full.split(" (")[0].strip()
        chosen_short = full2short.get(chosen_full)

        # Concluidos: presenca de data na coluna AC do status escolhido (mesma base do grafico)
        concl_sites = set()
        if chosen_short and chosen_short in wide_f.columns:
            concl_mask  = pd.to_datetime(wide_f[chosen_short], errors="coerce").notna()
            concl_sites = set(wide_f.loc[concl_mask, "SITE"].astype(str))

        # Faltando: exatamente os que estao PENDENTES nesse status (fase_curta == chosen_short)
        pend_mask = table_df["fase_curta"].astype(str) == str(chosen_short)

        esc = st.session_state.get("escopo", "Ambos")
        if esc == "Concluidos":
            table_df = table_df[table_df["SITE"].astype(str).isin(concl_sites)]
        elif esc == "Faltando":
            table_df = table_df[pend_mask]
        else:  # Ambos
            table_df = table_df[pend_mask | table_df["SITE"].astype(str).isin(concl_sites)]

    # (Opcional) badge de situacao para o status selecionado
    # if not _is_all_label(st.session_state.sel_phase_full) and chosen_short and chosen_short in wide_f.columns:
    #     table_df["sit_selected"] = table_df["SITE"].astype(str).map(
    #         lambda s: "Concluido" if s in concl_sites else ("Faltando" if s in set(table_df.loc[pend_mask, "SITE"].astype(str)) else "")
    #     )

    # Fallback de current_status (se faltar)
    if "current_status" not in table_df.columns:
        table_df["current_status"] = table_df.get("current_full", table_df.get("fase_label"))

    cols_order = [c for c in [
        "SITE","UF","Regional","current_status","fase_label","fase_curta",
        "last_date","delay_days","year","Subcon","Type","Model","PO", "Carimbo"
        # "sit_selected",  # (opcional) se ativar o badge acima
    ] if c in table_df.columns]

    st.dataframe(table_df.reset_index(drop=True)[cols_order], use_container_width=True, height=430)



    # ========== 7) Analise de lead time ==========
    if st.session_state.get("show_lead", True):
        render_lead_analysis(df_raw)
        
        
        
        
        
    # ---- Tabela Fiel/Real (respeita filtros) ----
    if st.session_state.get("show_fiel", True):
        render_fiel_real(df_raw, sites_f)

        
    


# ---------------- Router ----------------
if st.session_state.route == "rollout": 
    page_rollout()
else:
    page_rollout()





