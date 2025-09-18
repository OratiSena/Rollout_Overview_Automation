import pandas as pd
from datetime import datetime, timedelta

# ============== Limpeza ==============
def clean_rollout_dataframe(df_raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    header_rows = 7
    df_header = df_raw.iloc[:header_rows, :]
    df_data = df_raw.iloc[header_rows:, :].copy()

    new_header = df_raw.iloc[6].tolist()              # linha 7 do Excel (index 6)
    df_data.columns = new_header
    df_data = df_data.loc[:, ~df_data.columns.isna()] # remove colunas sem header
    df_data = df_data.dropna(how="all").reset_index(drop=True)

    rename_map = {
        "Host Name": "host_name",
        "Site Name": "SITE",
        "State": "state",
        "Current Status": "current_status",
    }
    df_data = df_data.rename(columns=rename_map)
    return df_data, df_header

"""Funções de ETL do Rollout.

Notas de limpeza: constantes antigas de fases foram removidas por não serem
mais utilizadas; o mapeamento é derivado dinamicamente via
`get_explicit_phase_map(df_raw)`.
"""

# ============== Helpers Excel ==============
def excel_col_to_idx(col: str) -> int:
    col = col.strip().upper()
    acc = 0
    for ch in col:
        acc = acc * 26 + (ord(ch) - 64)
    return acc - 1

def _to_excel_datetime(val):
    if pd.isna(val):
        return pd.NaT
    if isinstance(val, (pd.Timestamp, datetime)):
        return pd.to_datetime(val)
    if isinstance(val, str):
        dt = pd.to_datetime(val, errors="coerce")
        return dt if pd.notna(dt) else pd.NaT
    if isinstance(val, (int, float)):
        base = datetime(1899, 12, 30)
        return pd.Timestamp(base + timedelta(days=float(val)))
    return pd.NaT

# ============== KPI por células explícitas (linha 6) ==============
_EXPLICIT_PHASES = [
    # This legacy constant is kept for backward compatibility but is no longer
    # used to compute KPIs. KPIs are now derived dynamically from the sheet
    # header by get_explicit_phase_map(df_raw).
]

def _excel_idx_to_col(idx: int) -> str:
    """Convert 0-based column index to Excel column letters."""
    idx += 1
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s

def kpi_from_explicit_cells(df_raw: pd.DataFrame, method: str = "header", sites_col_letter: str = "E") -> pd.DataFrame:
    # Descobre coluna de SITE por nome para robustez
    header = pd.Series(df_raw.iloc[6]).astype(str).str.strip().str.upper()
    site_idx = header[header.isin(["SITE", "SITE NAME"])].index
    if len(site_idx) > 0:
        sites_idx = int(site_idx[0])
    else:
        sites_idx = excel_col_to_idx(sites_col_letter)

    total_sites = int(pd.Series(df_raw.iloc[7:, sites_idx]).notna().sum())

    rows = []
    for full, short, cidx in get_explicit_phase_map(df_raw):
        if method == "header":
            raw_val = df_raw.iloc[5, cidx]  # linha 6 (contagem informada no topo)
            try:
                if isinstance(raw_val, str):
                    raw_val = raw_val.replace(" ", "").replace(".", "").replace(",", ".")
                qtd = int(float(raw_val))
            except Exception:
                qtd = int(pd.Series(df_raw.iloc[7:, cidx]).notna().sum())
        else:
            qtd = int(pd.Series(df_raw.iloc[7:, cidx]).notna().sum())

        faltam = max(0, total_sites - qtd)
        rows.append({
            "fase": full,
            "fase_curta": short,
            "col": _excel_idx_to_col(cidx),
            "qtd": qtd,
            "faltam": faltam,
            "total_sites": total_sites,
        })
    return pd.DataFrame(rows)

def get_explicit_phase_map(df_raw: pd.DataFrame) -> list[tuple[str,str,int]]:
    """Return ordered phases with column indexes taken from the sheet header.

    Extends the original 5.1..7.4 set adding initial phases so that sites
    before MRF don't all collapse into MRF. We match by header names "*-AC"
    (actual dates) found in the sheet header (row 7, index 6).

    Order:
      1.1 PPO-AC, 1.2 IPO-AC, 1.3 SV-AC, 1.4 IPP-AC, 1.5 APP-AC,
      2.3 RFP-AC,
      3.1 RFI-AC, 4.1 TXA-AC,
      5.1 MRQ-AC, 5.2 RDO-AC, 5.3 RPI-AC, 5.4 WHP-AC, 5.5 API-AC, 5.6 INV-AC, 5.7 DEL-AC, 5.8 POD-AC,
      6.1 I&C-AC, 6.2 INT-AC,
      7.1 IDP-AC, 7.2 ARQ-AC, 7.3 PAC-AC, 7.4 FAC-AC
    """
    desired = [
    # 1.x
    ("1.1- PWS Pre-PO",          "PPO",  ["PPO-AC", "PPWS-AC", "PWS PRE-PO-AC", "PWS PREPO-AC", "PWS PRE PO-AC"]),
    ("1.2- Issue PWS PO",        "IPO",  ["IPO-AC", "IPWS-AC", "PWS PO-AC", "PWSPO-AC"]),
    ("1.3- Site Survey",         "SV",   ["SSV-AC"]),
    ("1.4- Issue Project",       "IPP",  ["IPP-AC"]),
    ("1.5- Approve Project",     "APP",  ["APP-AC"]),

    # 2.x
    ("2.3- Infra Pre-PO",        "PINF", ["PINF-AC"]),

    # 3.x, 4.x ...
    ("3.1- Ready for Install",   "RFI",  ["RFI-AC"]),
    ("4.1- Transmission",        "TXA",  ["TXA-AC"]),

    # 5.x ...
    ("5.1- Material Request",    "MRF",  ["MRQ-AC"]),
    ("5.2- DOMRF",               "DO",   ["RDO-AC"]),
    ("5.3- Raise Pre-Invoice",   "RPI",  ["RPI-AC"]),
    ("5.4- WH Picking",          "WHP",  ["WHP-AC"]),
    ("5.5- Approve Pre-Invoice", "API",  ["API-AC"]),
    ("5.6- Issue Invoice",       "FAT",  ["INV-AC"]),
    ("5.7- Delivery",            "MOS",  ["DEL-AC"]),
    ("5.8- Upload PoD",          "POD",  ["POD-AC"]),

    # 6.x
    ("6.1- Install",             "INST", ["I&C-AC", "INC-AC", "INS-AC"]),
    ("6.2- Integrate",           "INT",  ["INT-AC"]),

    # 7.x
    ("7.1- Issue Final Project", "PDI",  ["IDP-AC"]),
    ("7.2- Raise Acceptance",    "OT",   ["ARQ-AC", "OT-AC"]),
    ("7.3- Provisional Accept",  "PAC",  ["PAC-AC"]),
    ("7.4- Final Accept",        "FAC",  ["FAC-AC"]),
    ]


    header = pd.Series(df_raw.iloc[6]).astype(str).str.strip()
    name_to_idx = {str(v).strip().upper(): i for i, v in enumerate(header)}

    out: list[tuple[str,str,int]] = []
    for full, short, keys in desired:
        idx = None
        for k in keys:
            idx = name_to_idx.get(str(k).strip().upper())
            if idx is not None:
                break
        if idx is not None:
            out.append((full, short, idx))
    return out

# ============== Base por fase (concluídos/faltantes) + Year (BR) ==============
def sites_for_phase_explicit(
    df_raw: pd.DataFrame, df_clean: pd.DataFrame, fase_full: str,
    site_col_letter: str = "E", uf_col_letter: str = "F", year_col_letter: str = "BR"
) -> pd.DataFrame:
    phase_map = {f: idx for (f, _short, idx) in get_explicit_phase_map(df_raw)}

    def _is_all_label(s) -> bool:
        return isinstance(s, str) and ("todas" in s.lower())

    all_label = _is_all_label(fase_full)
    if (not all_label) and (fase_full not in phase_map):
        return pd.DataFrame()

    def _find_idx_by_name(names: list[str]) -> int | None:
        header_rows = [6, 5, 7]
        for ridx in header_rows:
            if 0 <= ridx < len(df_raw):
                header = (
                    pd.Series(df_raw.iloc[ridx])
                    .astype(str)
                    .str.replace(' ', ' ', regex=False)
                    .str.strip()
                    .str.upper()
                )
                for nm in names:
                    nm_u = str(nm).strip().upper()
                    hit = header[header == nm_u]
                    if not hit.empty:
                        return int(hit.index[0])
        return None

    def _one_phase(full: str) -> pd.DataFrame:
        act_idx = phase_map[full]
        # preferir por nome ao invés de letra
        site_idx = _find_idx_by_name(["SITE", "SITE NAME"])
        uf_idx   = _find_idx_by_name(["STATE", "UF"])
        year_idx = _find_idx_by_name(["YEAR"]) if _find_idx_by_name(["YEAR"]) is not None else None
        carimbo_idx = _find_idx_by_name(["CARIMBO"])

        if site_idx is None:
            site_idx = excel_col_to_idx(site_col_letter)
        if uf_idx is None:
            uf_idx = excel_col_to_idx(uf_col_letter)
        if year_idx is None:
            try:
                year_idx = excel_col_to_idx(year_col_letter)
            except Exception:
                year_idx = None
        if carimbo_idx is None:
            try:
                carimbo_idx = excel_col_to_idx("CZ")
            except Exception:
                carimbo_idx = None
        # valida se os índices existem no DataFrame
        ncols = df_raw.shape[1]
        if not (isinstance(year_idx, int) and 0 <= year_idx < ncols):
            year_idx = None
        if not (isinstance(carimbo_idx, int) and 0 <= carimbo_idx < ncols):
            carimbo_idx = None

        sites = pd.Series(df_raw.iloc[7:, site_idx])
        uf    = pd.Series(df_raw.iloc[7:, uf_idx])
        act   = pd.Series(df_raw.iloc[7:, act_idx])
        yearv = pd.Series(df_raw.iloc[7:, year_idx]) if year_idx is not None else pd.Series([pd.NA]*len(act))
        carim = pd.Series(df_raw.iloc[7:, carimbo_idx]) if carimbo_idx is not None else pd.Series([pd.NA]*len(act))

        base = pd.DataFrame({"SITE": sites, "UF": uf, "actual_date": act, "year": yearv, "Carimbo": carim})
        base = base.dropna(subset=["SITE"]).reset_index(drop=True)
        base["concluded"] = base["actual_date"].notna()

        # normaliza ano (se vier vazio vira NaN)
        base["year"] = pd.to_numeric(base["year"], errors="coerce").astype("Int64")

        # junta com atributos do df_clean (quando existirem) — desduplicado por SITE
        if df_clean is not None and "SITE" in df_clean.columns:
            dfu = df_clean.drop_duplicates(subset=["SITE"])  # evita 1:N
            cols_keep = [c for c in dfu.columns if c in {
                "SITE","state","current_status","Group","Subcon","Type","Qty",
                "Model","SoW","SoW Type","Infra PO"
            }]
            base = base.merge(dfu[cols_keep], on="SITE", how="left")

        # renomear "Infra PO" para "PO" se existir
        if "Infra PO" in base.columns and "PO" not in base.columns:
            base = base.rename(columns={"Infra PO": "PO"})

        # rótulo da fase
        base["fase_label"] = full

        return base.reset_index(drop=True)

    # monta saída para uma fase específica ou para todas
    if all_label:
        frames = [_one_phase(full) for (full,_s,_c) in get_explicit_phase_map(df_raw)]
        out = pd.concat(frames, ignore_index=True)
    else:
        out = _one_phase(fase_full)
    return out

# ============== Snapshot e delay ==============
def _actuals_wide(df_raw: pd.DataFrame):
    phase = get_explicit_phase_map(df_raw)
    site_idx = excel_col_to_idx("E")
    # tente localizar por nome
    header = pd.Series(df_raw.iloc[6]).astype(str).str.strip().str.upper()
    site_hit = header[header.isin(["SITE", "SITE NAME"])]
    if not site_hit.empty:
        site_idx = int(site_hit.index[0])
    sites = pd.Series(df_raw.iloc[7:, site_idx]).rename("SITE")
    wide = pd.DataFrame({"SITE": sites})
    for _full, short, idx in phase:
        s = pd.Series(df_raw.iloc[7:, idx]).apply(_to_excel_datetime)
        wide[short] = s.values
    wide = wide.dropna(subset=["SITE"]).reset_index(drop=True)
    return sites.to_frame(), wide

def last_status_snapshot(df_raw: pd.DataFrame) -> pd.DataFrame:
    _, wide = _actuals_wide(df_raw)
    order = [s for (_f,s,_c) in get_explicit_phase_map(df_raw)]

    def _row_last(sr: pd.Series):
        s = sr[order]
        last_short = s.last_valid_index()
        if last_short is None:
            return pd.Series({"last_phase_short": None, "last_date": pd.NaT})
        return pd.Series({"last_phase_short": last_short, "last_date": s[last_short]})

    res = wide.apply(_row_last, axis=1)
    res["last_date"] = res["last_date"].dt.date  # <- só a data, sem horário
    full_map = {s:f for (f,s,_c) in get_explicit_phase_map(df_raw)}
    res["last_phase_full"] = res["last_phase_short"].map(full_map)
    return pd.concat([wide[["SITE"]], res], axis=1)


def last_delay_days(df_raw: pd.DataFrame, today: datetime | None = None) -> pd.DataFrame:
    if today is None:
        today = pd.Timestamp(datetime.utcnow().date())

    _, wide = _actuals_wide(df_raw)
    order = [s for (_f,s,_c) in get_explicit_phase_map(df_raw)]
    prev  = {order[i]: (order[i-1] if i>0 else None) for i in range(len(order))}
    out = pd.DataFrame({"SITE": wide["SITE"]})

    for s in order:
        cur = pd.to_datetime(wide[s], errors="coerce")
        ps  = prev[s]
        if ps:
            pr = pd.to_datetime(wide[ps], errors="coerce")
            diff = (cur - pr).dt.days
            need_fill = diff.isna() & pr.notna()
            diff.loc[need_fill] = (today - pr.loc[need_fill]).dt.days
        else:
            diff = (today - cur).dt.days
        out[f"delay_{s}"] = diff

    snap = last_status_snapshot(df_raw)[["SITE","last_phase_short"]]
    merged = out.merge(snap, on="SITE", how="left")

    def _pick(row):
        s = row["last_phase_short"]
        return row.get(f"delay_{s}", pd.NA) if pd.notna(s) else pd.NA

    merged["delay_days"] = merged.apply(_pick, axis=1)
    # garante não-negativo
    merged["delay_days"] = pd.to_numeric(merged["delay_days"], errors="coerce").clip(lower=0)
    return merged[["SITE","delay_days"]]


def stage_stay_days(df_raw: pd.DataFrame, today: datetime | None = None) -> pd.DataFrame:
    """Compute, for each site, how many days it stayed in each status.

    For a status S, we define stay_S = next(S)_date - S_date.
    If next(S)_date is missing and S_date exists, use (today - S_date).
    For the last status, use (today - last_status_date).
    Negative values are clipped to zero.
    """
    if today is None:
        today = pd.Timestamp(datetime.utcnow().date())

    _, wide = _actuals_wide(df_raw)
    order = [s for (_f, s, _c) in get_explicit_phase_map(df_raw)]
    out = pd.DataFrame({"SITE": wide["SITE"]})

    for i, s in enumerate(order):
        cur = pd.to_datetime(wide[s], errors="coerce")
        if i < len(order) - 1:
            nxt = pd.to_datetime(wide[order[i + 1]], errors="coerce")
            diff = (nxt - cur).dt.days
            need_fill = diff.isna() & cur.notna()
            diff.loc[need_fill] = (today - cur.loc[need_fill]).dt.days
        else:
            diff = (today - cur).dt.days
        out[f"stay_{s}"] = pd.to_numeric(diff, errors="coerce").clip(lower=0)

    return out