import pandas as pd
from datetime import datetime, timedelta

# ============== Limpeza ==============
def clean_rollout_dataframe(df_raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    header_rows = 7
    df_header = df_raw.iloc[:header_rows, :]
    df_data = df_raw.iloc[header_rows:, :].copy()

    new_header = df_raw.iloc[6].tolist()
    df_data.columns = new_header
    df_data = df_data.loc[:, ~df_data.columns.isna()]
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

Notas: mapeamento de fases derivado dinamicamente via get_explicit_phase_map().
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
_EXPLICIT_PHASES = []

def _excel_idx_to_col(idx: int) -> str:
    idx += 1
    s = ""
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s

def kpi_from_explicit_cells(df_raw: pd.DataFrame, method: str = "header", sites_col_letter: str = "E") -> pd.DataFrame:
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
            raw_val = df_raw.iloc[5, cidx]
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

    Includes phases before 5.1 and handles header name variants for "*-AC".
    """
    desired = [
        # 1.x
        ("1.1- PWS Pre-PO",          "PPO",  ["PPO-AC", "PPWS-AC", "PWS PRE-PO-AC", "PWS PREPO-AC", "PWS PRE PO-AC"]),
        ("1.2- Issue PWS PO",        "IPO",  ["IPO-AC", "IPWS-AC", "PWS PO-AC", "PWSPO-AC"]),
        ("1.3- Site Survey",         "SV",   ["SSV-AC", "SV-AC", "SITE SURVEY-AC", "SURVEY-AC"]),
        ("1.4- Issue Project",       "IPP",  ["IPP-AC"]),
        ("1.5- Approve Project",     "APP",  ["APP-AC"]),

        # 2.x
        ("2.3- Infra Pre-PO",        "PINF", ["PINF-AC", "RFP-AC", "INFRA PRE-PO-AC"]),

        # 3.x, 4.x
        ("3.1- Ready for Install",   "RFI",  ["RFI-AC"]),
        ("4.1- Transmission",        "TXA",  ["TXA-AC", "TX-AC", "TRANSMISSION-AC"]),

        # 5.x
        ("5.1- Material Request",    "MRF",  ["MRQ-AC", "MRF-AC", "MR-AC", "MAT REQUEST-AC"]),
        ("5.2- DOMRF",               "DO",   ["RDO-AC", "DO MRF-AC", "DOMRF-AC"]),
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

# ============== Base por fase + Year (BR) ==============
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
        header = pd.Series(df_raw.iloc[6]).astype(str).str.strip().str.upper()
        for nm in names:
            nm_u = str(nm).strip().upper()
            hit = header[header == nm_u]
            if not hit.empty:
                return int(hit.index[0])
        return None

    def _one_phase(full: str) -> pd.DataFrame:
        act_idx = phase_map[full]
        site_idx = _find_idx_by_name(["SITE", "SITE NAME"]) or excel_col_to_idx(site_col_letter)
        uf_idx   = _find_idx_by_name(["STATE", "UF"]) or excel_col_to_idx(uf_col_letter)
        year_idx = _find_idx_by_name(["YEAR"]) or None
        carimbo_idx = _find_idx_by_name(["CARIMBO"]) or None

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
        base["year"] = pd.to_numeric(base["year"], errors="coerce").astype("Int64")

        if df_clean is not None and "SITE" in df_clean.columns:
            dfu = df_clean.drop_duplicates(subset=["SITE"])
            cols_keep = [c for c in dfu.columns if c in {
                "SITE","state","current_status","Group","Subcon","Type","Qty",
                "Model","SoW","SoW Type","Infra PO"
            }]
            base = base.merge(dfu[cols_keep], on="SITE", how="left")

        if "Infra PO" in base.columns and "PO" not in base.columns:
            base = base.rename(columns={"Infra PO": "PO"})
        base["fase_label"] = full
        return base.reset_index(drop=True)

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
    res["last_date"] = res["last_date"].dt.date
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
    merged["delay_days"] = pd.to_numeric(merged["delay_days"], errors="coerce").clip(lower=0)
    return merged[["SITE","delay_days"]]

def stage_stay_days(df_raw: pd.DataFrame, today: datetime | None = None) -> pd.DataFrame:
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

