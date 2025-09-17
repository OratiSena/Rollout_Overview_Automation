import json
from pathlib import Path

import pandas as pd


def analyze(path: Path):
    ext = path.suffix.lower()
    engine = "pyxlsb" if ext == ".xlsb" else None
    try:
        df = pd.read_excel(path, sheet_name=0, header=None, nrows=12, engine=engine)
    except Exception as e:
        return {"path": str(path), "error": str(e)}

    header = pd.Series(df.iloc[6]).astype(str).str.strip()
    name_to_idx = {str(v).strip().upper(): int(i) for i, v in enumerate(header)}

    def find(names):
        for nm in names:
            u = str(nm).strip().upper()
            if u in name_to_idx:
                return name_to_idx[u]
        return None

    site_idx = find(["SITE", "SITE NAME"])
    uf_idx = find(["UF", "STATE"])
    year_idx = find(["YEAR"])
    carimbo_idx = find(["CARIMBO"])
    po_idx = find(["PO", "INFRA PO"])
    reg_idx = find(["REGIONAL", "GROUP"])

    ac_cols = [
        (int(i), str(v))
        for i, v in enumerate(header)
        if isinstance(v, str) and v.strip().upper().endswith("-AC")
    ]

    return {
        "path": path.name,
        "site_idx": site_idx,
        "uf_idx": uf_idx,
        "year_idx": year_idx,
        "carimbo_idx": carimbo_idx,
        "po_idx": po_idx,
        "reg_idx": reg_idx,
        "ac_count": len(ac_cols),
        "ac_examples": ac_cols[:12],
        "header_sample": [str(header[i]) for i in range(min(len(header), 50))],
    }


def main():
    files = [
        Path("ZTE Rollout Framework - Claro RAN Sites - Live (25).xlsb"),
        Path("Rollout Novo.xlsx"),
    ]
    out = {}
    for f in files:
        if f.exists():
            out[f.name] = analyze(f)
    print(json.dumps(out, indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()

