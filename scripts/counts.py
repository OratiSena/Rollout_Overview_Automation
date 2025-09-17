from pathlib import Path
import pandas as pd
from core.etl_rollout import get_explicit_phase_map, last_status_snapshot

path = Path("ZTE Rollout Framework - Claro RAN Sites - Live (25).xlsb")
df_raw = pd.read_excel(path, sheet_name=0, header=None, engine="pyxlsb")
ph = get_explicit_phase_map(df_raw)
order_short = [s for (_f,s,_i) in ph]
short2idx = {s:i for i,s in enumerate(order_short)}

snap = last_status_snapshot(df_raw)
snap = snap.rename(columns={"last_phase_short":"last_short"})
snap["last_idx"] = snap["last_short"].map(short2idx)
snap["current_idx"] = snap["last_idx"].fillna(-1).astype(int) + 1
snap["current_idx"] = snap["current_idx"].clip(0, len(order_short)-1)
snap["current_short"] = snap["current_idx"].map(lambda i: order_short[i])

print("total sites:", snap["SITE"].nunique())
print("faltando dist:")
print(snap["current_short"].value_counts().sort_index())
print("concl counts INT phase (cumulative >=):")
for s in order_short:
    idx = short2idx[s]
    concl = int((snap["last_idx"].fillna(-1) >= idx).sum())
    print(s, concl)
