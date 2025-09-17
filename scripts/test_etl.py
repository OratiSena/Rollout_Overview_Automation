from pathlib import Path
import pandas as pd
from core.etl_rollout import get_explicit_phase_map, last_status_snapshot

path = Path("ZTE Rollout Framework - Claro RAN Sites - Live (25).xlsb")
df_raw = pd.read_excel(path, sheet_name=0, header=None, engine="pyxlsb")

ph = get_explicit_phase_map(df_raw)
print("phases:")
for full, short, idx in ph:
    print(idx, short, full)

snap = last_status_snapshot(df_raw).head()
print("snapshot sample:")
print(snap)
