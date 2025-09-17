import pandas as pd
from pathlib import Path

path = Path("ZTE Rollout Framework - Claro RAN Sites - Live (25).xlsb")
df = pd.read_excel(path, sheet_name=0, header=None, engine="pyxlsb")
hdr = df.iloc[6].tolist()
print("cols", len(hdr))
for i, v in enumerate(hdr):
    if pd.isna(v):
        continue
    print(i, str(v))
