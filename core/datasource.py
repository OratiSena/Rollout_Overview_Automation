import pandas as pd
import streamlit as st
from pathlib import Path

@st.cache_data
def load_rollout_dataframe(path: str, sheet_name=None) -> pd.DataFrame:
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
    # .xlsb -> pyxlsb
    df = pd.read_excel(file_path, sheet_name=sheet_name or 0, engine="pyxlsb")
    return df