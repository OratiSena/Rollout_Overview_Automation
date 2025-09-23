"""ETL para análise de integração."""

import pandas as pd

def process_integration_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Processa os dados de integração para análise.

    Args:
        df (pd.DataFrame): DataFrame original com os dados de integração.

    Returns:
        pd.DataFrame: DataFrame processado com colunas relevantes e status calculados.
    """
    # Garantir que as colunas necessárias existam
    required_columns = [
        "Site Name", "Integration date", "MOS", "General Status", "4G Status", "2G Status"
    ]
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Coluna obrigatória ausente: {col}")

    # Filtrar colunas relevantes
    df = df[required_columns]

    # Converter datas para datetime
    df["Integration date"] = pd.to_datetime(df["Integration date"], errors="coerce")

    # Adicionar colunas calculadas, se necessário
    df["4G Status"] = df["4G Status"].fillna("Unknown")
    df["2G Status"] = df["2G Status"].fillna("Unknown")

    return df

def summarize_status(df: pd.DataFrame) -> pd.DataFrame:
    """
    Gera um resumo dos status de integração.

    Args:
        df (pd.DataFrame): DataFrame processado.

    Returns:
        pd.DataFrame: Resumo dos status.
    """
    status_summary = df["4G Status"].value_counts().reset_index()
    status_summary.columns = ["Status", "Count"]
    return status_summary