"""ETL para análise de integração."""

import pandas as pd

def process_integration_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Processa os dados de integração para análise.

    Args:
        df (pd.DataFrame): DataFrame original com os dados de integração.

    Returns:
        pd.DataFrame: DataFrame processado com todas as colunas relevantes.
    """
    # Garantir que as colunas necessárias existam
    required_columns = [
        "Site Name", "General Status", "Comment", "4G Status", "2G Status",
        "Alarm test", "Calling test", "IR", "SSV", "ARQ Number", "OT 4G", "OT 2G", "OT Status", "Pre-comissioned",
        "Region", "Related BSC", "BSC ID", "BSC SCTP", "MEIO TX", "MEID", "2G BTS ID", "LTE eNodeB ID",
        "OAM IP", "OAM IP netmask", "OAM Gateway", "VLAN", "GSM IP", "GSM IP netmask", "GSM IP Gateway",
        "LTE IP", "LTE IP netmask", "LTE IP Gateway"
    ]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Colunas obrigatórias ausentes: {', '.join(missing_columns)}")

    # Filtrar colunas relevantes
    df = df[required_columns]

    # Nota: não forçar conversão de Integration date/MOS aqui — a página lida com campos opcionais

    # Adicionar colunas calculadas, se necessário
    status_columns = ["4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", "OT 4G", "OT 2G", "OT Status"]
    for c in status_columns:
        if c in df.columns:
            df[c] = df[c].fillna("Unknown")

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