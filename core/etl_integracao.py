"""ETL para análise de integração."""

import pandas as pd
import warnings

def process_integration_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Processa os dados de integração para análise.

    Args:
        df (pd.DataFrame): DataFrame original com os dados de integração.

    Returns:
        pd.DataFrame: DataFrame processado com todas as colunas relevantes.
    """
    # Definir todas as colunas exatas conforme fornecidas pelo usuário
    expected_columns = [
        "Site Name", "Region", "General Status", "Comment", "4G Status", "2G Status",
        "Alarm test", "Calling test", "IR", "SSV", "ARQ Number", "IW Novo", "IW Reuso",
        "OT 2G", "OT 4G", "OT Date", "OT Status", "Pre-comissioned", "Related BSC",
        "BSC ID", "BSC SCTP", "MEIO TX", "MEID", "2G BTS ID", "LTE eNodeB ID",
        "OAM IP", "OAM IP netmask", "OAM Gateway", "VLAN", "GSM IP", "GSM IP netmask",
        "GSM IP Gateway", "LTE IP", "LTE IP netmask", "LTE IP Gateway"
    ]
    
    # Limpar nomes das colunas (remover espaços extras, normalizar)
    df.columns = df.columns.str.strip()
    
    # Verificar quais colunas estão presentes e quais estão faltando
    present_columns = [col for col in expected_columns if col in df.columns]
    missing_columns = [col for col in expected_columns if col not in df.columns]
    
    if missing_columns:
        warnings.warn(f"Colunas ausentes na planilha: {', '.join(missing_columns)}. Continuando sem elas.")
    
    # Usar apenas as colunas que existem
    df_filtered = df[present_columns].copy()
    
    # Adicionar colunas ausentes como NaN para manter a estrutura esperada
    for col in missing_columns:
        df_filtered[col] = pd.NA
    
    # Reordenar para manter a ordem esperada
    df_filtered = df_filtered.reindex(columns=expected_columns)
    
    # Normalizar valores nas colunas de status (trim whitespace, padronizar valores)
    status_columns = ["4G Status", "2G Status", "Alarm test", "Calling test", "IR", "SSV", 
                     "OT 2G", "OT 4G", "OT Status", "General Status"]
    
    for col in status_columns:
        if col in df_filtered.columns:
            # Converter para string, trim, e padronizar alguns valores
            df_filtered[col] = df_filtered[col].astype(str).str.strip()
            # Substituir valores vazios/nan por "Unknown"
            df_filtered[col] = df_filtered[col].replace(['nan', 'NaN', '', ' ', 'None'], 'Unknown')
            # Padronizar alguns valores comuns
            df_filtered[col] = df_filtered[col].str.replace('Finished', 'Finished', case=False)
            df_filtered[col] = df_filtered[col].str.replace('Pending', 'Pending', case=False)
            df_filtered[col] = df_filtered[col].str.replace('Unknown', 'Unknown', case=False)
    
    # Garantir que Site Name não tenha valores vazios
    if "Site Name" in df_filtered.columns:
        df_filtered = df_filtered[df_filtered["Site Name"].notna()]
        df_filtered = df_filtered[df_filtered["Site Name"].astype(str).str.strip() != ""]
    
    return df_filtered

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