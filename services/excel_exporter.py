import os
import pandas as pd
from typing import Dict

def salvar_em_excel_por_abas(resultados: Dict[str, pd.DataFrame], caminho_arquivo: str = "outputs/indicadores_empresas.xlsx") -> None:
    """
    Recebe um dicionário {ticker: dataframe} e gera um arquivo Excel 
    onde cada chave (ticker) vira uma aba (sheet) individual.
    """
    if not resultados:
        print("Aviso: Nenhum dado válido encontrado para exportar.")
        return

    # Garante que a pasta de destino exista
    pasta_destino = os.path.dirname(caminho_arquivo)
    if pasta_destino and not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    print(f"\nGerando arquivo Excel em: {caminho_arquivo}...")

    # Utiliza ExcelWriter para gerenciar múltiplas abas
    with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:
        for ticker, df in resultados.items():
            if df is not None and not df.empty:
                # Limite de 31 caracteres exigido pelo Excel para nomes de abas
                nome_aba = ticker[:31]
                
                # Salva o DataFrame na aba correspondente ao ticker
                print(df.loc["LPA"])
                df.to_excel(writer, sheet_name=nome_aba, index=True)

    print(f"Planilha exportada com sucesso! Total de abas salvas: {len(resultados)}")