import os
import pandas as pd

def _aplicar_formatacao(
    worksheet, df: pd.DataFrame, formato_moeda: str, formato_porcentagem: str
):
    """Aplica formatação numérica de Moeda, Porcentagem e ajusta a largura das colunas."""
    header_colunas = [df.index.name or "ticker"] + list(df.columns)

    for col_idx, col_name in enumerate(header_colunas, start=1):
        if col_name in ["valor_de_mercado", "divida_liquida", "ebit"]:
            for row_idx in range(2, len(df) + 2):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.number_format = formato_moeda

        elif col_name == "earning_yield":
            for row_idx in range(2, len(df) + 2):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.number_format = formato_porcentagem

    for col in worksheet.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max(max_len + 5, 18)

def salvar_em_excel_por_abas(
    resultados: dict[str, pd.DataFrame], caminho_arquivo: str
):
    """Cria um arquivo Excel salvando a primeira aba ('Ranking') com todas as empresas

    ordenadas do menor para o maior Earning Yield, seguida pelas abas individuais.
    """
    if not resultados:
        print("Nenhum dado para salvar no arquivo Excel.")
        return

    diretorio = os.path.dirname(caminho_arquivo)
    if diretorio:
        os.makedirs(diretorio, exist_ok=True)

    # 1. Calcula o Earning Yield para cada DataFrame no dicionário
    for ticker, df in resultados.items():
        ey = (
            df["ebit"] / (df["divida_liquida"] + df["valor_de_mercado"])
        ).values[0]
        df["earning_yield"] = ey

    # 2. Junta todas as empresas em um único DataFrame para montar o Ranking
    df_ranking = pd.concat(resultados.values())

    # 3. Ordena do menor para o maior Earning Yield
    df_ranking.sort_values(by="earning_yield", ascending=False, inplace=True)

    # Formatadores do Excel
    formato_moeda = 'R$ #,##0.00;[Red]-R$ #,##0.00;"-"'
    formato_porcentagem = "0.00%"

    with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:
        # --- PRIMEIRA ABA: RANKING ---
        df_ranking.to_excel(writer, sheet_name="Ranking", index=True)
        _aplicar_formatacao(
            writer.sheets["Ranking"],
            df_ranking,
            formato_moeda,
            formato_porcentagem,
        )

        # --- ABAS INDIVIDUAIS POR EMPRESA ---
        for ticker, df in resultados.items():
            nome_aba = str(ticker)[:31]
            df.to_excel(writer, sheet_name=nome_aba, index=True)
            _aplicar_formatacao(
                writer.sheets[nome_aba],
                df,
                formato_moeda,
                formato_porcentagem,
            )

    print(
        f"\n[SUCESSO] Arquivo Excel exportado com sucesso para: '{caminho_arquivo}'"
    )
    