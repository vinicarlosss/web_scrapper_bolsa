import os
import pandas as pd


def _aplicar_formatacao(
    worksheet,
    df: pd.DataFrame,
    formato_moeda: str,
    formato_porcentagem: str,
    start_row: int = 1,
):
    """Aplica formatação numérica de Moeda, Porcentagem e ajusta a largura das colunas."""
    indicadores_pct = [
        "earning_yield",
        "payout",
        "margem bruta",
        "margem líquida",
        "margem_bruta",
        "margem_liquida",
    ]

    is_index_indicadores = df.index.name in ["Indicador", "indicador"]

    if is_index_indicadores:
        for row_idx, idx_val in enumerate(df.index, start=start_row + 1):
            nome_indicador = str(idx_val).strip().lower()

            if (
                nome_indicador in indicadores_pct
                or "margem" in nome_indicador
                or "payout" in nome_indicador
            ):
                for col_idx in range(2, len(df.columns) + 2):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    if isinstance(cell.value, (int, float)):
                        if cell.value > 1.0 or cell.value < -1.0:
                            cell.value = cell.value / 100.0
                        cell.number_format = formato_porcentagem
    else:
        header_colunas = [df.index.name or "ticker"] + list(df.columns)

        for col_idx, col_name in enumerate(header_colunas, start=1):
            nome_coluna = str(col_name).strip().lower()

            if nome_coluna in ["valor_de_mercado", "divida_liquida", "ebit"]:
                for row_idx in range(start_row + 1, start_row + len(df) + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = formato_moeda

            elif (
                nome_coluna in indicadores_pct
                or "margem" in nome_coluna
                or "payout" in nome_coluna
            ):
                for row_idx in range(start_row + 1, start_row + len(df) + 1):
                    cell = worksheet.cell(row=row_idx, column=col_idx)
                    if isinstance(cell.value, (int, float)):
                        if cell.value > 1.0 or cell.value < -1.0:
                            cell.value = cell.value / 100.0
                        cell.number_format = formato_porcentagem


def _ajustar_largura_colunas(worksheet):
    """Ajusta a largura de todas as colunas na aba de acordo com o maior conteúdo."""
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = col[0].column_letter
        worksheet.column_dimensions[col_letter].width = max(max_len + 4, 15)


def salvar_em_excel_por_abas(
    resultados: dict[str, dict], caminho_arquivo: str
):
    """Gera o arquivo Excel salvando a aba Ranking e as abas individuais

    incluindo a coluna 'Atual' e os anos de 2021 a 2025.
    """
    if not resultados:
        print("Nenhum dado para salvar no arquivo Excel.")
        return

    diretorio = os.path.dirname(caminho_arquivo)
    if diretorio:
        os.makedirs(diretorio, exist_ok=True)

    dfs_ranking = []

    # 1. Processa o Earning Yield para cada empresa
    for ticker, dados in resultados.items():
        df_ey = dados["ey"].copy()

        divida = df_ey["divida_liquida"].fillna(0).values[0]
        vme = df_ey["valor_de_mercado"].values[0]
        ebit = df_ey["ebit"].values[0]

        ey = ebit / (divida + vme)
        df_ey["earning_yield"] = ey

        dfs_ranking.append(df_ey)

    # 2. Concatena e ordena o Ranking (maior para o menor EY)
    df_ranking = pd.concat(dfs_ranking)
    df_ranking.sort_values(by="earning_yield", ascending=False, inplace=True)

    formato_moeda = 'R$ #,##0.00;[Red]-R$ #,##0.00;"-"'
    formato_porcentagem = "0.00%"

    # Incluído 'Atual' junto dos anos de 2021 a 2025
    colunas_desejadas = ["Atual", "2025", "2024", "2023", "2022", "2021"]

    with pd.ExcelWriter(caminho_arquivo, engine="openpyxl") as writer:
        # --- PRIMEIRA ABA: RANKING ---
        df_ranking.to_excel(writer, sheet_name="Ranking", index=True)
        ws_ranking = writer.sheets["Ranking"]
        _aplicar_formatacao(
            ws_ranking,
            df_ranking,
            formato_moeda,
            formato_porcentagem,
            start_row=1,
        )
        _ajustar_largura_colunas(ws_ranking)

        # --- ABAS INDIVIDUAIS POR EMPRESA ---
        for ticker, dados in resultados.items():
            nome_aba = str(ticker)[:31]
            df_ey = dados["ey"]
            df_indicadores = dados.get("indicadores")

            # A) Tabela de Earning Yield (Topo)
            df_ey.to_excel(writer, sheet_name=nome_aba, startrow=0, index=True)
            ws_empresa = writer.sheets[nome_aba]
            _aplicar_formatacao(
                ws_empresa,
                df_ey,
                formato_moeda,
                formato_porcentagem,
                start_row=1,
            )

            # B) Tabela de Indicadores (Abaixo)
            if df_indicadores is not None and not df_indicadores.empty:
                df_ind_filtrado = df_indicadores.copy()

                # Filtra apenas as colunas que correspondem a "Atual" e anos 2021-2025
                cols_presentes = [
                    c for c in colunas_desejadas if c in df_ind_filtrado.columns
                ]
                if cols_presentes:
                    df_ind_filtrado = df_ind_filtrado[cols_presentes]

                start_row_ind = len(df_ey) + 4

                ws_empresa.cell(
                    row=start_row_ind - 1,
                    column=1,
                    value="Indicadores Fundamentalistas (Atual e 2021 - 2025)",
                )

                df_ind_filtrado.to_excel(
                    writer,
                    sheet_name=nome_aba,
                    startrow=start_row_ind,
                    index=True,
                )

                _aplicar_formatacao(
                    ws_empresa,
                    df_ind_filtrado,
                    formato_moeda,
                    formato_porcentagem,
                    start_row=start_row_ind + 1,
                )

            _ajustar_largura_colunas(ws_empresa)

    print(
        f"\n[SUCESSO] Arquivo Excel exportado com sucesso para: '{caminho_arquivo}'"
    )
