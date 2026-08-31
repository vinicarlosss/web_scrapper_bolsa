import os
import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


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


def _adicionar_graficos_individuais(
    worksheet, start_row: int, df_indicadores: pd.DataFrame, num_colunas: int
):
    """Cria um gráfico de linha individual para cada indicador e os posiciona

    todos em uma única linha horizontal, começando da Coluna A.
    """
    categories = Reference(
        worksheet,
        min_col=2,
        min_row=start_row,
        max_col=num_colunas + 1,
        max_row=start_row,
    )

    # Quantas colunas do Excel cada gráfico ocupa de largura na horizontal
    LARGURA_EM_COLUNAS = 10

    # Linha fixa para TODOS os gráficos (logo abaixo da tabela de indicadores)
    linha_excel = start_row + len(df_indicadores) + 3

    for idx, nome_indicador in enumerate(df_indicadores.index):
        row_num = start_row + 1 + idx

        chart = LineChart()
        chart.title = f"Evolução: {nome_indicador}"
        chart.style = 13
        chart.legend = None

        # Configuração de Eixos
        chart.x_axis.title = "Período"
        chart.x_axis.delete = False

        chart.y_axis.title = "Valor"
        chart.y_axis.delete = False

        # Dados da linha do indicador atual
        data = Reference(
            worksheet,
            min_col=1,
            min_row=row_num,
            max_col=num_colunas + 1,
            max_row=row_num,
        )

        chart.add_data(data, titles_from_data=True, from_rows=True)
        chart.set_categories(categories)

        # Tamanho de cada gráfico
        chart.height = 8.5
        chart.width = 16.5

        # --- CÁLCULO DA POSIÇÃO HORIZONTAL ---
        # Começa na Coluna 1 (Coluna A) e avança LARGURA_EM_COLUNAS para cada novo gráfico
        coluna_excel_num = 1 + (idx * LARGURA_EM_COLUNAS)
        letra_coluna = get_column_letter(coluna_excel_num)

        posicao_grafico = f"{letra_coluna}{linha_excel}"

        worksheet.add_chart(chart, posicao_grafico)

def salvar_em_excel_por_abas(
    resultados: dict[str, dict], caminho_arquivo: str
):
    """Gera o arquivo Excel com a aba Ranking, aba Priorização (Matriz de Decisão),

    abas individuais por empresa e gráficos dedicados para cada indicador.
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
    colunas_desejadas = ["2021", "2022", "2023", "2024", "2025", "Atual"]

    # Lista de tickers para a aba de Priorização
    tickers_lista = list(resultados.keys())

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

        # --- SEGUNDA ABA: PRIORIZAÇÃO (MATRIZ DE DECISÃO) ---
        _gerar_aba_priorizacao(writer, tickers_lista)

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

            # B) Tabela de Indicadores e Gráficos Individuais
            if df_indicadores is not None and not df_indicadores.empty:
                df_ind_filtrado = df_indicadores.copy()

                cols_presentes = [
                    c for c in colunas_desejadas if c in df_ind_filtrado.columns
                ]
                if cols_presentes:
                    df_ind_filtrado = df_ind_filtrado[cols_presentes]

                start_row_ind = len(df_ey) + 4

                ws_empresa.cell(
                    row=start_row_ind - 1,
                    column=1,
                    value="Indicadores Fundamentalistas (2021 - Atual)",
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

                # C) Gerar os Gráficos Individuais por Indicador
                _adicionar_graficos_individuais(
                    worksheet=ws_empresa,
                    start_row=start_row_ind + 1,
                    df_indicadores=df_ind_filtrado,
                    num_colunas=len(cols_presentes),
                )

            _ajustar_largura_colunas(ws_empresa)

    print(
        f"\n[SUCESSO] Arquivo Excel exportado com sucesso para: '{caminho_arquivo}'"
    )


def _gerar_aba_priorizacao(writer, tickers: list[str]):
    """Cria a aba 'Priorizacao' com perguntas de pontuação, dropdowns de Sim/Não

    e cálculo automático de pontos para priorização de compras.
    """
    wb = writer.book
    ws = wb.create_sheet(title="Priorizacao")

    # 1. Título e Tabela Explicativa das Regras
    ws["A1"] = "MATRIZ DE PRIORIZAÇÃO DE COMPRA"
    ws["A1"].font = openpyxl.styles.Font(bold=True, size=14)

    ws["A3"] = "Critério de Pontuação"
    ws["B3"] = "Pontos se Sim"
    ws["C3"] = "Pontos se Não"

    regras = [
        ("1 - Está em recuperação judicial?", -1, 1),
        ("2 - Lucro por ação crescente?", 1, -1),
        ("3 - Despesa operacional crescente?", -1, 1),
        ("4 - Distribui dividendos?", 1, -1),
        ("5 - Está no ranking de 20 ações mais baratas da bolsa?", 1, -1),
    ]

    for idx, (pergunta, p_sim, p_nao) in enumerate(regras, start=4):
        ws.cell(row=idx, column=1, value=pergunta)
        ws.cell(row=idx, column=2, value=p_sim)
        ws.cell(row=idx, column=3, value=p_nao)

    # Estiliza cabeçalho das regras
    for col in range(1, 4):
        ws.cell(row=3, column=col).font = openpyxl.styles.Font(
            bold=True, color="FFFFFF"
        )
        ws.cell(row=3, column=col).fill = openpyxl.styles.PatternFill(
            start_color="1F497D", end_color="1F497D", fill_type="solid"
        )

    # 2. Tabela de Avaliação das Ações
    start_row_tabela = 11

    headers_tabela = [
        "Ticker",
        "Q1: Rec. Judicial? (-1)",
        "Q2: LPA Crescente? (+1)",
        "Q3: Despesa Crescente? (-1)",
        "Q4: Dividendos? (+1)",
        "Q5: Top 20 Baratas? (+1)",
        "Pontuação Total",
    ]

    for col_idx, header in enumerate(headers_tabela, start=1):
        cell = ws.cell(row=start_row_tabela, column=col_idx, value=header)
        cell.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        cell.fill = openpyxl.styles.PatternFill(
            start_color="366092", end_color="366092", fill_type="solid"
        )

    # Criação do Dropdown de seleção (Sim / Não)
    dv = DataValidation(
        type="list", formula1='"Sim,Não"', allow_blank=True, showDropDown=True
    )
    ws.add_data_validation(dv)

    linha_atual = start_row_tabela + 1

    for ticker in tickers:
        ws.cell(row=linha_atual, column=1, value=ticker).font = (
            openpyxl.styles.Font(bold=True)
        )

        # Preenche dropdown de Sim/Não padrão para Q1 a Q5 (Colunas B até F)
        for col_idx in range(2, 7):
            c_cell = ws.cell(row=linha_atual, column=col_idx, value="Não")
            dv.add(c_cell)

        # Fórmula Excel para calcular a pontuação dinâmica conforme as escolhas
        # Q1: Sim = -1
        # Q2: Sim = +1
        # Q3: Sim = -1
        # Q4: Sim = +1
        # Q5: Sim = +1
        formula = (
            f'=IF(B{linha_atual}="Sim", -1, 0) + '
            f'IF(C{linha_atual}="Sim", 1, 0) + '
            f'IF(D{linha_atual}="Sim", -1, 0) + '
            f'IF(E{linha_atual}="Sim", 1, 0) + '
            f'IF(F{linha_atual}="Sim", 1, 0)'
        )

        cell_total = ws.cell(row=linha_atual, column=7, value=formula)
        cell_total.font = openpyxl.styles.Font(bold=True)
        cell_total.alignment = openpyxl.styles.Alignment(horizontal="center")

        linha_atual += 1

    # Adiciona Autofiltro na tabela para você poder ordenar por pontuação
    max_linha = linha_atual - 1
    ws.auto_filter.ref = f"A{start_row_tabela}:G{max_linha}"

    _ajustar_largura_colunas(ws)