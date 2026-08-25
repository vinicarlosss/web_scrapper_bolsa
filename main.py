import time
from services.http_client import get_data_indicadores, get_data_earning_yield
from utils.empresas import empresas, empresas_teste
import pandas as pd
import os
from utils.excell import salvar_em_excel_por_abas


def calcular_earning_yield(df: pd.DataFrame) -> float | None:
    """Calcula o Earning Yield: EBIT / (Dívida Líquida + Valor de Mercado)"""
    try:
        ebit = df["ebit"].values[0]
        divida_liquida = df["divida_liquida"].values[0]
        valor_mercado = df["valor_de_mercado"].values[0]

        # Trata valores ausentes ou nulos do Pandas (pd.NA / None)
        if pd.isna(ebit) or pd.isna(divida_liquida) or pd.isna(valor_mercado):
            return None

        ev = float(divida_liquida) + float(valor_mercado)

        if ev == 0:
            return None

        return float(ebit) / ev
    except (IndexError, KeyError, ValueError, TypeError):
        return None


def main():
    resultados = {}

    for empresa in empresas_teste:
        ticker = empresa[0]
        nome = empresa[1]

        print(f"Buscando dados de {nome} ({ticker})...")
        data = get_data_earning_yield(ticker)

        if data is not None and not data.empty:
            resultados[ticker] = data
        else:
            print(f"Aviso: Não foi possível obter dados para {ticker}.")

        # Intervalo para evitar rate limit na API
        time.sleep(2)

    # Chamada do módulo de exportação isolado
    salvar_em_excel_por_abas(resultados, "outputs/indicadores_bolsa.xlsx")


if __name__ == "__main__":
    main()