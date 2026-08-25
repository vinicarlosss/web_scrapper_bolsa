import time
from services.http_client import get_data_indicadores, get_data_earning_yield
from utils.empresas import empresas, empresas_teste
from utils.excell import salvar_em_excel_por_abas


def main():
    resultados = {}

    for empresa in empresas:
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