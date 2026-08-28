import time
from services.http_client import get_data_indicadores, get_data_earning_yield
from utils.empresas import empresas, empresas_teste
from utils.excell import salvar_em_excel_por_abas


def main():
    resultados = {}

    for empresa in empresas_teste:
        ticker = empresa[0]
        nome = empresa[1]
        api_indicadores_code = empresa[2]

        print(f"Buscando dados de {nome} ({ticker})...")

        # 1. Busca Earning Yield (Status Invest)
        df_ey = get_data_earning_yield(ticker)

        # 2. Busca Indicadores Históricos (Investidor10)
        df_indicadores = get_data_indicadores(ticker, api_indicadores_code)

        if df_ey is not None and not df_ey.empty:
            # Armazena ambos em um dicionário estruturado para a exportação
            resultados[ticker] = {
                "ey": df_ey,
                "indicadores": df_indicadores,  # Pode ser None se falhar
            }
        else:
            print(f"Aviso: Não foi possível obter dados de EY para {ticker}.")

        # Intervalo para evitar rate limit nas APIs
        time.sleep(2)

    # Exportação para o Excel
    salvar_em_excel_por_abas(resultados, "outputs/indicadores_bolsa.xlsx")

if __name__ == "__main__":
    main()