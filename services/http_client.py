import time
import requests
from services.parse.indicadores_parse import parse_historico_indicadores

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
}

def get_data_earning_yeld(ticker: str):
    return None


def get_data_indicadores(ticker: str) -> str | None:
    """Realiza a requisição HTTP para a página da empresa no Investidor 10."""
    session = requests.Session()
    session.headers.update(HEADERS)
    ticker_clean = ticker.strip().lower()
    url_indicadores = "https://investidor10.com.br/api/historico-indicadores/160/5/?v=2"
    indicadores_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Referer": f"https://investidor10.com.br/acoes/{ticker_clean}/",
    "X-Requested-With": "XMLHttpRequest"
    }

    
    
    try:
        indicadores_response = session.get(url_indicadores, headers=indicadores_headers, timeout=10)
        indicadores_response.raise_for_status()
        dados_indicadores_brutos = indicadores_response.json()
        # Retornando em formato de DataFrame
        df_indicadores = parse_historico_indicadores(dados_indicadores_brutos)
        return df_indicadores

    except requests.exceptions.HTTPError as http_err:
        if indicadores_response.status_code == 404:
            print(f"[ERRO 404] Ticker não encontrado: {ticker.upper()}")
        elif indicadores_response.status_code == 429:
            print(f"[Aviso 429] Excesso de requisições. Aguardando 10 segundos...")
            time.sleep(10)
        else:
            print(f"[ERRO HTTP] {http_err}")

    except requests.exceptions.RequestException as err:
        print(f"[FALHA CONEXÃO] Erro ao acessar {ticker.upper()}: {err}")
        
    return None