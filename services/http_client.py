import time
import requests

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
}

def obter_html_empresa(session: requests.Session, ticker: str) -> str | None:
    """Realiza a requisição HTTP para a página da empresa no Investidor 10."""
    ticker_clean = ticker.strip().lower()
    url = f"https://investidor10.com.br/acoes/{ticker_clean}/"
    
    try:
        response = session.get(url, timeout=10)
        response.raise_for_status()
        return response.text

    except requests.exceptions.HTTPError as http_err:
        if response.status_code == 404:
            print(f"[ERRO 404] Ticker não encontrado: {ticker.upper()}")
        elif response.status_code == 429:
            print(f"[Aviso 429] Excesso de requisições. Aguardando 10 segundos...")
            time.sleep(10)
        else:
            print(f"[ERRO HTTP] {http_err}")

    except requests.exceptions.RequestException as err:
        print(f"[FALHA CONEXÃO] Erro ao acessar {ticker.upper()}: {err}")
        
    return None