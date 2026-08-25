import time
import requests
from services.parse.indicadores_parse import parse_historico_indicadores
from bs4 import BeautifulSoup
import pandas as pd
import re

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
}

def _limpar_valor_status_invest(texto: str) -> float | None:
    """Limpa a string formatada do Status Invest e converte para float numérico."""
    if not texto or texto.strip() in ["-", "", "--", "&nbsp;"]:
        return None

    texto = texto.strip()

    multiplicador = 1.0
    if "M" in texto:
        multiplicador = 1_000_000.0
    elif "B" in texto:
        multiplicador = 1_000_000_000.0
    elif "K" in texto:
        multiplicador = 1_000.0

    limpo = re.sub(r"[^\d,-]", "", texto)

    if "," in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    else:
        limpo = limpo.replace(".", "")

    try:
        return float(limpo) * multiplicador
    except ValueError:
        return None


def _get_ebit_from_api(ticker: str) -> float | None:
    """Consulta a API interna de DRE do Status Invest e retorna o EBIT dos últimos 12 meses."""
    url_dre = (
        f"https://statusinvest.com.br/acao/getdre?code={ticker.lower()}&type=0"
    )

    try:
        response = requests.get(url_dre, headers=HEADERS, timeout=10)
        if response.status_code != 200:
            return None

        res_json = response.json()
        api_data = res_json.get("data", {})
        grid = api_data.get("grid", [])

        for row in grid:
            grid_line = row.get("gridLineModel", {})
            title = grid_line.get("name", "").upper() or grid_line.get(
                "key", ""
            ).upper()

            # Valida a palavra exata EBIT para evitar pegar EBITDA
            if re.search(r"\bEBIT\b", title):
                # O array de valores fica dentro de gridLineModel
                values = grid_line.get("values", [])
                if values:
                    # O primeiro elemento refere-se aos Últimos 12 Meses (Últ. 12M)
                    primeiro_valor = values[0]

                    # Se for um dicionário (ex: {'value': 12345.0, ...})
                    if isinstance(primeiro_valor, dict):
                        val = primeiro_valor.get("value")
                        if val is not None:
                            return float(val)

                        val_str = primeiro_valor.get("valueFormat")
                        return _limpar_valor_status_invest(val_str)

                    # Se o array já trouxer os números diretamente (ex: [12345.0, 9876.0])
                    elif isinstance(primeiro_valor, (int, float)):
                        return float(primeiro_valor)

                    # Caso venha como string
                    elif isinstance(primeiro_valor, str):
                        return _limpar_valor_status_invest(primeiro_valor)

    except Exception as e:
        print(f"[{ticker}] Erro ao buscar EBIT via API interna: {e}")

    return None

def get_data_earning_yield(
    ticker: str, default_divida_banco: float = 0.0
) -> pd.DataFrame:
    """Acessa o Status Invest e retorna um DataFrame do Pandas com os dados do ticker."""
    url_pagina = f"https://statusinvest.com.br/acoes/{ticker.lower()}"

    dados = {
        "ticker": [ticker.upper()],
        "valor_de_mercado": [pd.NA],
        "divida_liquida": [pd.NA],
        "ebit": [pd.NA],
    }

    try:
        # 1. Requisição da página HTML para Valor de Mercado e Dívida Líquida
        response = requests.get(url_pagina, headers=HEADERS, timeout=10)
        if response.status_code != 200:
            print(
                f"[{ticker}] Erro HTTP ao acessar Status Invest: {response.status_code}"
            )
            return pd.DataFrame(dados).set_index("ticker")

        soup = BeautifulSoup(response.text, "html.parser")

        valor_mercado = None
        divida_liquida = None
        encontrou_divida = False

        container_top_info = soup.find("div", class_="info-3")

        if container_top_info:
            blocos = container_top_info.find_all("div", class_="info")
            for bloco in blocos:
                title_tag = bloco.find("h3", class_="title")
                value_tag = bloco.find("strong", class_="value")

                if not title_tag or not value_tag:
                    continue

                titulo = title_tag.get_text(strip=True).upper()
                valor_str = value_tag.get_text(strip=True)

                if titulo == "VALOR DE MERCADO":
                    valor_mercado = _limpar_valor_status_invest(valor_str)

                elif titulo in ["DÍVIDA LÍQUIDA", "DIVIDA LIQUIDA"]:
                    encontrou_divida = True
                    divida_liquida = _limpar_valor_status_invest(valor_str)

        if not encontrou_divida:
            divida_liquida = default_divida_banco

        # 2. Requisição direta à API interna para obter o EBIT
        ebit = _get_ebit_from_api(ticker)

        # Atribuição dos valores finais ao dicionário
        dados["valor_de_mercado"] = [
            valor_mercado if valor_mercado is not None else pd.NA
        ]
        dados["divida_liquida"] = [
            divida_liquida if divida_liquida is not None else pd.NA
        ]
        dados["ebit"] = [ebit if ebit is not None else pd.NA]

    except Exception as e:
        print(f"[{ticker}] Erro durante o scraping no Status Invest: {e}")

    df = pd.DataFrame(dados)
    df.set_index("ticker", inplace=True)
    return df

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