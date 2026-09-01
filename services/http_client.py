import time
import requests
from services.parse.indicadores_parse import parse_historico_indicadores
from bs4 import BeautifulSoup
import pandas as pd
import re
import datetime

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
}

def _limpar_valor_status_invest(val) -> float | None:
    if val is None or pd.isna(val) or val == "-":
        return None
    if isinstance(val, (int, float)):
        return float(val)
    try:
        val_str = (
            str(val)
            .replace(".", "")
            .replace(",", ".")
            .replace("%", "")
            .strip()
        )
        return float(val_str)
    except Exception:
        return None


def get_balanco_patrimonial(
    ticker: str, min_year: int = 2021, max_year: int = None
) -> dict[str, list[float | None]]:
    """Busca o Balanço Patrimonial (Ativos) via endpoint /getativos do Status Invest.

    Retorna um dicionário mapeando a chave/nome do indicador para a lista de valores flutuantes.
    Exemplo:
        {
            'AtivoTotal': [2455143253000.0, 2398719197000.0, ...],
            'AtivoCirculante': [59635525000.0, 83167243000.0, ...],
            'CaixaeEquivalentesdeCaixa': [19737849000.0, 20079736000.0, ...]
        }
    """
    if max_year is None:
        max_year = datetime.datetime.now().year

    url_bs = (
        f"https://statusinvest.com.br/acao/getativos?"
        f"code={ticker.lower()}&type=0&futureData=false&"
        f"range.min={min_year}&range.max={max_year}"
    )

    bs_dict = {}

    try:
        response = requests.get(url_bs, headers=HEADERS, timeout=10)
        if response.status_code != 200:
            print(
                f"[{ticker}] Erro HTTP {response.status_code} ao buscar ativos."
            )
            return bs_dict

        res_json = response.json()

        if not res_json.get("success", False):
            print(f"[{ticker}] Resposta inválida da API do Status Invest.")
            return bs_dict

        grid = res_json.get("data", {}).get("grid", [])

        for row in grid:
            # Ignora linhas de cabeçalho
            if row.get("isHeader", False):
                continue

            grid_line = row.get("gridLineModel")
            if not grid_line:
                continue

            # Utiliza a chave limpa (ex: 'AtivoTotal', 'CaixaeEquivalentesdeCaixa')
            # Fallback para 'name' em caixa alta caso 'key' não exista
            key_name = grid_line.get("key") or str(
                grid_line.get("name", "")
            ).strip().upper()
            values = grid_line.get("values", [])

            if key_name:
                bs_dict[key_name] = values

    except Exception as e:
        print(f"[{ticker}] Erro ao processar Balanço Patrimonial: {e}")

    return bs_dict


def get_dre(ticker: str) -> dict[str, list[float | None]]:
    """Busca a DRE completa do Status Invest.

    Retorna um dicionário no formato:
    {
        'RECEITA LÍQUIDA': [val_12m, val_2023, val_2022, ...],
        'EBIT': [val_12m, val_2023, val_2022, ...],
        ...
    }
    """
    url_dre = f"https://statusinvest.com.br/acao/getdre?code={ticker.lower()}&type=0"
    dre_dict = {}

    try:
        response = requests.get(url_dre, headers=HEADERS, timeout=10)
        if response.status_code != 200:
            print(
                f"[{ticker}] Erro HTTP {response.status_code} ao buscar DRE."
            )
            return dre_dict

        res_json = response.json()
        grid = res_json.get("data", {}).get("grid", [])

        for row in grid:
            grid_line = row.get("gridLineModel", {})
            title = (
                grid_line.get("name", "").strip().upper()
                or grid_line.get("key", "").strip().upper()
            )
            raw_values = grid_line.get("values", [])

            valores_limpos = []
            for item in raw_values:
                if isinstance(item, dict):
                    val = item.get("value")
                    if val is None:
                        val = _limpar_valor_status_invest(
                            item.get("valueFormat")
                        )
                    else:
                        val = float(val)
                else:
                    val = _limpar_valor_status_invest(item)

                valores_limpos.append(val)

            dre_dict[title] = valores_limpos

    except Exception as e:
        print(f"[{ticker}] Erro ao buscar DRE completa: {e}")

    return dre_dict


def get_data_earning_yield(
    ticker: str, dre_dict: dict = None, default_divida_banco: float = 0.0
) -> pd.DataFrame:
    """Acessa o Status Invest e retorna DataFrame de EY extraindo o EBIT da DRE fornecida."""
    url_pagina = f"https://statusinvest.com.br/acoes/{ticker.lower()}"

    dados = {
        "ticker": [ticker.upper()],
        "valor_de_mercado": [pd.NA],
        "divida_liquida": [pd.NA],
        "ebit": [pd.NA],
    }

    try:
        response = requests.get(url_pagina, headers=HEADERS, timeout=10)
        if response.status_code == 200:
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

            dados["valor_de_mercado"] = [
                valor_mercado if valor_mercado is not None else pd.NA
            ]
            dados["divida_liquida"] = [
                divida_liquida if divida_liquida is not None else pd.NA
            ]

        # Extrai o EBIT diretamente da DRE já carregada
        ebit = None
        if dre_dict:
            for k, v in dre_dict.items():
                if re.search(r"\bEBIT\b", k) and v:
                    ebit = v[0]  # Pega o valor mais recente (Últimos 12 Meses)
                    break

        dados["ebit"] = [ebit if ebit is not None else pd.NA]

    except Exception as e:
        print(f"[{ticker}] Erro no processamento de EY: {e}")

    df = pd.DataFrame(dados)
    df.set_index("ticker", inplace=True)
    return df

def get_data_indicadores(ticker: str, api_code: int) -> str | None:
    """Realiza a requisição HTTP para a página da empresa no Investidor 10."""
    session = requests.Session()
    session.headers.update(HEADERS)
    ticker_clean = ticker.strip().lower()
    url_indicadores = f"https://investidor10.com.br/api/historico-indicadores/{api_code}/5/?v=2"
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