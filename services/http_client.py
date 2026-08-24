import time
import requests
from services.parse.indicadores_parse import parse_historico_indicadores
from bs4 import BeautifulSoup
from typing import Optional
import pandas as pd

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
}

def _limpar_valor_detalhado(texto_valor: str) -> Optional[float]:
    """
    Converte strings exatas do tipo 'R$ 574.237.659.000' ou 'R$ 312.769.000.000'
    para o tipo float do Python (ex: 574237659000.0).
    """
    if not texto_valor or texto_valor.strip() == "-" or "N/A" in texto_valor:
        return None

    # Remove 'R$', espaços em branco e pontos de milhar
    texto_limpo = (
        texto_valor.replace("R$", "")
        .replace(".", "")
        .replace(",", ".")
        .strip()
    )

    try:
        return float(texto_limpo)
    except ValueError:
        return None


def get_data_earning_yield(
    ticker: str, default_divida_banco: float = 0.0
) -> pd.DataFrame:
    """Acessa o Investidor10 e retorna um DataFrame do Pandas com os dados do ticker.

    Ações de bancos/financeiras recebem `default_divida_banco` (padrão: 0.0) na
    Dívida Líquida.
    """
    url = f"https://investidor10.com.br/acoes/{ticker.lower()}/"

    # Estrutura padrão para retorno em caso de falha de requisição
    dados = {
        "ticker": [ticker.upper()],
        "valor_de_mercado": [pd.NA],
        "divida_liquida": [pd.NA],
        "ebit": [pd.NA],
    }

    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        if response.status_code != 200:
            print(f"[{ticker}] Erro HTTP: {response.status_code}")
            return pd.DataFrame(dados)

        soup = BeautifulSoup(response.text, "html.parser")

        # 1. Extração de Valor de Mercado e Dívida Líquida no container de indicadores
        container = soup.find("div", id="table-indicators-company")

        valor_mercado = None
        divida_liquida = None
        encontrou_divida = False

        if container:
            cells = container.find_all("div", class_="cell")
            for cell in cells:
                title_tag = cell.find("span", class_="title")
                if not title_tag:
                    continue

                titulo = title_tag.get_text(strip=True).upper()

                if titulo == "VALOR DE MERCADO":
                    detail_tag = cell.find("div", class_="detail-value")
                    if detail_tag:
                        valor_mercado = _limpar_valor_detalhado(
                            detail_tag.get_text(strip=True)
                        )

                elif titulo in ["DÍVIDA LÍQUIDA", "DIVIDA LIQUIDA"]:
                    encontrou_divida = True
                    detail_tag = cell.find("div", class_="detail-value")
                    if detail_tag:
                        divida_liquida = _limpar_valor_detalhado(
                            detail_tag.get_text(strip=True)
                        )

        # Trata o caso de bancos/financeiras onde o elemento não existe na página
        if not encontrou_divida:
            divida_liquida = default_divida_banco

        # 2. Extração do EBIT na tabela DRE/Balanço
        ebit = None
        tabela_balanco = soup.find("table", id="table-balance-results")

        # Se não achar por id, busca em qualquer tabela do documento
        linhas_tabela = (
            tabela_balanco.find_all("tr")
            if tabela_balanco
            else soup.find_all("tr")
        )

        for linha in linhas_tabela:
            coluna_nome = linha.find("td", class_="column-value")
            if not coluna_nome:
                continue

            texto_coluna = coluna_nome.get_text(strip=True).upper()

            # Garante que é exatamente EBIT e não EBITDA ou Margem EBIT
            if (
                "EBIT" in texto_coluna
                and "EBITDA" not in texto_coluna
                and "MARGEM" not in texto_coluna
            ):
                # O primeiro valor com classe 'detail-value' corresponde aos últimos 12 meses (ÚLT. 12M)
                detail_tag = linha.find("div", class_="detail-value")
                if detail_tag:
                    ebit = _limpar_valor_detalhado(
                        detail_tag.get_text(strip=True)
                    )
                break

        # Atribuição dos valores finais ao dicionário
        dados["valor_de_mercado"] = [
            valor_mercado if valor_mercado is not None else pd.NA
        ]
        dados["divida_liquida"] = [
            divida_liquida if divida_liquida is not None else pd.NA
        ]
        dados["ebit"] = [ebit if ebit is not None else pd.NA]

    except Exception as e:
        print(f"[{ticker}] Erro durante o scraping: {e}")

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