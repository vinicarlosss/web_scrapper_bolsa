import pandas as pd
from typing import Dict, Any, Union, List

# Lista padrão de indicadores com os nomes exatos retornados pelo Investidor 10
INDICADORES_PADRAO = [
    "LPA",
    "VPA",
    "P/L",
    "Payout",
    "Margem Líquida",
    "Margem Bruta"
]

ANOS_PADRAO = ["Atual", "2025", "2024", "2023", "2022", "2021"]


def _converter_para_float(val: Any) -> Union[float, None]:
    """Converte valores brutos (string, int ou float) recebidos da API para float de forma segura."""
    if val is None:
        return None

    if isinstance(val, (int, float)):
        return round(float(val), 2)

    if isinstance(val, str):
        val_str = val.strip().replace("%", "").replace("R$", "")

        if val_str in ["", "-", "N/A", "NaN", "null"]:
            return None

        # Trata formato numérico brasileiro (ex: "1.234,56" -> "1234.56" ou "12,34" -> "12.34")
        if "," in val_str and "." in val_str:
            val_str = val_str.replace(".", "").replace(",", ".")
        elif "," in val_str:
            val_str = val_str.replace(",", ".")

        try:
            return round(float(val_str), 2)
        except ValueError:
            return None

    return None


def parse_historico_indicadores(
    dados_api: Dict[str, Any],
    indicadores: List[str] = None,
    anos: List[str] = None,
    retornar_dataframe: bool = True,
) -> Union[pd.DataFrame, Dict[str, Dict[str, float]]]:
    """Filtra e estrutura o JSON retornado pela API de indicadores do Investidor 10.

    :param dados_api: Dicionário bruto retornado pela API.
    :param indicadores: Lista de nomes de indicadores para extrair. Se None,
    usa INDICADORES_PADRAO.
    :param anos: Lista de anos para extrair. Se None, usa ANOS_PADRAO.
    :param retornar_dataframe: Se True, retorna pd.DataFrame. Se False, retorna
    Dict.
    :return: DataFrame pivoteado ou Dicionário estruturado por [indicador][ano].
    """
    indicadores_alvo = indicadores or INDICADORES_PADRAO
    anos_alvo = anos or ANOS_PADRAO

    registros = []
    resultado_dict = {}

    for ind in indicadores_alvo:
        if ind in dados_api:
            resultado_dict[ind] = {}
            for item in dados_api[ind]:
                ano = item.get("year")
                if ano in anos_alvo:
                    val = item.get("value")

                    # Converte string/num para float tratado de forma segura
                    valor_limpo = _converter_para_float(val)

                    # Preenche dicionário
                    resultado_dict[ind][ano] = valor_limpo

                    # Preenche lista para DataFrame
                    registros.append(
                        {"Indicador": ind, "Ano": ano, "Valor": valor_limpo}
                    )

    if retornar_dataframe:
        if not registros:
            return pd.DataFrame()

        df = pd.DataFrame(registros)
        df_pivot = df.pivot(index="Indicador", columns="Ano", values="Valor")
        # Garante a ordenação das colunas conforme a lista de anos solicitada
        colunas_ordenadas = [
            col for col in anos_alvo if col in df_pivot.columns
        ]
        return df_pivot[colunas_ordenadas]

    return resultado_dict