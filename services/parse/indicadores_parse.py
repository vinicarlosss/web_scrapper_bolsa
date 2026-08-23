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


def parse_historico_indicadores(
    dados_api: Dict[str, Any],
    indicadores: List[str] = None,
    anos: List[str] = None,
    retornar_dataframe: bool = True
) -> Union[pd.DataFrame, Dict[str, Dict[str, float]]]:
    """
    Filtra e estrutura o JSON retornado pela API de indicadores do Investidor 10.

    :param dados_api: Dicionário bruto retornado pela API.
    :param indicadores: Lista de nomes de indicadores para extrair. Se None, usa INDICADORES_PADRAO.
    :param anos: Lista de anos para extrair. Se None, usa ANOS_PADRAO.
    :param retornar_dataframe: Se True, retorna pd.DataFrame. Se False, retorna Dict.
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
                    valor_limpo = round(val, 2) if val is not None else None
                    
                    # Preenche dicionário
                    resultado_dict[ind][ano] = valor_limpo
                    
                    # Preenche lista para DataFrame
                    registros.append({
                        "Indicador": ind,
                        "Ano": ano,
                        "Valor": valor_limpo
                    })

    if retornar_dataframe:
        if not registros:
            return pd.DataFrame()
        
        df = pd.DataFrame(registros)
        df_pivot = df.pivot(index="Indicador", columns="Ano", values="Valor")
        # Garante a ordenação das colunas conforme a lista de anos solicitada
        colunas_ordenadas = [col for col in anos_alvo if col in df_pivot.columns]
        return df_pivot[colunas_ordenadas]

    return resultado_dict