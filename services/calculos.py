import re
import pandas as pd


def limpar_valor_status_invest(val) -> float | None:
    """Higieniza e converte valores brutos do Status Invest em float."""
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


def calcular_coeficientes(
    balanco: dict, dre: dict
) -> dict[str, list[float | None]]:
    # 1. Atribuição direta do Lucro Líquido (sem loop no dicionário)
    lucro_liquido = (
        dre.get("LUCRO LÍQUIDO - (R$)")
        or dre.get("LUCRO LIQUIDO - (R$)")
        or []
    )

    # 2. Dados do Balanço Patrimonial
    ativo_circ = balanco.get("AtivoCirculante", [])
    passivo_circ = balanco.get("PassivoCiruclante") or balanco.get(
        "PassivoCirculante", []
    )
    passivo_nao_circ = balanco.get("PassivoNaoCiruclante") or balanco.get(
        "PassivoNaoCirculante", []
    )
    ativo_total = balanco.get("AtivoTotal", [])
    pl = balanco.get("PatrimonioLiquidoConsolidado") or balanco.get(
        "PatrimonioLiquido", []
    )

    num_anos = len(ativo_total)

    coeficientes = {
        "liquidez_corrente": [],
        "endividamento": [],
        "roe": [],
        "alavancagem": [],
    }

    # Loop apenas para casar os valores dos anos
    for i in range(num_anos):
        ac = ativo_circ[i] if i < len(ativo_circ) else None
        pc = passivo_circ[i] if i < len(passivo_circ) else None
        pnc = passivo_nao_circ[i] if i < len(passivo_nao_circ) else None
        at = ativo_total[i] if i < len(ativo_total) else None
        patrimonio = pl[i] if i < len(pl) else None
        ll = lucro_liquido[i] if i < len(lucro_liquido) else None

        # Exigível Total
        if pc is not None and pnc is not None:
            exigivel_total = pc + pnc
        else:
            exigivel_total = pc if pc is not None else pnc

        # Coeficiente de Liquidez Corrente
        coeficientes["liquidez_corrente"].append(
            round(ac / pc, 4)
            if (ac is not None and pc is not None and pc > 0)
            else None
        )

        # Coeficiente de Endividamento
        coeficientes["endividamento"].append(
            round(exigivel_total / at, 4)
            if (exigivel_total is not None and at is not None and at > 0)
            else None
        )

        # ROE
        coeficientes["roe"].append(
            round(ll / patrimonio, 4)
            if (ll is not None and patrimonio is not None and patrimonio != 0)
            else None
        )

        # Alavancagem
        coeficientes["alavancagem"].append(
            round(at / patrimonio, 4)
            if (at is not None and patrimonio is not None and patrimonio != 0)
            else None
        )

    return coeficientes
