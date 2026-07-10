import numpy as np
import pandas as pd


def cumplimiento(valor, esperado):
    """Calcula el cumplimiento de una competencia en escala 0-1."""
    if pd.isna(valor) or pd.isna(esperado) or esperado <= 0:
        return np.nan
    brecha = valor - esperado
    if brecha >= 0:
        return 1.0
    return max(0.0, 1.0 - round(abs(brecha) / 10, 2))


def normalizar_pesos(competencias, pesos):
    """Devuelve pesos relativos que suman 1."""
    limpios = {comp: max(0.0, float(pesos.get(comp, 0.0))) for comp in competencias}
    total = sum(limpios.values())
    if total <= 0:
        raise ValueError("Al menos una competencia debe tener un peso mayor que cero.")
    return {comp: valor / total for comp, valor in limpios.items()}


def pesos_iguales_porcentaje(competencias, decimales=1):
    """Distribuye 100% en partes iguales y ajusta el último valor por redondeo."""
    competencias = list(competencias)
    if not competencias:
        return {}
    peso = round(100.0 / len(competencias), decimales)
    resultado = {comp: peso for comp in competencias}
    resultado[competencias[-1]] = round(
        100.0 - peso * (len(competencias) - 1), decimales
    )
    return resultado


def calcular(df, competencias, esperados, pesos=None):
    """Calcula CAP por competencia y CAP global ponderado."""
    df2 = df.copy()
    if not competencias:
        df2["CAP_global"] = np.nan
        return df2

    pesos = pesos or {comp: 1.0 for comp in competencias}
    pesos_norm = normalizar_pesos(competencias, pesos)
    numerador = pd.Series(0.0, index=df2.index)
    denominador = pd.Series(0.0, index=df2.index)

    for comp in competencias:
        esperado = esperados[comp]
        cumplimiento_col = f"{comp}__cumpl_sim"
        cap_col = f"{comp}__cap"
        df2[cumplimiento_col] = df2[f"{comp}__valor"].apply(
            lambda valor: cumplimiento(valor, esperado)
        )
        df2[cap_col] = df2[cumplimiento_col] * 100

        validos = df2[cap_col].notna()
        numerador = numerador.add(df2[cap_col].fillna(0) * pesos_norm[comp])
        denominador = denominador.add(validos.astype(float) * pesos_norm[comp])

    df2["CAP_global"] = numerador.div(denominador.replace(0, np.nan))
    return df2.round(3)
