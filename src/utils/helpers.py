# -*- coding: utf-8 -*-

"""
SportsLabResearch-BloodPressure-Analyzer
Módulo de funciones auxiliares.
"""

from __future__ import annotations

import re
import sys
import numpy as np
import pandas as pd


def configurar_consola():
    for nombre in ("stdin", "stdout", "stderr"):
        flujo = getattr(sys, nombre, None)

        if hasattr(flujo, "reconfigure"):
            try:
                flujo.reconfigure(encoding="utf-8")
            except Exception:
                pass


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().lower()
    texto = texto.replace("\n", " ")

    reemplazos = (
        ("á", "a"),
        ("é", "e"),
        ("í", "i"),
        ("ó", "o"),
        ("ú", "u"),
        ("ü", "u"),
    )

    for origen, destino in reemplazos:
        texto = texto.replace(origen, destino)

    return re.sub(r"\s+", " ", texto)


def detectar_columna(columnas, patrones):
    patrones = [normalizar_texto(x) for x in patrones]

    for columna in columnas:
        nombre = normalizar_texto(columna)

        for patron in patrones:
            if patron in nombre:
                return columna

    return None


def limpiar_numero(serie):
    return pd.to_numeric(
        serie.astype(str)
        .str.replace(",", ".", regex=False)
        .str.replace(" ", "", regex=False),
        errors="coerce",
    )


def regresion_lineal(x, y):

    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)

    mascara = np.isfinite(x) & np.isfinite(y)

    x = x[mascara]
    y = y[mascara]

    if len(x) < 2:
        return np.nan, np.nan, np.nan

    pendiente, intercepto = np.polyfit(x, y, 1)

    y_pred = pendiente * x + intercepto

    ss_res = np.sum((y - y_pred) ** 2)
    ss_tot = np.sum((y - np.mean(y)) ** 2)

    r2 = (
        1 - ss_res / ss_tot
        if ss_tot > 0
        else np.nan
    )

    return pendiente, intercepto, r2


def nombre_seguro(texto):

    texto = re.sub(
        r"[^A-Za-z0-9_áéíóúÁÉÍÓÚüÜñÑ -]",
        "",
        str(texto),
    )

    texto = texto.strip().replace(" ", "_")

    return texto