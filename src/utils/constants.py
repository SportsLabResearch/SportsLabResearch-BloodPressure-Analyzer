# -*- coding: utf-8 -*-

"""
Constantes globales del proyecto.

SportsLabResearch-BloodPressure-Analyzer
"""

from pathlib import Path


VERSION = "0.8.0-alpha"

CARPETA_DATOS = Path("data/input")
CARPETA_RESULTADOS = Path("results")
CARPETA_CONVERTIDOS = Path("data/output/converted")

UMBRAL_HOME_PAS = 135
UMBRAL_HOME_PAD = 85

UMBRAL_URGENTE_PAS = 180
UMBRAL_URGENTE_PAD = 120