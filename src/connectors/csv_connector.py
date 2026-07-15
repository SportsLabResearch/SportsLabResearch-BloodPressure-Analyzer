# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

import pandas as pd

from .excel_connector import ExcelConnector
from .registry import register_connector


@register_connector
class CSVConnector(ExcelConnector):
    """Conector para archivos CSV de presión arterial."""

    @property
    def name(self) -> str:
        return "CSV Blood Pressure Connector"

    @property
    def supported_extensions(self) -> tuple[str, ...]:
        return (".csv",)

    def detect(self, file_path: str | Path) -> bool:
        path = Path(file_path)
        return (
            path.is_file()
            and path.suffix.lower() in self.supported_extensions
        )

    def read(self, file_path: str | Path) -> pd.DataFrame:
        path = Path(file_path)
        errores = []

        configuraciones = [
            ("utf-8-sig", None),
            ("utf-8", None),
            ("cp1252", None),
            ("latin-1", None),
            ("utf-8-sig", ";"),
            ("utf-8", ";"),
            ("cp1252", ";"),
            ("latin-1", ";"),
            ("utf-8-sig", ","),
            ("utf-8", ","),
            ("cp1252", ","),
            ("latin-1", ","),
            ("utf-8-sig", "\t"),
            ("utf-8", "\t"),
            ("cp1252", "\t"),
            ("latin-1", "\t"),
        ]

        for encoding, separator in configuraciones:
            try:
                opciones = {
                    "encoding": encoding,
                }

                if separator is None:
                    opciones["sep"] = None
                    opciones["engine"] = "python"
                else:
                    opciones["sep"] = separator

                dataframe = pd.read_csv(
                    path,
                    **opciones,
                )

                dataframe = (
                    dataframe
                    .dropna(how="all")
                    .dropna(axis=1, how="all")
                    .copy()
                )

                if dataframe.empty:
                    continue

                if len(dataframe.columns) < 2:
                    continue

                return dataframe

            except Exception as exc:
                errores.append(
                    f"{encoding} / {repr(separator)}: {exc}"
                )

        detalle = errores[-1] if errores else "Formato no reconocido."

        raise ValueError(
            f"No se pudo leer el archivo CSV '{path.name}'. "
            f"Último error: {detalle}"
        )

    def normalize(
        self,
        dataframe: pd.DataFrame,
        file_path: str | Path,
    ) -> pd.DataFrame:
        normalized = super().normalize(
            dataframe,
            file_path,
        )

        normalized["source"] = "CSV"
        normalized["file_name"] = Path(file_path).name

        return normalized
