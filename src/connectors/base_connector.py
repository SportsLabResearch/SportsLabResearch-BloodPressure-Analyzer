from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

import pandas as pd


class BaseConnector(ABC):
    """Contrato base para todos los conectores de datos."""

    STANDARD_COLUMNS = [
        "subject",
        "date",
        "time",
        "session",
        "measurement",
        "sbp",
        "dbp",
        "heart_rate",
        "hrv",
        "device",
        "source",
        "file_name",
        "record_id",
    ]

    REQUIRED_COLUMNS = [
        "subject",
        "date",
        "time",
        "session",
        "measurement",
        "sbp",
        "dbp",
        "heart_rate",
        "device",
        "source",
        "file_name",
        "record_id",
    ]

    @property
    @abstractmethod
    def name(self) -> str:
        raise NotImplementedError

    @property
    @abstractmethod
    def supported_extensions(self) -> tuple[str, ...]:
        raise NotImplementedError

    @abstractmethod
    def detect(self, file_path: str | Path) -> bool:
        raise NotImplementedError

    @abstractmethod
    def read(self, file_path: str | Path) -> pd.DataFrame:
        raise NotImplementedError

    @abstractmethod
    def validate(self, dataframe: pd.DataFrame) -> None:
        raise NotImplementedError

    @abstractmethod
    def normalize(
        self,
        dataframe: pd.DataFrame,
        file_path: str | Path,
    ) -> pd.DataFrame:
        raise NotImplementedError

    def load(self, file_path: str | Path) -> pd.DataFrame:
        path = self._validate_file_path(file_path)
        raw_dataframe = self.read(path)

        if not isinstance(raw_dataframe, pd.DataFrame):
            raise TypeError(
                f"{self.__class__.__name__}.read() debe devolver un pandas.DataFrame."
            )

        self.validate(raw_dataframe)
        normalized_dataframe = self.normalize(raw_dataframe.copy(), path)

        if not isinstance(normalized_dataframe, pd.DataFrame):
            raise TypeError(
                f"{self.__class__.__name__}.normalize() debe devolver un pandas.DataFrame."
            )

        normalized_dataframe = self._prepare_standard_dataframe(
            normalized_dataframe
        )
        self.validate_standard_dataframe(normalized_dataframe)
        return normalized_dataframe

    def validate_standard_dataframe(self, dataframe: pd.DataFrame) -> None:
        missing = [
            column
            for column in self.REQUIRED_COLUMNS
            if column not in dataframe.columns
        ]
        if missing:
            raise ValueError(
                "Faltan columnas obligatorias en el DataFrame normalizado: "
                + ", ".join(missing)
            )

        if dataframe.empty:
            raise ValueError("El DataFrame normalizado no contiene registros.")

        if dataframe["subject"].astype(str).str.strip().eq("").any():
            raise ValueError("Existen registros sin sujeto.")

        for column in ("sbp", "dbp", "heart_rate"):
            if dataframe[column].isna().any():
                raise ValueError(
                    f"La columna '{column}' contiene valores ausentes."
                )

        if (dataframe["session"] < 1).any():
            raise ValueError("La columna 'session' contiene valores menores que 1.")

        if (dataframe["measurement"] < 1).any():
            raise ValueError(
                "La columna 'measurement' contiene valores menores que 1."
            )

        if dataframe["record_id"].duplicated().any():
            raise ValueError("La columna 'record_id' contiene duplicados.")

    def _validate_file_path(self, file_path: str | Path) -> Path:
        path = Path(file_path)

        if not path.exists():
            raise FileNotFoundError(f"No existe el archivo: {path}")

        if not path.is_file():
            raise ValueError(f"La ruta no corresponde a un archivo: {path}")

        extension = path.suffix.lower()
        supported = tuple(
            ext.lower() if ext.startswith(".") else f".{ext.lower()}"
            for ext in self.supported_extensions
        )

        if extension not in supported:
            raise ValueError(
                f"Extensión no compatible con {self.name}: {extension}. "
                f"Extensiones admitidas: {', '.join(supported)}"
            )

        return path

    def _prepare_standard_dataframe(
        self,
        dataframe: pd.DataFrame,
    ) -> pd.DataFrame:
        prepared = dataframe.copy()

        if "hrv" not in prepared.columns:
            prepared["hrv"] = pd.NA

        for column in self.STANDARD_COLUMNS:
            if column not in prepared.columns:
                prepared[column] = pd.NA

        prepared = prepared[self.STANDARD_COLUMNS].copy()
        prepared["subject"] = prepared["subject"].astype("string").str.strip()
        prepared["date"] = pd.to_datetime(
            prepared["date"],
            errors="coerce",
            dayfirst=True,
        ).dt.date
        prepared["time"] = prepared["time"].astype("string").str.strip()

        for column in ("session", "measurement", "record_id"):
            prepared[column] = pd.to_numeric(
                prepared[column],
                errors="coerce",
            ).astype("Int64")

        for column in ("sbp", "dbp", "heart_rate", "hrv"):
            prepared[column] = pd.to_numeric(
                prepared[column],
                errors="coerce",
            )

        for column in ("device", "source", "file_name"):
            prepared[column] = prepared[column].astype("string").str.strip()

        return prepared.reset_index(drop=True)

    def info(self) -> dict[str, Any]:
        return {
            "connector": self.__class__.__name__,
            "name": self.name,
            "supported_extensions": list(self.supported_extensions),
            "standard_columns": list(self.STANDARD_COLUMNS),
        }
