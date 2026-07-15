from __future__ import annotations

from pathlib import Path
import pandas as pd

from .factory import ConnectorFactory


class ConnectorLoader:
    """
    Punto único de entrada para la adquisición de datos.

    El resto del proyecto nunca accederá directamente a un conector
    concreto. Siempre utilizará ConnectorLoader.
    """

    def __init__(self):
        self._connector = None

    @property
    def connector(self):
        return self._connector

    def load(self, file_path: str | Path) -> pd.DataFrame:
        """
        Detecta automáticamente el conector adecuado y devuelve un
        DataFrame normalizado.
        """
        file_path = Path(file_path)

        self._connector = ConnectorFactory.detect(file_path)

        dataframe = self._connector.load(file_path)

        if not isinstance(dataframe, pd.DataFrame):
            raise TypeError(
                f"{self._connector.__class__.__name__} "
                "debe devolver un pandas.DataFrame."
            )

        return dataframe

    def info(self) -> dict:
        """
        Información del último conector utilizado.
        """
        if self._connector is None:
            return {}

        return {
            "connector": self._connector.__class__.__name__,
            "name": getattr(self._connector, "name", ""),
            "extensions": getattr(
                self._connector,
                "supported_extensions",
                [],
            ),
        }
