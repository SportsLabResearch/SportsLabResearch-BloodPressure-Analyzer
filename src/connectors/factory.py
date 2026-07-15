from __future__ import annotations

from pathlib import Path

from .registry import ConnectorRegistry


class ConnectorFactory:
    """
    Selecciona automáticamente el conector adecuado para un archivo.
    """

    @classmethod
    def detect(cls, file_path: str | Path):
        file_path = Path(file_path)

        for connector_cls in ConnectorRegistry.get_connectors():
            connector = connector_cls()

            try:
                if connector.detect(file_path):
                    return connector
            except Exception:
                continue

        raise ValueError(
            f"No existe ningún conector compatible con '{file_path.name}'."
        )

    @classmethod
    def supported_connectors(cls) -> list[str]:
        return [
            connector.__name__
            for connector in ConnectorRegistry.get_connectors()
        ]

    @classmethod
    def supported_extensions(cls) -> list[str]:
        extensiones = set()

        for connector_cls in ConnectorRegistry.get_connectors():
            connector = connector_cls()
            extensiones.update(connector.supported_extensions)

        return sorted(extensiones)
