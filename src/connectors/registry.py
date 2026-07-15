from __future__ import annotations

from typing import Type

from .base_connector import BaseConnector


class ConnectorRegistry:
    """
    Registro central de conectores disponibles.

    Permite registrar, consultar y recuperar conectores sin modificar
    el núcleo del proyecto.
    """

    _connectors: list[Type[BaseConnector]] = []

    @classmethod
    def register(cls, connector: Type[BaseConnector]) -> None:
        if not issubclass(connector, BaseConnector):
            raise TypeError(
                f"{connector.__name__} debe heredar de BaseConnector."
            )

        if connector not in cls._connectors:
            cls._connectors.append(connector)

    @classmethod
    def unregister(cls, connector: Type[BaseConnector]) -> None:
        if connector in cls._connectors:
            cls._connectors.remove(connector)

    @classmethod
    def get_connectors(cls) -> list[Type[BaseConnector]]:
        return list(cls._connectors)

    @classmethod
    def clear(cls) -> None:
        cls._connectors.clear()


def register_connector(
    connector: Type[BaseConnector],
) -> Type[BaseConnector]:
    """
    Decorador para registrar automáticamente un conector.

    Ejemplo:

        @register_connector
        class BP2Connector(BaseConnector):
            ...
    """
    ConnectorRegistry.register(connector)
    return connector
