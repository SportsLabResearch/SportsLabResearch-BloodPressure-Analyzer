# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
import re

import numpy as np
import pandas as pd

from .base_connector import BaseConnector
from .registry import register_connector


@register_connector
class ExcelConnector(BaseConnector):
    """Conector para archivos Excel de presión arterial."""

    @property
    def name(self) -> str:
        return "Excel Blood Pressure Connector"

    @property
    def supported_extensions(self) -> tuple[str, ...]:
        return (".xlsx", ".xls")

    def detect(self, file_path: str | Path) -> bool:
        path = Path(file_path)
        return path.is_file() and path.suffix.lower() in self.supported_extensions

    def read(self, file_path: str | Path) -> pd.DataFrame:
        path = Path(file_path)

        try:
            dataframe = pd.read_excel(path)
        except Exception as exc:
            raise ValueError(
                f"No se pudo leer el archivo Excel '{path.name}': {exc}"
            ) from exc

        dataframe = (
            dataframe
            .dropna(how="all")
            .dropna(axis=1, how="all")
            .copy()
        )

        if dataframe.empty:
            raise ValueError(
                f"El archivo Excel '{path.name}' no contiene datos."
            )

        return dataframe

    def validate(self, dataframe: pd.DataFrame) -> None:
        if not isinstance(dataframe, pd.DataFrame):
            raise TypeError(
                "El contenido leído debe ser un pandas.DataFrame."
            )

        if dataframe.empty:
            raise ValueError(
                "El archivo Excel no contiene registros válidos."
            )

        sbp_column = self._find_column(
            dataframe.columns,
            exact_names=("sys(mmhg)", "sys", "pas", "sbp"),
            patterns=(
                "presion arterial sistolica",
                "systolic",
                "sistolica",
            ),
        )

        dbp_column = self._find_column(
            dataframe.columns,
            exact_names=("dia(mmhg)", "pad", "dbp"),
            patterns=(
                "presion arterial diastolica",
                "diastolic",
                "diastolica",
            ),
        )

        heart_rate_column = self._find_column(
            dataframe.columns,
            exact_names=("pr(bpm)", "pr", "fc", "hr", "heart_rate"),
            patterns=(
                "frecuencia cardiaca",
                "heart rate",
                "pulso",
                "pulse",
            ),
        )

        missing = []

        if sbp_column is None:
            missing.append("PAS/SYS/SBP")

        if dbp_column is None:
            missing.append("PAD/DIA/DBP")

        if heart_rate_column is None:
            missing.append("FC/PR/HR")

        if missing:
            raise ValueError(
                "No se detectaron las columnas obligatorias: "
                + ", ".join(missing)
            )

    def normalize(
        self,
        dataframe: pd.DataFrame,
        file_path: str | Path,
    ) -> pd.DataFrame:
        path = Path(file_path)
        columns = dataframe.columns

        subject_column = self._find_column(
            columns,
            exact_names=(
                "subject",
                "sujeto",
                "nombre",
                "nombre o id",
                "id sujeto",
            ),
            patterns=(
                "participante",
                "patient",
                "paciente",
                "athlete",
                "deportista",
            ),
        )

        datetime_column = self._find_column(
            columns,
            exact_names=(
                "fecha/hora",
                "fecha hora",
                "datetime",
                "timestamp",
                "date time",
            ),
        )

        date_column = self._find_column(
            columns,
            exact_names=("fecha", "date"),
        )

        time_column = self._find_column(
            columns,
            exact_names=("hora", "time"),
        )

        session_column = self._find_column(
            columns,
            exact_names=(
                "sesion",
                "session",
                "dia",
                "day",
            ),
        )

        measurement_column = self._find_column(
            columns,
            exact_names=(
                "medicion",
                "measurement",
                "toma",
            ),
        )

        sbp_column = self._find_column(
            columns,
            exact_names=("sys(mmhg)", "sys", "pas", "sbp"),
            patterns=(
                "presion arterial sistolica",
                "systolic",
                "sistolica",
            ),
        )

        dbp_column = self._find_column(
            columns,
            exact_names=("dia(mmhg)", "pad", "dbp"),
            patterns=(
                "presion arterial diastolica",
                "diastolic",
                "diastolica",
            ),
        )

        heart_rate_column = self._find_column(
            columns,
            exact_names=("pr(bpm)", "pr", "fc", "hr", "heart_rate"),
            patterns=(
                "frecuencia cardiaca",
                "heart rate",
                "pulso",
                "pulse",
            ),
        )

        hrv_column = self._find_column(
            columns,
            exact_names=(
                "hrv",
                "rmssd",
                "lnrmssd",
                "variabilidad fc",
            ),
            patterns=(
                "variabilidad",
                "heart rate variability",
            ),
        )

        device_column = self._find_column(
            columns,
            exact_names=(
                "device",
                "dispositivo",
                "monitor",
                "modelo",
            ),
        )

        normalized = pd.DataFrame(index=dataframe.index)

        normalized["subject"] = self._build_subject(
            dataframe,
            subject_column,
            path,
        )

        date_values, time_values = self._build_datetime(
            dataframe,
            datetime_column,
            date_column,
            time_column,
        )

        normalized["date"] = date_values
        normalized["time"] = time_values
        normalized["sbp"] = self._clean_numeric(dataframe[sbp_column])
        normalized["dbp"] = self._clean_numeric(dataframe[dbp_column])
        normalized["heart_rate"] = self._clean_numeric(
            dataframe[heart_rate_column]
        )

        if hrv_column is not None:
            normalized["hrv"] = self._clean_numeric(
                dataframe[hrv_column]
            )
        else:
            normalized["hrv"] = np.nan

        normalized["session"] = self._build_session(
            dataframe,
            normalized["date"],
            session_column,
        )

        normalized["measurement"] = self._build_measurement(
            dataframe,
            normalized["session"],
            measurement_column,
        )

        if device_column is not None:
            device_values = (
                dataframe[device_column]
                .astype("string")
                .fillna("")
                .str.strip()
            )
            normalized["device"] = device_values.mask(
                device_values.eq(""),
                "No especificado",
            )
        else:
            normalized["device"] = "No especificado"

        normalized["source"] = "Excel"
        normalized["file_name"] = path.name
        normalized["record_id"] = range(1, len(normalized) + 1)

        normalized = self._apply_physiological_filters(normalized)

        normalized = normalized.dropna(
            subset=["subject", "sbp", "dbp", "heart_rate"]
        ).copy()

        normalized["session"] = (
            pd.to_numeric(
                normalized["session"],
                errors="coerce",
            )
            .fillna(1)
            .astype(int)
        )

        normalized["measurement"] = (
            normalized
            .groupby("session")
            .cumcount()
            .add(1)
        )

        normalized["record_id"] = range(1, len(normalized) + 1)

        if normalized.empty:
            raise ValueError(
                f"El archivo '{path.name}' no contiene registros válidos."
            )

        return normalized.reset_index(drop=True)

    @staticmethod
    def _normalize_text(value) -> str:
        if pd.isna(value):
            return ""

        text = str(value).strip().lower()

        for source, target in (
            ("á", "a"),
            ("é", "e"),
            ("í", "i"),
            ("ó", "o"),
            ("ú", "u"),
            ("ü", "u"),
            ("ñ", "n"),
        ):
            text = text.replace(source, target)

        return re.sub(r"\s+", " ", text)

    def _find_column(
        self,
        columns,
        exact_names: tuple[str, ...] = (),
        patterns: tuple[str, ...] = (),
    ):
        normalized_columns = {
            column: self._normalize_text(column)
            for column in columns
        }

        exact_targets = {
            self._normalize_text(name)
            for name in exact_names
        }

        for column, normalized_name in normalized_columns.items():
            if normalized_name in exact_targets:
                return column

        normalized_patterns = [
            self._normalize_text(pattern)
            for pattern in patterns
        ]

        for column, normalized_name in normalized_columns.items():
            for pattern in normalized_patterns:
                if pattern and pattern in normalized_name:
                    return column

        return None

    @staticmethod
    def _clean_numeric(series: pd.Series) -> pd.Series:
        return pd.to_numeric(
            series
            .astype(str)
            .str.replace(",", ".", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace(r"[^0-9.\-]", "", regex=True),
            errors="coerce",
        )

    def _build_subject(
        self,
        dataframe: pd.DataFrame,
        subject_column,
        file_path: Path,
    ) -> pd.Series:
        fallback = (
            file_path.stem
            .replace("_", " ")
            .replace("-", " ")
            .strip()
        )

        if subject_column is None:
            return pd.Series(
                fallback,
                index=dataframe.index,
                dtype="string",
            )

        subjects = (
            dataframe[subject_column]
            .astype("string")
            .fillna("")
            .str.strip()
        )

        return subjects.mask(subjects.eq(""), fallback)

    @staticmethod
    def _parse_datetime_series(series: pd.Series) -> pd.Series:
        """
        Convierte fechas heterogéneas sin generar UserWarning.

        Primero intenta formato ISO; después formato europeo; finalmente
        usa format='mixed' para valores válidos no cubiertos.
        """
        raw = series.astype("string").str.strip()
        result = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

        iso_mask = raw.str.match(
            r"^\d{4}-\d{2}-\d{2}(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$",
            na=False,
        )
        if iso_mask.any():
            result.loc[iso_mask] = pd.to_datetime(
                raw.loc[iso_mask],
                errors="coerce",
                format="mixed",
                dayfirst=False,
            )

        european_mask = raw.str.match(
            r"^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$",
            na=False,
        )
        european_mask &= ~iso_mask
        if european_mask.any():
            result.loc[european_mask] = pd.to_datetime(
                raw.loc[european_mask],
                errors="coerce",
                format="mixed",
                dayfirst=True,
            )

        remaining = result.isna() & raw.ne("")
        if remaining.any():
            result.loc[remaining] = pd.to_datetime(
                raw.loc[remaining],
                errors="coerce",
                format="mixed",
                dayfirst=True,
            )

        return result

    @staticmethod
    def _parse_time_series(series: pd.Series) -> pd.Series:
        raw = series.astype("string").fillna("").str.strip()
        parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

        hhmmss = raw.str.match(r"^\d{1,2}:\d{2}:\d{2}$", na=False)
        if hhmmss.any():
            parsed.loc[hhmmss] = pd.to_datetime(
                raw.loc[hhmmss],
                errors="coerce",
                format="%H:%M:%S",
            )

        hhmm = raw.str.match(r"^\d{1,2}:\d{2}$", na=False)
        if hhmm.any():
            parsed.loc[hhmm] = pd.to_datetime(
                raw.loc[hhmm],
                errors="coerce",
                format="%H:%M",
            )

        remaining = parsed.isna() & raw.ne("")
        if remaining.any():
            parsed.loc[remaining] = pd.to_datetime(
                raw.loc[remaining],
                errors="coerce",
                format="mixed",
            )

        formatted = parsed.dt.strftime("%H:%M:%S")
        return formatted.fillna(raw)

    @classmethod
    def _build_datetime(
        cls,
        dataframe,
        datetime_column,
        date_column,
        time_column,
    ):
        if datetime_column is not None:
            values = cls._parse_datetime_series(
                dataframe[datetime_column]
            )
            return values.dt.date, values.dt.strftime("%H:%M:%S")

        if date_column is not None:
            dates = cls._parse_datetime_series(
                dataframe[date_column]
            ).dt.date
        else:
            dates = pd.Series(pd.NaT, index=dataframe.index)

        if time_column is not None:
            times = cls._parse_time_series(
                dataframe[time_column]
            )
        else:
            times = pd.Series(
                "",
                index=dataframe.index,
                dtype="string",
            )

        return dates, times

    def _build_session(
        self,
        dataframe,
        normalized_dates,
        session_column,
    ):
        if session_column is not None:
            sessions = self._clean_numeric(
                dataframe[session_column]
            )
            if sessions.notna().any():
                return sessions

        dates = pd.to_datetime(
            normalized_dates,
            errors="coerce",
        ).dt.normalize()

        if dates.notna().any():
            unique_dates = sorted(dates.dropna().unique())
            date_map = {
                date: index
                for index, date in enumerate(unique_dates, start=1)
            }
            return dates.map(date_map).ffill().bfill().fillna(1)

        return pd.Series(
            (np.arange(len(dataframe)) // 2) + 1,
            index=dataframe.index,
        )

    def _build_measurement(
        self,
        dataframe,
        sessions,
        measurement_column,
    ):
        if measurement_column is not None:
            measurements = self._clean_numeric(
                dataframe[measurement_column]
            )
            if measurements.notna().all():
                return measurements

        temporary = pd.DataFrame(
            {"session": sessions},
            index=dataframe.index,
        )

        return temporary.groupby("session").cumcount().add(1)

    @staticmethod
    def _apply_physiological_filters(dataframe):
        filtered = dataframe.copy()

        filtered.loc[
            (filtered["sbp"] < 70) | (filtered["sbp"] > 260),
            "sbp",
        ] = np.nan

        filtered.loc[
            (filtered["dbp"] < 40) | (filtered["dbp"] > 160),
            "dbp",
        ] = np.nan

        filtered.loc[
            (filtered["heart_rate"] < 35)
            | (filtered["heart_rate"] > 190),
            "heart_rate",
        ] = np.nan

        return filtered
