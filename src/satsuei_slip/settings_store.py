from __future__ import annotations

import json
from dataclasses import dataclass

from PySide6.QtCore import QSize, QSettings


@dataclass(slots=True)
class AppSettings:
    company_name: str = ""
    project_name: str = ""
    company_names: list[str] | None = None
    project_names: list[str] | None = None
    last_pdf_dir: str = ""
    sender_footer: str = ""
    head_trim_frames: int = 8
    window_size: QSize | None = None


class SettingsStore:
    def __init__(self) -> None:
        self._settings = QSettings("SatsueiSlip", "SatsueiSlip")

    def load(self) -> AppSettings:
        size = self._settings.value("window_size")
        if not isinstance(size, QSize):
            size = None
        return AppSettings(
            company_name=str(self._settings.value("company_name", "")),
            project_name=str(self._settings.value("project_name", "")),
            company_names=self._load_string_list("company_names", "company_name"),
            project_names=self._load_string_list("project_names", "project_name"),
            last_pdf_dir=str(self._settings.value("last_pdf_dir", "")),
            sender_footer=str(self._settings.value("sender_footer", "")),
            head_trim_frames=int(self._settings.value("head_trim_frames", 8)),
            window_size=size,
        )

    def save(self, settings: AppSettings) -> None:
        self._settings.setValue("company_name", settings.company_name)
        self._settings.setValue("project_name", settings.project_name)
        self._settings.setValue("company_names", json.dumps(settings.company_names or [], ensure_ascii=False))
        self._settings.setValue("project_names", json.dumps(settings.project_names or [], ensure_ascii=False))
        self._settings.setValue("last_pdf_dir", settings.last_pdf_dir)
        self._settings.setValue("sender_footer", settings.sender_footer)
        self._settings.setValue("head_trim_frames", settings.head_trim_frames)
        if settings.window_size is not None:
            self._settings.setValue("window_size", settings.window_size)

    def _load_string_list(self, key: str, fallback_key: str) -> list[str]:
        raw_value = str(self._settings.value(key, ""))
        values: list[str] = []
        if raw_value:
            try:
                parsed = json.loads(raw_value)
                if isinstance(parsed, list):
                    values = [str(item).strip() for item in parsed if str(item).strip()]
            except json.JSONDecodeError:
                values = []

        fallback_value = str(self._settings.value(fallback_key, "")).strip()
        if fallback_value and fallback_value not in values:
            values.insert(0, fallback_value)
        return values
