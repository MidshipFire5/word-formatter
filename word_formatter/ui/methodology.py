"""Создание профиля из методички: кнопка с перетаскиванием и окно проверки.

Приложение читает методичку и показывает, что в ней нашло: каждое требование
с цитатой, расхождения между текстом и оформлением файла и список того, что
найти не удалось. Профиль сохраняется только после подтверждения человеком.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from word_formatter.core.profiles.extraction import (
    ORIGIN_SAMPLE,
    ORIGIN_TEXT,
    SAMPLE_SUFFIXES,
    TEXT_SUFFIXES,
    DraftValue,
    ProfileDraft,
)

DROP_SUFFIXES = tuple(TEXT_SUFFIXES)
SAMPLE_DROP_SUFFIXES = tuple(SAMPLE_SUFFIXES)

ORIGIN_TITLES = {
    ORIGIN_TEXT: "из текста методички",
    ORIGIN_SAMPLE: "из оформления файла",
}

VALUE_TITLES = {
    "left": "по левому краю",
    "center": "по центру",
    "right": "по правому краю",
    "justify": "по ширине",
    "above": "над элементом",
    "below": "под элементом",
    "keep": "не переносить",
    "bottom-center": "внизу по центру",
    "bottom-left": "внизу слева",
    "bottom-right": "внизу справа",
    "top-center": "вверху по центру",
    "top-left": "вверху слева",
    "top-right": "вверху справа",
}

EMPTY_MISSING = "Ничего не пропущено: все перечисленные требования найдены в методичке."
EMPTY_CONFLICTS = (
    "Расхождений нет: текст методички и оформление её файла говорят одно и то же."
)


def human_value(value) -> str:
    """Значение в виде, понятном человеку."""
    if isinstance(value, bool):
        return "да" if value else "нет"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, str):
        return VALUE_TITLES.get(value, value)
    return str(value)


class DropFileButton(QPushButton):
    """Кнопка, на которую можно перетащить файл методички."""

    fileDropped = Signal(str)

    def __init__(self, text: str = "", parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(38)
        self.setCursor(Qt.PointingHandCursor)

    @staticmethod
    def dropped_file(mime) -> str | None:
        """Путь к методичке из перетаскиваемых данных."""
        if mime is None or not mime.hasUrls():
            return None
        for url in mime.urls():
            local = url.toLocalFile()
            if local and Path(local).suffix.lower() in DROP_SUFFIXES:
                return local
        return None

    def dragEnterEvent(self, event) -> None:  # noqa: N802 - имя задано Qt
        if self.dropped_file(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:  # noqa: N802
        path = self.dropped_file(event.mimeData())
        if path:
            self.fileDropped.emit(path)
            event.acceptProposedAction()


class MethodologyDraftDialog(QDialog):
    """Показывает найденное в методичке и даёт сохранить профиль."""

    def __init__(
        self,
        draft: ProfileDraft,
        *,
        source_name: str,
        suggested_name: str,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._draft = draft
        self.setWindowTitle("Профиль из методички")
        self.setMinimumSize(760, 620)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        title = QLabel(f"Что найдено в файле «{source_name}»")
        title.setObjectName("cardTitle")
        layout.addWidget(title)

        summary = QLabel(self._summary_text())
        summary.setObjectName("cardHint")
        summary.setWordWrap(True)
        layout.addWidget(summary)

        name_label = QLabel("Название профиля")
        name_label.setObjectName("cardHint")
        layout.addWidget(name_label)

        self.name_edit = QLineEdit(suggested_name)
        self.name_edit.setMinimumHeight(34)
        layout.addWidget(self.name_edit)

        layout.addWidget(self._build_tabs(), 1)

        buttons = QDialogButtonBox()
        self.save_button = buttons.addButton("Сохранить профиль", QDialogButtonBox.AcceptRole)
        self.save_button.setObjectName("primaryButton")
        buttons.addButton("Отмена", QDialogButtonBox.RejectRole)
        buttons.accepted.connect(self._on_save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def profile_name(self) -> str:
        return self.name_edit.text().strip()

    # --- содержимое окна ------------------------------------------------------

    def _summary_text(self) -> str:
        parts = [f"Найдено требований: {len(self._draft.values)}"]
        if self._draft.sample_measured:
            parts.append("оформление файла измерено")
        if self._draft.conflicts:
            parts.append(f"расхождений текста и оформления: {len(self._draft.conflicts)}")
        if self._draft.missing:
            parts.append(f"не найдено: {len(self._draft.missing)}")
        return " · ".join(parts) + ". Проверьте значения перед сохранением."

    def _build_tabs(self) -> QTabWidget:
        tabs = QTabWidget()
        tabs.addTab(self._build_values_table(), "Найденные требования")
        tabs.addTab(
            self._build_text_view(
                "\n".join(item.describe() for item in self._draft.conflicts) or EMPTY_CONFLICTS
            ),
            "Расхождения",
        )
        tabs.addTab(self._build_text_view(self._missing_text()), "Не найдено")
        return tabs

    def _build_values_table(self) -> QTableWidget:
        rows = self._draft.value_list()
        table = QTableWidget(len(rows), 3)
        table.setHorizontalHeaderLabels(["Требование", "Значение", "Откуда взято"])
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setAlternatingRowColors(True)

        for row, item in enumerate(rows):
            table.setItem(row, 0, QTableWidgetItem(item.title or item.path))
            table.setItem(row, 1, QTableWidgetItem(human_value(item.value)))
            table.setItem(row, 2, QTableWidgetItem(self._origin_and_evidence(item)))

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        return table

    @staticmethod
    def _origin_and_evidence(item: DraftValue) -> str:
        origin = ORIGIN_TITLES.get(item.origin, item.origin)
        evidence = (item.evidence or "").strip()
        return f"{origin}: «{evidence}»" if evidence else origin

    def _missing_text(self) -> str:
        if not self._draft.missing:
            return EMPTY_MISSING
        lines = [
            "Эти требования приложение в методичке не нашло.",
            "В профиле остались распространённые значения — проверьте их:",
            "",
        ]
        lines.extend(f"  • {title} ({section})" for section, title in self._draft.missing)
        return "\n".join(lines)

    @staticmethod
    def _build_text_view(text: str) -> QPlainTextEdit:
        view = QPlainTextEdit(text)
        view.setReadOnly(True)
        return view

    # --- сохранение -----------------------------------------------------------

    def _on_save(self) -> None:
        if not self.profile_name():
            self.name_edit.setFocus()
            self.name_edit.setPlaceholderText("Введите название профиля")
            return
        self.accept()
