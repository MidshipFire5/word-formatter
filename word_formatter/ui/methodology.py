"""Создание профиля из методички: кнопка с перетаскиванием и окно проверки.

Приложение читает методичку и показывает, что в ней нашло: каждое требование
с цитатой, расхождения между текстом и оформлением файла и строки для того,
что найти не удалось. Значения можно **исправить прямо в таблице**, а пустые
строки — заполнить своими: приложение не подставляет молча то, о чём
методичка не говорит.

Профиль сохраняется только после подтверждения человеком.
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
    ORIGIN_EDIT,
    ORIGIN_SAMPLE,
    ORIGIN_TEXT,
    SAMPLE_SUFFIXES,
    TEXT_SUFFIXES,
    DraftRow,
    DraftValue,
    ProfileDraft,
)
from word_formatter.core.profiles.extraction.draft import (
    CODE_TITLES,
    VALUE_CODES,
    coerce_value,
)

DROP_SUFFIXES = tuple(TEXT_SUFFIXES)
SAMPLE_DROP_SUFFIXES = tuple(SAMPLE_SUFFIXES)

ORIGIN_TITLES = {
    ORIGIN_TEXT: "из текста методички",
    ORIGIN_SAMPLE: "из оформления файла",
    ORIGIN_EDIT: "исправлено вами",
}

VALUE_TITLES = CODE_TITLES
NOT_FOUND_HINT = "не найдено — впишите своё"

EMPTY_MISSING = "Ничего не пропущено: все перечисленные требования найдены в методичке."
EMPTY_CONFLICTS = (
    "Расхождений нет: текст методички и оформление её файла говорят одно и то же."
)

EDIT_HINT = (
    "Значение можно исправить прямо в таблице. Пустое поле означает, что "
    "требование в методичке не найдено: впишите своё значение, и оно попадёт "
    "в профиль. Если оставить пустым — подставится типовое, о чём будет "
    "сказано в ограничениях профиля."
)


def human_value(value) -> str:
    """Значение в виде, понятном человеку."""
    if isinstance(value, bool):
        return "да" if value else "нет"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, str):
        return VALUE_TITLES.get(value, value)
    return "" if value is None else str(value)


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
        self._rows = draft.editable_rows()
        self.setWindowTitle("Профиль из методички")
        self.setMinimumSize(820, 680)

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

        hint = QLabel(EDIT_HINT)
        hint.setObjectName("cardHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)

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

    def edited_values(self) -> dict:
        """Значения из таблицы: исправленные и вписанные человеком.

        Пустые поля в результат не попадают: для них останется типовое
        значение из основы профиля.
        """
        edits: dict = {}
        for row, item in zip(self._rows, self._table_items):
            text = item.text()
            original = human_value(row.value)
            if text.strip() == original.strip():
                continue
            value = coerce_value(row.path, text)
            if value is not None:
                edits[row.path] = value
        return edits

    # --- содержимое окна ------------------------------------------------------

    def _summary_text(self) -> str:
        parts = [f"Найдено требований: {len(self._draft.values)}"]
        if self._draft.sample_measured:
            parts.append("оформление файла измерено")
        if self._draft.conflicts:
            parts.append(f"расхождений текста и оформления: {len(self._draft.conflicts)}")
        if self._draft.missing:
            parts.append(f"не найдено: {len(self._draft.missing)} — их можно вписать")
        return " · ".join(parts) + "."

    def _build_tabs(self) -> QTabWidget:
        tabs = QTabWidget()
        tabs.addTab(self._build_values_table(), "Требования и значения")
        tabs.addTab(
            self._build_text_view(
                "\n".join(item.describe() for item in self._draft.conflicts) or EMPTY_CONFLICTS
            ),
            "Расхождения",
        )
        tabs.addTab(self._build_text_view(self._missing_text()), "Что не найдено")
        return tabs

    def _build_values_table(self) -> QTableWidget:
        table = QTableWidget(len(self._rows), 3)
        table.setHorizontalHeaderLabels(["Требование", "Значение", "Откуда взято"])
        table.verticalHeader().setVisible(False)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setAlternatingRowColors(True)

        self._table_items: list[QTableWidgetItem] = []
        for index, row in enumerate(self._rows):
            name_item = QTableWidgetItem(row.title or row.path)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(index, 0, name_item)

            value_item = QTableWidgetItem(human_value(row.value))
            value_item.setFlags(value_item.flags() | Qt.ItemIsEditable)
            if not row.found:
                value_item.setToolTip(row.placeholder)
                value_item.setForeground(Qt.gray)
            table.setItem(index, 1, value_item)
            self._table_items.append(value_item)

            source_item = QTableWidgetItem(self._source_text(row))
            source_item.setFlags(source_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(index, 2, source_item)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Interactive)
        header.resizeSection(1, 150)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        self._table = table
        return table

    @staticmethod
    def _source_text(row: DraftRow) -> str:
        if not row.found:
            return NOT_FOUND_HINT
        origin = ORIGIN_TITLES.get(row.origin, row.origin)
        evidence = (row.evidence or "").strip()
        return f"{origin}: «{evidence}»" if evidence else origin

    def _missing_text(self) -> str:
        if not self._draft.missing:
            return EMPTY_MISSING
        lines = [
            "Эти требования приложение в методичке не нашло — их нужно проверить",
            "и, если они есть, вписать в таблицу на первой вкладке:",
            "",
        ]
        lines.extend(f"  • {title} ({section})" for section, title in self._draft.missing)
        lines.extend(
            [
                "",
                "Если оставить поля пустыми, в профиле останутся распространённые",
                "значения, а в ограничениях профиля появится запись об этом.",
            ]
        )
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
