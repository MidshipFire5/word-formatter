"""Главное окно приложения.

Окно устроено по одному экрану: выбрать профиль, выбрать документ,
применить оформление, прочитать отчёт. Никаких настроек и командной строки
обычному пользователю не требуется.

Внешний вид вынесен в модуль theme: здесь только структура и поведение.
"""

from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFontDatabase
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from word_formatter import APP_NAME, APP_VERSION
from word_formatter.core.docx.opening import ensure_supported_format
from word_formatter.core.errors import FormatterError
from word_formatter.core.processing import ProcessingResult, default_output_name
from word_formatter.core.profiles import Profile, ProfileCatalog, ProfileInfo
from word_formatter.ui import theme
from word_formatter.ui.worker import ProcessingTask

logger = logging.getLogger(__name__)

DOCX_FILTER = "Документ Word (*.docx)"
REPORT_FILTER = "Текстовый файл (*.txt)"

# Сколько названий профилей должно быть видно в списке без прокрутки.
PROFILE_ROWS_VISIBLE = 5
# Высота строки, когда список ещё пуст и измерить её не по чему.
PROFILE_FALLBACK_ROW_HEIGHT = 46
# Минимальная высота карточки отчёта: она не должна схлопываться в полоску.
REPORT_CARD_MIN_HEIGHT = 200

HINT_START = "Выберите профиль оформления и документ Word."
HINT_NEED_FILE = "Сначала выберите документ Word."
HINT_NEED_PROFILE = "Выберите профиль оформления."
HINT_READY = "Можно применять оформление."
HINT_WORKING = "Идёт обработка документа…"
HINT_DONE = "Готово. Результат сохранён в отдельный файл, исходный документ не изменялся."
HINT_FAILED = "Обработка не выполнена."

REPORT_PLACEHOLDER = (
    "Здесь появится отчёт: что изменено, что требует ручной проверки "
    "и прошла ли проверка сохранности содержимого."
)

FILE_NOT_CHOSEN = "Документ не выбран"

WINDOW_MINIMUM = (880, 620)
WINDOW_START = (1060, 780)


def _card(title: str, hint: str = "") -> tuple[QFrame, QVBoxLayout]:
    """Создаёт карточку с заголовком: белая панель со скруглёнными углами."""
    frame = QFrame()
    frame.setObjectName("card")
    layout = QVBoxLayout(frame)
    layout.setContentsMargins(18, 16, 18, 18)
    layout.setSpacing(10)

    title_label = QLabel(title)
    title_label.setObjectName("cardTitle")
    layout.addWidget(title_label)

    if hint:
        hint_label = QLabel(hint)
        hint_label.setObjectName("cardHint")
        hint_label.setWordWrap(True)
        layout.addWidget(hint_label)

    return frame, layout


class MainWindow(QMainWindow):
    """Главное окно: профили, выбор документа, обработка и отчёт."""

    processing_finished = Signal(object)
    processing_failed = Signal(str)

    def __init__(self, catalog: ProfileCatalog | None = None, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"{APP_NAME} — оформление документов Word")
        self.setMinimumSize(*WINDOW_MINIMUM)
        self.resize(*WINDOW_START)

        self._catalog = catalog or ProfileCatalog()
        self._infos: list[ProfileInfo] = []
        self._profile: Profile | None = None
        self._source_file: Path | None = None
        self._result: ProcessingResult | None = None
        self._task: ProcessingTask | None = None
        self._busy = False

        self._build_ui()
        self.refresh_profiles()
        self._update_controls()

    # --- построение интерфейса ------------------------------------------------

    def _build_ui(self) -> None:
        content = QWidget()
        content.setObjectName("background")

        layout = QVBoxLayout(content)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(16)

        layout.addLayout(self._build_header())
        layout.addLayout(self._build_profile_area(), 2)
        layout.addWidget(self._build_document_card())

        report_card = self._build_report_card()
        report_card.setMinimumHeight(REPORT_CARD_MIN_HEIGHT)
        layout.addWidget(report_card, 4)

        # Содержимое прокручивается. Без этого уменьшение окна сжимало бы
        # список профилей до одной строки: лучше прокрутка, чем неудобный выбор.
        scroll = QScrollArea()
        scroll.setObjectName("contentScroll")
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setWidget(content)
        self.setCentralWidget(scroll)

    def _build_header(self) -> QHBoxLayout:
        title = QLabel(APP_NAME)
        title.setObjectName("appTitle")

        subtitle = QLabel(
            "Оформление документов Word по требованиям выбранного профиля. "
            "Текст документа не переписывается."
        )
        subtitle.setObjectName("appSubtitle")
        subtitle.setWordWrap(True)

        text_column = QVBoxLayout()
        text_column.setSpacing(2)
        text_column.addWidget(title)
        text_column.addWidget(subtitle)

        version = QLabel(f"версия {APP_VERSION}")
        version.setObjectName("versionLabel")
        version.setAlignment(Qt.AlignRight | Qt.AlignTop)

        header = QHBoxLayout()
        header.setSpacing(16)
        header.addLayout(text_column, 1)
        header.addWidget(version)
        return header

    def _build_profile_area(self) -> QHBoxLayout:
        profiles_card, profiles_layout = _card(
            "Профиль оформления", "Выберите требования, по которым оформлять документ."
        )
        self.profile_list = QListWidget()
        self.profile_list.setMinimumWidth(300)
        # Названия профилей длинные: переносим строки, а не прокручиваем вбок.
        self.profile_list.setWordWrap(True)
        self.profile_list.setTextElideMode(Qt.ElideNone)
        self.profile_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.profile_list.setVerticalScrollMode(QListWidget.ScrollPerPixel)
        self.profile_list.currentItemChanged.connect(self._on_profile_changed)
        profiles_layout.addWidget(self.profile_list, 1)

        from word_formatter.ui.methodology import DropFileButton

        self.create_profile_button = DropFileButton("Создать профиль из методички…")
        self.create_profile_button.setToolTip(
            "Выберите файл методички — Word, PDF или текстовый — "
            "или перетащите его прямо на эту кнопку"
        )
        self.create_profile_button.clicked.connect(self.choose_methodology_file)
        self.create_profile_button.fileDropped.connect(self.create_profile_from_methodology)
        profiles_layout.addWidget(self.create_profile_button)

        profile_hint = QLabel(
            "Нет вашей методички в списке? Перетащите её файл на кнопку выше — "
            "приложение прочитает требования и покажет, что нашло."
        )
        profile_hint.setObjectName("cardHint")
        profile_hint.setWordWrap(True)
        profiles_layout.addWidget(profile_hint)

        description_card, description_layout = _card(
            "Что делает профиль", "Источник требований и ограничения выбранного профиля."
        )
        self.profile_description = QPlainTextEdit()
        self.profile_description.setReadOnly(True)
        self.profile_description.setPlaceholderText("Выберите профиль, чтобы увидеть его описание.")
        description_layout.addWidget(self.profile_description, 1)

        area = QHBoxLayout()
        area.setSpacing(16)
        area.addWidget(profiles_card, 2)
        area.addWidget(description_card, 3)
        return area

    def _build_document_card(self) -> QFrame:
        self.document_card = QFrame()
        self.document_card.setObjectName("documentCard")
        card_layout = QVBoxLayout(self.document_card)
        card_layout.setContentsMargins(18, 16, 18, 18)
        card_layout.setSpacing(12)

        title = QLabel("Документ Word")
        title.setObjectName("cardTitle")
        card_layout.addWidget(title)

        hint = QLabel(
            "Выберите файл «.docx» или перетащите его в это окно. "
            "Исходный документ не изменяется — результат сохраняется отдельно."
        )
        hint.setObjectName("cardHint")
        hint.setWordWrap(True)
        card_layout.addWidget(hint)

        self.choose_button = QPushButton("Выбрать файл Word")
        self.choose_button.clicked.connect(self.choose_source_file)
        self.choose_button.setMinimumHeight(38)

        self.file_label = QLabel(FILE_NOT_CHOSEN)
        self.file_label.setObjectName("fileLabel")
        theme.set_state(self.file_label, "empty")

        file_row = QHBoxLayout()
        file_row.setSpacing(12)
        file_row.addWidget(self.choose_button)
        file_row.addWidget(self.file_label, 1)
        card_layout.addLayout(file_row)

        self.apply_button = QPushButton("Применить оформление")
        self.apply_button.setObjectName("primaryButton")
        self.apply_button.clicked.connect(self.apply_formatting)
        self.apply_button.setMinimumHeight(42)

        self.save_report_button = QPushButton("Сохранить отчёт…")
        self.save_report_button.clicked.connect(self.save_report_dialog)
        self.save_report_button.setMinimumHeight(42)

        buttons_row = QHBoxLayout()
        buttons_row.setSpacing(12)
        buttons_row.addWidget(self.apply_button, 2)
        buttons_row.addWidget(self.save_report_button, 1)
        card_layout.addLayout(buttons_row)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)  # режим «идёт работа»: точных долей мы не знаем
        self.progress.setVisible(False)
        card_layout.addWidget(self.progress)

        self.status_label = QLabel(HINT_START)
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        card_layout.addWidget(self.status_label)

        self.setAcceptDrops(True)
        return self.document_card

    def _build_report_card(self) -> QFrame:
        card, card_layout = _card(
            "Отчёт об обработке",
            "Что изменено, что требует ручной проверки и целы ли содержимое и оформление.",
        )
        self.report_view = QPlainTextEdit()
        self.report_view.setReadOnly(True)
        self.report_view.setFont(QFontDatabase.systemFont(QFontDatabase.FixedFont))
        self.report_view.setPlainText(REPORT_PLACEHOLDER)
        card_layout.addWidget(self.report_view, 1)
        return card

    # --- перетаскивание файла -------------------------------------------------

    @staticmethod
    def _dropped_documents(mime) -> list[Path]:
        """Пути документов Word из перетаскиваемых данных."""
        if mime is None or not mime.hasUrls():
            return []
        paths: list[Path] = []
        for url in mime.urls():
            local = url.toLocalFile()
            if local and Path(local).suffix.lower() == ".docx":
                paths.append(Path(local))
        return paths

    def dragEnterEvent(self, event) -> None:  # noqa: N802 - имя задано Qt
        if self._busy:
            event.ignore()
            return
        if self._dropped_documents(event.mimeData()):
            theme.set_state(self.document_card, "dragover")
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:  # noqa: N802
        theme.set_state(self.document_card, "set")
        event.accept()

    def dropEvent(self, event) -> None:  # noqa: N802
        theme.set_state(self.document_card, "set")
        documents = self._dropped_documents(event.mimeData())
        if documents:
            self.set_source_file(documents[0])
            event.acceptProposedAction()

    # --- профили --------------------------------------------------------------

    def refresh_profiles(self) -> None:
        """Заполняет список профилей из каталога."""
        self._infos = self._catalog.list_infos()
        self.profile_list.clear()

        for info in self._infos:
            title = info.title
            if info.error:
                title = f"⚠ {title} — ошибка в файле профиля"
            item = QListWidgetItem(title)
            item.setData(Qt.UserRole, info.id)
            self.profile_list.addItem(item)

        if self.profile_list.count():
            self.profile_list.setCurrentRow(0)

        self._fit_profile_list()

    def _fit_profile_list(self) -> None:
        """Оставляет в списке профилей место хотя бы для пяти названий.

        Без минимальной высоты при уменьшении окна список сжимается до одной
        строки, и выбирать профиль становится неудобно. Названия переносятся
        по строкам, поэтому высота строки берётся по самой высокой из первых
        строк списка.
        """
        rows_visible = PROFILE_ROWS_VISIBLE
        rows = min(self.profile_list.count(), rows_visible)
        heights = [self.profile_list.sizeHintForRow(index) for index in range(rows)]
        row_height = max(heights) if heights else PROFILE_FALLBACK_ROW_HEIGHT
        self.profile_list.setMinimumHeight(
            row_height * rows_visible + 2 * self.profile_list.frameWidth() + 8
        )

    def select_profile(self, profile_id: str) -> bool:
        """Выбирает профиль по идентификатору. Возвращает признак успеха."""
        for row in range(self.profile_list.count()):
            item = self.profile_list.item(row)
            if item.data(Qt.UserRole) == profile_id:
                self.profile_list.setCurrentRow(row)
                return True
        return False

    def current_profile_id(self) -> str | None:
        item = self.profile_list.currentItem()
        return item.data(Qt.UserRole) if item is not None else None

    def _on_profile_changed(self, current: QListWidgetItem | None, _previous=None) -> None:
        self._profile = None
        if current is None:
            self.profile_description.setPlainText("")
            self._update_controls()
            return

        info = next((item for item in self._infos if item.id == current.data(Qt.UserRole)), None)
        if info is None:
            self._update_controls()
            return

        if info.error:
            self.profile_description.setPlainText(
                f"{info.name}\n\nПрофиль не удалось прочитать.\n\n{info.error}\n\nФайл: {info.path}"
            )
            self._update_controls()
            return

        try:
            self._profile = self._catalog.load(info.id)
        except FormatterError as error:
            self.profile_description.setPlainText(f"{info.name}\n\n{error.user_message}")
            self._update_controls()
            return

        self.profile_description.setPlainText(self._describe_profile(info, self._profile))
        self._update_controls()

    def _describe_profile(self, info: ProfileInfo, profile: Profile) -> str:
        lines = [
            info.name,
            "",
            f"Версия: {profile.version}",
            f"Тип: {'демонстрационный профиль' if profile.demo else 'профиль оформления'}"
            + (", встроенный" if info.origin == "builtin" else ", личный"),
            f"Источник требований: {profile.source_title}",
        ]
        if profile.applies_to:
            lines.append("Применяется к: " + ", ".join(profile.applies_to))
        if profile.description:
            lines.extend(["", profile.description])
        if profile.limitations:
            lines.append("")
            lines.append("Что профиль не делает:")
            lines.extend(f"  • {item}" for item in profile.limitations)
        lines.extend(
            [
                "",
                "Оформление:",
                f"  • страница {profile.page.format}, поля "
                f"{profile.page.margins.left:g}/{profile.page.margins.right:g}/"
                f"{profile.page.margins.top:g}/{profile.page.margins.bottom:g} мм",
                f"  • основной текст {profile.body.font} {profile.body.size_pt:g} пт, "
                f"интервал {profile.body.line_spacing:g}, отступ "
                f"{profile.body.first_line_indent_cm:g} см",
                f"  • таблицы {profile.tables.font_size_pt:g} пт, "
                f"сноски {profile.footnotes.font_size_pt:g} пт",
                f"  • заголовки уровней: {profile.headings.max_level}",
            ]
        )
        return "\n".join(lines)

    # --- создание профиля из методички ----------------------------------------

    def choose_methodology_file(self) -> None:
        """Открывает диалог выбора файла методички."""
        from word_formatter.core.profiles.extraction import FILE_FILTER

        selected, _filter = QFileDialog.getOpenFileName(
            self, "Выберите файл методички", str(Path.home()), FILE_FILTER
        )
        if selected:
            self.create_profile_from_methodology(selected)

    def create_profile_from_methodology(self, path: str | Path) -> str | None:
        """Читает методичку, показывает найденное и сохраняет профиль.

        Возвращает идентификатор созданного профиля или None, если человек
        отказался или файл прочитать не удалось.
        """
        from PySide6.QtWidgets import QApplication, QDialog

        from word_formatter.core.errors import FormatterError
        from word_formatter.core.profiles.extraction import (
            build_profile_draft,
            save_draft_profile,
        )
        from word_formatter.ui.methodology import MethodologyDraftDialog

        source = Path(path)
        self._set_status(f"Читаю методичку «{source.name}»…", "busy")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            draft = build_profile_draft(source)
        except FormatterError as error:
            self._set_status("Методичку прочитать не удалось.", "error")
            QMessageBox.warning(self, "Не удалось прочитать методичку", error.user_message)
            return None
        finally:
            QApplication.restoreOverrideCursor()

        dialog = MethodologyDraftDialog(
            draft,
            source_name=source.stem,
            suggested_name=f"Моя методичка: {source.stem}",
            parent=self,
        )
        if dialog.exec() != QDialog.Accepted:
            self._set_status("Создание профиля отменено.")
            return None

        name = dialog.profile_name()
        directory = self._catalog.directories[1]
        try:
            saved = save_draft_profile(draft, name=name, directory=directory)
        except FormatterError as error:
            self._set_status("Профиль сохранить не удалось.", "error")
            QMessageBox.warning(self, "Не удалось сохранить профиль", error.user_message)
            return None

        self.refresh_profiles()
        if self.select_profile(saved.stem):
            self._set_status(
                f"Профиль «{name}» создан и выбран. Теперь выберите документ для оформления.",
                "success",
            )
        else:
            self._set_status(f"Профиль сохранён: {saved}", "success")
        return saved.stem

    # --- выбор документа ------------------------------------------------------

    def choose_source_file(self) -> None:
        """Открывает диалог выбора документа Word."""
        start_directory = str(self._source_file.parent) if self._source_file else str(Path.home())
        selected, _filter = QFileDialog.getOpenFileName(
            self, "Выберите документ Word", start_directory, DOCX_FILTER
        )
        if selected:
            self.set_source_file(Path(selected))

    def set_source_file(self, path: Path) -> bool:
        """Запоминает выбранный документ. Возвращает признак успеха."""
        try:
            ensure_supported_format(path)
        except FormatterError as error:
            QMessageBox.warning(self, "Формат не поддерживается", error.user_message)
            return False

        if not path.is_file():
            QMessageBox.warning(
                self,
                "Файл не найден",
                f"Файл «{path.name}» не найден. Возможно, он был перемещён или удалён.",
            )
            return False

        self._source_file = path
        self._result = None
        self._update_file_label()
        self._update_controls()
        return True

    def _update_file_label(self) -> None:
        """Показывает имя и размер выбранного файла, полный путь — в подсказке."""
        if self._source_file is None:
            self.file_label.setText(FILE_NOT_CHOSEN)
            self.file_label.setToolTip("")
            theme.set_state(self.file_label, "empty")
            return

        size_kb = self._source_file.stat().st_size / 1024
        self.file_label.setText(f"{self._source_file.name}  ·  {size_kb:.0f} КБ")
        self.file_label.setToolTip(str(self._source_file))
        theme.set_state(self.file_label, "set")

    # --- обработка ------------------------------------------------------------

    def default_output_path(self) -> Path | None:
        if self._source_file is None:
            return None
        return self._source_file.with_name(default_output_name(self._source_file))

    def apply_formatting(self) -> None:
        """Спрашивает, куда сохранить результат, и запускает обработку."""
        if self._source_file is None:
            self._set_status(HINT_NEED_FILE, "error")
            return
        if self._profile is None:
            self._set_status(HINT_NEED_PROFILE, "error")
            return

        suggested = self.default_output_path()
        selected, _filter = QFileDialog.getSaveFileName(
            self,
            "Сохранить оформленный документ",
            str(suggested) if suggested else "",
            DOCX_FILTER,
        )
        if not selected:
            self._set_status(HINT_READY)
            return

        self.start_processing(Path(selected))

    def start_processing(self, output_path: Path) -> ProcessingTask | None:
        """Запускает обработку в отдельном потоке."""
        if self._source_file is None or self._profile is None or self._busy:
            return None

        self._set_busy(True)
        self._set_status(HINT_WORKING, "busy")
        self.report_view.setPlainText("")

        task = ProcessingTask(self._source_file, output_path, self._profile)
        task.progress.connect(self._set_status)
        task.finished.connect(self._on_finished)
        task.failed.connect(self._on_failed)
        self._task = task
        task.start()
        return task

    def _on_finished(self, result: ProcessingResult) -> None:
        self._result = result
        self._set_busy(False)
        self.report_view.setPlainText(result.report.as_text())
        if result.verification.passed:
            self._set_status(HINT_DONE, "success")
        else:
            self._set_status(HINT_FAILED, "error")
        self._update_controls()
        self.processing_finished.emit(result)

        if result.verification.passed:
            QMessageBox.information(
                self,
                "Оформление применено",
                f"Документ оформлен.\n\nРезультат: {result.output_path}\n"
                "Исходный документ не изменялся.\n\n"
                "Проверка сохранности содержимого пройдена. Подробности — в отчёте.",
            )
        else:
            QMessageBox.warning(
                self,
                "Проверьте результат",
                "Документ сохранён, но проверка сохранности обнаружила расхождения:\n\n"
                + "\n".join(f"• {problem}" for problem in result.verification.problems)
                + "\n\nПроверьте эти места вручную.",
            )

    def _on_failed(self, message: str) -> None:
        self._set_busy(False)
        self._set_status(HINT_FAILED, "error")
        self.processing_failed.emit(message)
        QMessageBox.critical(self, "Не удалось обработать документ", message)

    def save_report_dialog(self) -> None:
        """Сохраняет отчёт в текстовый файл."""
        if self._result is None:
            self._set_status("Отчёт появится после обработки документа.")
            return

        source_name = self._result.source_path.stem
        suggested = self._result.output_path.with_name(f"{source_name}_отчёт.txt")
        selected, _filter = QFileDialog.getSaveFileName(
            self, "Сохранить отчёт", str(suggested), REPORT_FILTER
        )
        if not selected:
            return
        self.save_report(Path(selected))

    def save_report(self, path: Path) -> bool:
        """Записывает отчёт в файл."""
        if self._result is None:
            return False
        try:
            path.write_text(self._result.report.as_text(), encoding="utf-8")
        except OSError as error:
            QMessageBox.critical(
                self,
                "Не удалось сохранить отчёт",
                f"Не удалось записать файл «{path.name}»: {error.strerror or error}.",
            )
            return False
        self._set_status(f"Отчёт сохранён: {path}", "success")
        return True

    # --- служебное ------------------------------------------------------------

    def _set_status(self, message: str, state: str = "info") -> None:
        self.status_label.setText(message)
        theme.set_state(self.status_label, state)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        self.progress.setVisible(busy)
        self._update_controls()

    def _update_controls(self) -> None:
        can_apply = (
            not self._busy and self._profile is not None and self._source_file is not None
        )
        self.apply_button.setEnabled(can_apply)
        self.choose_button.setEnabled(not self._busy)
        self.profile_list.setEnabled(not self._busy)
        self.save_report_button.setEnabled(self._result is not None and not self._busy)
        self.save_report_button.setToolTip(
            "Сохранить отчёт в текстовый файл"
            if self._result is not None
            else "Отчёт появится после обработки документа"
        )

        # Когда всё выбрано, подсказываем следующий шаг вместо начальной строки.
        if can_apply and self._result is None and self.status_label.text() in (
            HINT_START,
            HINT_NEED_FILE,
            HINT_NEED_PROFILE,
        ):
            self._set_status(HINT_READY)

    def closeEvent(self, event) -> None:  # noqa: N802 - имя задано Qt
        if self._task is not None and self._busy:
            self._task.wait(5000)
        event.accept()
