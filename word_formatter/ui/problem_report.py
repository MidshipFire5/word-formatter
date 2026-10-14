"""Сообщение о проблеме: сбор отчёта и отправка по почте.

Отчёт собирается на машине пользователя и **никуда не отправляется сам**.
Приложение открывает почтовую программу с готовым письмом, а человек видит
текст до отправки и решает сам, отправлять ли его. Те же сведения можно
скопировать или сохранить в файл — на случай, если почтовая программа
не настроена.

Из отчёта убираются личные данные: домашняя папка и имя пользователя
заменяются на «<домашняя папка>» и «<пользователь>».
"""

from __future__ import annotations

import os
import platform
import sys
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import quote

from PySide6.QtCore import Qt, QUrl, qVersion
from PySide6.QtGui import QDesktopServices, QGuiApplication
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QLabel,
    QPlainTextEdit,
    QVBoxLayout,
    QWidget,
)

from word_formatter import APP_NAME, APP_VERSION
from word_formatter.core.paths import user_log_dir

SUPPORT_EMAIL = "midshipfire5@mail.ru"
LOG_FILE_NAME = "word-formatter.log"
LOG_TAIL_LINES = 25
# Почтовые программы ограничивают длину ссылки mailto, поэтому в письмо
# попадает короткий хвост журнала, а полный отчёт доступен кнопками
# «Скопировать» и «Сохранить в файл».
LOG_TAIL_LINES_IN_MAIL = 10
MAILTO_BODY_LIMIT = 1800

NOT_SET = "не указано"
REPORT_PLACEHOLDER = (
    "Что вы делали, что ожидали и что получилось. Своими словами — "
    "это самая полезная часть сообщения."
)


@dataclass
class ProblemContext:
    """Что приложению известно о моменте, когда всё пошло не так."""

    action: str = ""
    error: str = ""
    details: str = ""
    document: str = ""
    document_size: str = ""
    profile: str = ""
    result: str = ""


def mask_private(text: str) -> str:
    """Убирает из текста домашнюю папку и имя пользователя."""
    if not text:
        return text

    home = str(Path.home())
    masked = text.replace(home, "<домашняя папка>")
    masked = masked.replace(home.replace("\\", "/"), "<домашняя папка>")
    masked = masked.replace(Path.home().as_posix(), "<домашняя папка>")

    names = {
        os.environ.get("USERNAME"),
        os.environ.get("USER"),
        os.environ.get("LOGNAME"),
    }
    for name in sorted((item for item in names if item), key=len, reverse=True):
        masked = masked.replace(name, "<пользователь>")
    return masked


def environment_report() -> str:
    """Версии программы, системы и библиотек."""
    lines = [
        f"{APP_NAME} {APP_VERSION}",
        f"Система: {platform.platform()}",
        f"Python: {sys.version.split()[0]}",
        f"Qt: {qVersion()}",
    ]
    return "\n".join(mask_private(line) for line in lines)


def log_tail(lines: int = LOG_TAIL_LINES) -> str:
    """Последние строки журнала приложения."""
    if lines <= 0:
        return ""
    path = user_log_dir() / LOG_FILE_NAME
    if not path.is_file():
        return ""

    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return ""

    tail = [line for line in text.splitlines() if line.strip()][-lines:]
    return mask_private("\n".join(tail))


def collect_report(context: ProblemContext, *, log_lines: int = LOG_TAIL_LINES) -> str:
    """Собирает техническую часть отчёта."""
    parts = [environment_report(), ""]

    parts.append("Что происходило:")
    parts.append(f"  Последнее действие: {context.action or NOT_SET}")
    parts.append(f"  Ошибка: {context.error or NOT_SET}")
    if context.details:
        parts.append(f"  Подробности: {context.details}")
    parts.append(f"  Документ: {context.document or NOT_SET}")
    if context.document_size:
        parts.append(f"  Размер документа: {context.document_size}")
    parts.append(f"  Профиль: {context.profile or NOT_SET}")
    if context.result:
        parts.append(f"  Итог последней обработки: {context.result}")

    journal = log_tail(log_lines)
    if journal:
        parts.extend(["", f"Последние строки журнала ({log_lines}):", journal])

    return mask_private("\n".join(parts))


def build_message(user_text: str, context: ProblemContext, *, log_lines: int = LOG_TAIL_LINES) -> str:
    """Полный текст отчёта: сначала слова человека, потом техника."""
    description = user_text.strip() or "Описание не заполнено."
    return (
        f"{description}\n\n"
        f"--- Что приложила программа ---\n\n"
        f"{collect_report(context, log_lines=log_lines)}"
    )


def mailto_url(body: str, *, recipient: str = SUPPORT_EMAIL, subject: str | None = None) -> str:
    """Ссылка mailto с готовым письмом."""
    topic = subject or f"{APP_NAME}: сообщение о проблеме"
    trimmed = body[:MAILTO_BODY_LIMIT]
    return (
        f"mailto:{recipient}"
        f"?subject={quote(topic)}"
        f"&body={quote(trimmed)}"
    )


class ProblemReportDialog(QDialog):
    """Окно сообщения о проблеме."""

    def __init__(self, context: ProblemContext, *, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._context = context
        self.setWindowTitle("Сообщить об ошибке")
        self.setMinimumSize(720, 640)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        title = QLabel("Сообщить об ошибке")
        title.setObjectName("cardTitle")
        layout.addWidget(title)

        hint = QLabel(
            "Опишите проблему своими словами. Отчёт никуда не отправляется сам: "
            "приложение откроет почтовую программу с готовым письмом, а вы решите, "
            "отправлять ли его. Личные данные — домашняя папка и имя пользователя — "
            "из отчёта убраны."
        )
        hint.setObjectName("cardHint")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        description_label = QLabel("Что случилось")
        description_label.setObjectName("cardHint")
        layout.addWidget(description_label)

        self.description_edit = QPlainTextEdit()
        self.description_edit.setPlaceholderText(REPORT_PLACEHOLDER)
        self.description_edit.setMinimumHeight(110)
        layout.addWidget(self.description_edit)

        report_label = QLabel("Что приложится к письму")
        report_label.setObjectName("cardHint")
        layout.addWidget(report_label)

        self.report_view = QPlainTextEdit(collect_report(context))
        self.report_view.setReadOnly(True)
        layout.addWidget(self.report_view, 1)

        self.status_label = QLabel(f"Письмо будет адресовано: {SUPPORT_EMAIL}")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        buttons = QDialogButtonBox()
        self.send_button = buttons.addButton("Отправить по почте", QDialogButtonBox.AcceptRole)
        self.send_button.setObjectName("primaryButton")
        buttons.addButton("Скопировать отчёт", QDialogButtonBox.ActionRole).clicked.connect(
            self.copy_report
        )
        buttons.addButton("Сохранить в файл…", QDialogButtonBox.ActionRole).clicked.connect(
            self.save_report
        )
        buttons.addButton("Закрыть", QDialogButtonBox.RejectRole).clicked.connect(self.reject)
        buttons.accepted.connect(self.send_by_mail)
        layout.addWidget(buttons)

    # --- действия -------------------------------------------------------------

    def user_text(self) -> str:
        return self.description_edit.toPlainText().strip()

    def full_report(self) -> str:
        return build_message(self.user_text(), self._context)

    def send_by_mail(self) -> bool:
        """Открывает почтовую программу с готовым письмом."""
        if not self.user_text():
            self._set_status("Опишите, пожалуйста, что произошло — без этого письмо бесполезно.", "error")
            self.description_edit.setFocus()
            return False

        body = build_message(self.user_text(), self._context, log_lines=LOG_TAIL_LINES_IN_MAIL)
        opened = QDesktopServices.openUrl(QUrl(mailto_url(body)))
        if opened:
            self._set_status(
                f"Почтовая программа открыта, письмо адресовано {SUPPORT_EMAIL}. "
                "Если она не открылась — скопируйте отчёт и отправьте его вручную.",
                "success",
            )
        else:
            self._set_status(
                "Почтовую программу открыть не удалось. Скопируйте отчёт "
                f"и отправьте его на {SUPPORT_EMAIL}.",
                "error",
            )
        return bool(opened)

    def copy_report(self) -> None:
        clipboard = QGuiApplication.clipboard()
        if clipboard is not None:
            clipboard.setText(self.full_report())
        self._set_status("Отчёт скопирован — вставьте его в письмо, мессенджер или чат.", "success")

    def save_report(self) -> bool:
        suggested = str(Path.home() / "сообщение об ошибке Word Formatter.txt")
        selected, _filter = QFileDialog.getSaveFileName(
            self, "Сохранить отчёт", suggested, "Текстовый файл (*.txt)"
        )
        if not selected:
            return False

        path = Path(selected)
        try:
            path.write_text(self.full_report(), encoding="utf-8")
        except OSError as error:
            self._set_status(f"Не удалось сохранить файл: {error.strerror or error}.", "error")
            return False

        self._set_status(f"Отчёт сохранён: {path}. Отправьте его на {SUPPORT_EMAIL}.", "success")
        return True

    # --- служебное ------------------------------------------------------------

    def _set_status(self, message: str, state: str = "info") -> None:
        from word_formatter.ui import theme

        self.status_label.setText(message)
        theme.set_state(self.status_label, state)
