"""Оформление интерфейса: цвета, шрифт и таблица стилей.

Стиль собран в одном месте, чтобы окно выглядело одинаково в Windows и Linux
и чтобы цвета и скругления не расползались по коду. Используется Qt Style
Sheets — штатный механизм оформления Qt.

Про скругление самого окна: рамку окна рисует операционная система.
В Windows 11 она и так со скруглёнными углами, в Linux — зависит от темы
рабочего стола. Своя рамка (frameless) сломала бы перетаскивание и изменение
размера, поэтому окно остаётся обычным, а скругления получают карточки,
кнопки, поля и полоса хода работы.
"""

from __future__ import annotations

from PySide6.QtGui import QFont, QFontDatabase
from PySide6.QtWidgets import QApplication, QWidget

# Цвета
BACKGROUND = "#F4F6FB"
CARD = "#FFFFFF"
BORDER = "#E3E8F2"
BORDER_STRONG = "#CFD8E8"
TEXT = "#1A2233"
TEXT_MUTED = "#6B7688"
ACCENT = "#2F6BFF"
ACCENT_HOVER = "#2760E8"
ACCENT_PRESSED = "#1F51C9"
ACCENT_SOFT = "#EDF3FF"
SUCCESS = "#17915C"
DANGER = "#CC4B45"
FIELD = "#F8FAFD"

RADIUS_CARD = 14
RADIUS_CONTROL = 10

FONT_SIZE = 10.5
PREFERRED_FAMILIES = (
    "Segoe UI",
    "Inter",
    "Noto Sans",
    "Ubuntu",
    "Cantarell",
    "DejaVu Sans",
    "Verdana",
    "Arial",
)


def pick_font_family() -> str:
    """Выбирает первый доступный шрифт из списка предпочтений.

    Список семейств можно получить только при созданном приложении Qt:
    без него обращение к QFontDatabase аварийно завершает программу.
    """
    if QApplication.instance() is None:
        return QFont().family()

    available = set(QFontDatabase.families())
    for family in PREFERRED_FAMILIES:
        if family in available:
            return family
    return QFont().family()


def application_font() -> QFont:
    """Шрифт приложения."""
    font = QFont(pick_font_family())
    font.setPointSizeF(FONT_SIZE)
    return font


def stylesheet() -> str:
    """Таблица стилей приложения."""
    return f"""
    QWidget#background {{
        background: {BACKGROUND};
    }}

    QLabel {{
        color: {TEXT};
    }}

    QLabel#appTitle {{
        font-size: 20px;
        font-weight: 600;
        color: {TEXT};
    }}

    QLabel#appSubtitle {{
        color: {TEXT_MUTED};
        font-size: 11px;
    }}

    QLabel#versionLabel {{
        color: {TEXT_MUTED};
        padding-top: 6px;
    }}

    QLabel#cardTitle {{
        font-size: 12px;
        font-weight: 600;
        color: {TEXT};
    }}

    QLabel#cardHint {{
        color: {TEXT_MUTED};
        font-size: 10px;
    }}

    QLabel#fileLabel {{
        color: {TEXT};
        padding: 8px 12px;
        background: {FIELD};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_CONTROL}px;
    }}

    QLabel#fileLabel[state="empty"] {{
        color: {TEXT_MUTED};
    }}

    QLabel#statusLabel {{
        color: {TEXT_MUTED};
    }}

    QLabel#statusLabel[state="success"] {{
        color: {SUCCESS};
        font-weight: 600;
    }}

    QLabel#statusLabel[state="error"] {{
        color: {DANGER};
        font-weight: 600;
    }}

    QLabel#statusLabel[state="busy"] {{
        color: {ACCENT};
        font-weight: 600;
    }}

    /* Карточки */
    QFrame#card, QFrame#documentCard {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_CARD}px;
    }}

    QFrame#documentCard[state="dragover"] {{
        border: 2px dashed {ACCENT};
        background: {ACCENT_SOFT};
    }}

    /* Список профилей */
    QListWidget {{
        background: transparent;
        border: none;
        outline: none;
        padding: 2px;
    }}

    QListWidget::item {{
        color: {TEXT};
        padding: 10px 12px;
        border-radius: {RADIUS_CONTROL}px;
        margin: 2px 0;
    }}

    QListWidget::item:hover {{
        background: {ACCENT_SOFT};
    }}

    QListWidget::item:selected {{
        background: {ACCENT};
        color: #FFFFFF;
    }}

    QListWidget::item:selected:!active {{
        background: {ACCENT};
        color: #FFFFFF;
    }}

    QListWidget::item:disabled {{
        color: {TEXT_MUTED};
    }}

    /* Текстовые поля */
    QPlainTextEdit {{
        background: {FIELD};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_CONTROL}px;
        padding: 10px 12px;
        color: {TEXT};
        selection-background-color: {ACCENT};
        selection-color: #FFFFFF;
    }}

    QPlainTextEdit:focus {{
        border: 1px solid {ACCENT};
    }}

    /* Кнопки */
    QPushButton {{
        background: {CARD};
        color: {TEXT};
        border: 1px solid {BORDER_STRONG};
        border-radius: {RADIUS_CONTROL}px;
        padding: 8px 16px;
    }}

    QPushButton:hover {{
        background: {ACCENT_SOFT};
        border: 1px solid {ACCENT};
        color: {ACCENT_PRESSED};
    }}

    QPushButton:pressed {{
        background: #DFE9FF;
    }}

    QPushButton:disabled {{
        background: {FIELD};
        color: #A9B2C3;
        border: 1px solid {BORDER};
    }}

    QPushButton#primaryButton {{
        background: {ACCENT};
        color: #FFFFFF;
        border: none;
        border-radius: {RADIUS_CONTROL}px;
        padding: 11px 22px;
        font-weight: 600;
        font-size: 12px;
    }}

    QPushButton#primaryButton:hover {{
        background: {ACCENT_HOVER};
    }}

    QPushButton#primaryButton:pressed {{
        background: {ACCENT_PRESSED};
    }}

    QPushButton#primaryButton:disabled {{
        background: #C9D6EE;
        color: #6E82A8;
    }}

    /* Полоса хода работы */
    QProgressBar {{
        background: #E7ECF7;
        border: none;
        border-radius: 6px;
        height: 12px;
        text-align: center;
        color: transparent;
    }}

    QProgressBar::chunk {{
        background: {ACCENT};
        border-radius: 6px;
    }}

    /* Полосы прокрутки */
    QScrollBar:vertical {{
        background: transparent;
        width: 10px;
        margin: 4px 2px 4px 0;
    }}

    QScrollBar::handle:vertical {{
        background: {BORDER_STRONG};
        border-radius: 5px;
        min-height: 30px;
    }}

    QScrollBar::handle:vertical:hover {{
        background: #B4C0D6;
    }}

    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical,
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
        background: transparent;
        height: 0;
    }}

    QScrollBar:horizontal {{
        background: transparent;
        height: 10px;
        margin: 0 4px 2px 4px;
    }}

    QScrollBar::handle:horizontal {{
        background: {BORDER_STRONG};
        border-radius: 5px;
        min-width: 30px;
    }}

    QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal,
    QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {{
        background: transparent;
        width: 0;
    }}

    /* Подсказки */
    QToolTip {{
        background: {TEXT};
        color: #FFFFFF;
        border: none;
        border-radius: 6px;
        padding: 6px 8px;
    }}

    /* Поля ввода и таблицы: нужны в окне проверки методички */
    QLineEdit {{
        background: {FIELD};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_CONTROL}px;
        padding: 7px 10px;
        color: {TEXT};
        selection-background-color: {ACCENT};
        selection-color: #FFFFFF;
    }}

    QLineEdit:focus {{
        border: 1px solid {ACCENT};
    }}

    QTabWidget::pane {{
        border: 1px solid {BORDER};
        border-radius: {RADIUS_CONTROL}px;
        background: {CARD};
        top: -1px;
    }}

    QTabBar::tab {{
        background: transparent;
        color: {TEXT_MUTED};
        padding: 8px 14px;
        margin-right: 4px;
        border: 1px solid transparent;
        border-top-left-radius: {RADIUS_CONTROL}px;
        border-top-right-radius: {RADIUS_CONTROL}px;
    }}

    QTabBar::tab:selected {{
        background: {CARD};
        color: {TEXT};
        border: 1px solid {BORDER};
        border-bottom-color: {CARD};
    }}

    QTabBar::tab:hover:!selected {{
        color: {ACCENT_PRESSED};
    }}

    QTableWidget {{
        background: {CARD};
        alternate-background-color: {FIELD};
        border: none;
        gridline-color: {BORDER};
        color: {TEXT};
        selection-background-color: {ACCENT_SOFT};
        selection-color: {TEXT};
    }}

    QHeaderView::section {{
        background: {FIELD};
        color: {TEXT_MUTED};
        border: none;
        border-bottom: 1px solid {BORDER};
        padding: 7px 8px;
        font-weight: 600;
    }}
    """


def set_state(widget: QWidget, state: str) -> None:
    """Переключает оформление виджета по состоянию.

    В Qt смена свойства сама по себе не перерисовывает виджет, поэтому
    после установки нужен повторный разбор стиля.
    """
    if widget.property("state") == state:
        return
    widget.setProperty("state", state)
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)


def apply_theme(application: QApplication) -> None:
    """Применяет оформление ко всему приложению."""
    # Fusion рисует элементы одинаково во всех системах, поэтому наша таблица
    # стилей выглядит предсказуемо и в Windows, и в Linux.
    application.setStyle("Fusion")
    application.setFont(application_font())
    application.setStyleSheet(stylesheet())
