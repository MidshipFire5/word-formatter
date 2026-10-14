"""Контекст применения правил: документ, профиль, инвентарь и журнал изменений."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from docx.text.paragraph import Paragraph

from word_formatter.core.docx import DocumentInventory
from word_formatter.core.profiles import Profile
from word_formatter.core.rules.changes import ChangeLog


@dataclass
class RuleContext:
    """Всё, что нужно правилу для работы."""

    document: Any
    profile: Profile
    inventory: DocumentInventory
    changes: ChangeLog = field(default_factory=ChangeLog)
    order_changes: int = 0
    """Сколько раз правила меняли порядок элементов (перенос подписей)."""

    def paragraph(self, index: int) -> Paragraph:
        """Абзац верхнего уровня по его номеру из инвентаря."""
        return self.document.paragraphs[index]

    def table(self, index: int):
        return self.document.tables[index]
