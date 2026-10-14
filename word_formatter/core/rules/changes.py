"""Журнал изменений: что именно приложение сделало с документом.

Отчёт для пользователя строится из этого журнала, поэтому каждое изменение
записывается один раз и с достаточными подробностями, чтобы человек мог
найти место в документе и проверить его.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Change:
    """Одно изменение документа."""

    category: str
    description: str
    paragraph_index: int | None = None
    text: str = ""
    detail: str = ""


class ChangeLog:
    """Список изменений с группировкой по разделам отчёта."""

    def __init__(self) -> None:
        self._items: list[Change] = []

    def add(
        self,
        category: str,
        description: str,
        *,
        paragraph_index: int | None = None,
        text: str = "",
        detail: str = "",
    ) -> None:
        self._items.append(
            Change(
                category=category,
                description=description,
                paragraph_index=paragraph_index,
                text=text,
                detail=detail,
            )
        )

    @property
    def items(self) -> tuple[Change, ...]:
        return tuple(self._items)

    def __len__(self) -> int:
        return len(self._items)

    def by_category(self) -> dict[str, list[Change]]:
        grouped: dict[str, list[Change]] = {}
        for item in self._items:
            grouped.setdefault(item.category, []).append(item)
        return grouped

    def count(self, category: str) -> int:
        return sum(1 for item in self._items if item.category == category)

    def category_counts(self) -> dict[str, int]:
        return {category: len(items) for category, items in self.by_category().items()}
