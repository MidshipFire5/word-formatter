"""Положение подписей относительно рисунков и таблиц.

Методичка требует размещать название таблицы **над** таблицей, а подпись
рисунка — **под** рисунком. Если в документе порядок обратный, приложение
переносит подпись на нужное место.

Это единственное правило, которое меняет порядок элементов, поэтому:

- оно включается только самим профилем (`captions.*.position`);
- переносятся только подписи — текст, таблицы и рисунки не двигаются;
- подпись переносится лишь тогда, когда она стоит вплотную к своему
  рисунку или таблице; если рядом оказались две подписи, приложение
  отказывается угадывать и сообщает об этом в отчёте;
- каждое перемещение попадает в отчёт.
"""

from __future__ import annotations

from word_formatter.core.docx import Role
from word_formatter.core.rules.body import BodyElement, collect_body
from word_formatter.core.rules.context import RuleContext

CATEGORY = "caption-positions"
CATEGORY_MANUAL = "caption-positions-manual"

FIGURE_CAPTION_ROLES = (Role.FIGURE_CAPTION,)
TABLE_CAPTION_ROLES = (Role.TABLE_CAPTION, Role.TABLE_CONTINUATION)


def _neighbour_indexes(items: list[BodyElement]) -> tuple[list[int | None], list[int | None]]:
    """Для каждого элемента — ближайший непустой сосед слева и справа."""
    previous: list[int | None] = [None] * len(items)
    following: list[int | None] = [None] * len(items)

    last: int | None = None
    for index, item in enumerate(items):
        previous[index] = last
        if not item.is_blank:
            last = index

    last = None
    for index in range(len(items) - 1, -1, -1):
        following[index] = last
        if not items[index].is_blank:
            last = index

    return previous, following


def _is_claimed(element, claimed: list) -> bool:
    """Уже отнесён ли этот абзац к какому-нибудь рисунку или таблице."""
    return any(element is item for item in claimed)


def _report_move(context: RuleContext, description: str, caption: BodyElement, detail: str) -> None:
    context.changes.add(
        CATEGORY,
        description,
        paragraph_index=caption.index if caption.is_paragraph else None,
        text=caption.text,
        detail=detail,
    )


def _move(context: RuleContext, caption: BodyElement, element, *, before: bool, description: str, detail: str) -> None:
    if before:
        element.addprevious(caption.element)
    else:
        element.addnext(caption.element)
    _report_move(context, description, caption, detail)


def apply_caption_positions(context: RuleContext) -> int:
    """Переносит подписи на место, которое требует профиль. Возвращает число переносов."""
    figure_position = context.profile.captions.figure.position
    table_position = context.profile.captions.table.position
    if figure_position == "keep" and table_position == "keep":
        return 0

    items = collect_body(context)
    previous, following = _neighbour_indexes(items)

    claimed: list = []
    moves = 0
    conflicts = 0

    def neighbour(index: int | None) -> BodyElement | None:
        return items[index] if index is not None else None

    for index, item in enumerate(items):
        before = neighbour(previous[index])
        after = neighbour(following[index])

        if item.is_paragraph and item.role is Role.FIGURE and figure_position != "keep":
            before_caption = before if before is not None and before.role in FIGURE_CAPTION_ROLES else None
            after_caption = after if after is not None and after.role in FIGURE_CAPTION_ROLES else None

            if before_caption is not None and after_caption is not None:
                conflicts += 1
                continue
            if before_caption is None and after_caption is None:
                continue

            caption = before_caption or after_caption
            if _is_claimed(caption.element, claimed):
                continue
            claimed.append(caption.element)

            if figure_position == "below":
                if before_caption is not None:
                    _move(
                        context,
                        caption,
                        item.element,
                        before=False,
                        description="подпись рисунка перенесена под рисунок",
                        detail="было: над рисунком — так требует оформление рисунков",
                    )
                    moves += 1
            else:
                if after_caption is not None:
                    _move(
                        context,
                        caption,
                        item.element,
                        before=True,
                        description="подпись рисунка перенесена над рисунком",
                        detail="было: под рисунком",
                    )
                    moves += 1

        elif item.is_table and table_position != "keep":
            before_caption = before if before is not None and before.role in TABLE_CAPTION_ROLES else None
            after_caption = after if after is not None and after.role in TABLE_CAPTION_ROLES else None

            if before_caption is not None and after_caption is not None:
                conflicts += 1
                continue
            if before_caption is None and after_caption is None:
                continue

            caption = before_caption or after_caption
            if _is_claimed(caption.element, claimed):
                continue
            claimed.append(caption.element)

            if table_position == "above":
                if after_caption is not None:
                    _move(
                        context,
                        caption,
                        item.element,
                        before=True,
                        description="название таблицы перенесено над таблицу",
                        detail="было: под таблицей — методичка требует название над таблицей",
                    )
                    moves += 1
            else:
                if before_caption is not None:
                    _move(
                        context,
                        caption,
                        item.element,
                        before=False,
                        description="название таблицы перенесено под таблицу",
                        detail="было: над таблицей",
                    )
                    moves += 1

    if moves:
        context.changes.add(
            CATEGORY,
            f"подписи расставлены по местам: перенесено {moves} — подпись рисунка ставится "
            "под рисунком, название таблицы над таблицей",
            detail="это единственное изменение порядка элементов: сама подпись и её текст не меняются",
        )
    if conflicts:
        context.changes.add(
            CATEGORY_MANUAL,
            "рядом с рисунком или таблицей оказались две подписи — приложение не стало угадывать, "
            "проверьте это место вручную",
            detail=f"таких мест: {conflicts}",
        )

    return moves
