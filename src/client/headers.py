"""v1-parity table header normalization for the ``getWorksheets`` sync action.

Design spec: ``docs/superpowers/specs/2026-08-17-wr-onedrive-v2-design.md`` §5 (sync actions —
"ASCII header normalization"). This is a direct port of ``keboola.wr-onedrive`` (v1, PHP)'s
``Api\\Helpers::toAscii`` + ``Api\\Model\\TableHeader::parseColumns`` — v1 feature parity is an
acceptance criterion for the sync-action output shapes, so the normalization rules below
intentionally match v1's byte-for-byte rather than inventing a new scheme.

Kept as a standalone module (rather than folded into ``excel_writer.py``) because it has no
dependency on :class:`~client.graph_client.GraphClient` — it's pure string transformation, unit
tested directly against v1's own fixture values.
"""

import re
import unicodedata
from collections.abc import Sequence

# v1's `Helpers::toAscii`: after NFD normalization and combining-mark removal, any *run* of
# characters outside this set collapses to a single `_` (not one `_` per character).
_NON_ASCII_RUN_RE = re.compile(r"[^a-zA-Z0-9\-.]+")


def to_ascii(value: str) -> str:
    """Port of v1's ``Helpers::toAscii``: NFD-normalize, strip combining marks, ASCII-fold.

    1. Unicode NFD normalization (e.g. "š" -> "s" + a combining caron).
    2. Strip every Unicode combining mark (``unicodedata.combining(ch) != 0``) — this is what
       actually drops the diacritic, leaving the base Latin letter.
    3. Any run of characters other than ``[a-zA-Z0-9\\-.]`` collapses to a single ``_``.
    4. Leading/trailing ``_`` are trimmed (v1: ``trim($str, '_')``) — *only* leading/trailing;
       underscores produced mid-string by step 3 are left alone.
    """
    normalized = unicodedata.normalize("NFD", value)
    stripped = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    replaced = _NON_ASCII_RUN_RE.sub("_", stripped)
    return replaced.strip("_")


def normalize_header_row(cells: Sequence[str]) -> list[str]:
    """Port of v1's ``TableHeader::parseColumns``, applied to one worksheet's first-row cells.

    ``cells`` is the raw ``text`` row Graph returns for
    ``usedRange(valuesOnly=true)/row(row=0)`` — a flat list of cell strings for a genuinely
    empty sheet, Graph still returns a single-cell row (``[""]``), which v1 special-cases as
    "no header at all" (``[]``) rather than a one-column sheet named ``column-1``.

    Each non-empty cell is ASCII-folded via :func:`to_ascii`; a cell that normalizes to the
    empty string becomes ``column-{1-based index}``; a name that collides with an
    already-emitted column gets a ``-1``, ``-2``, ... suffix (first-come-first-served, in column
    order).
    """
    if len(cells) <= 1 and (not cells or cells[0] == ""):
        return []

    output: list[str] = []
    for index, cell in enumerate(cells):
        normalized = to_ascii(str(cell))
        normalized = normalized or f"column-{index + 1}"
        candidate = normalized
        suffix = 1
        while candidate in output:
            candidate = f"{normalized}-{suffix}"
            suffix += 1
        output.append(candidate)
    return output
