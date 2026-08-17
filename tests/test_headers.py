"""Unit tests for `client.headers` — v1-parity table-header normalization (plan Task 8).

Values are taken directly from ``keboola.wr-onedrive`` (v1, PHP)'s own test fixtures
(``tests/datadir/get-worksheets-*``, copied into ``tests/fixtures/v1_parity``) so this is a
byte-for-byte port of ``Api\\Helpers::toAscii`` + ``Api\\Model\\TableHeader::parseColumns``, not a
reinvention.
"""

from client.headers import normalize_header_row, to_ascii


class TestToAscii:
    def test_space_becomes_underscore(self):
        assert to_ascii("Col 1") == "Col_1"

    def test_diacritics_are_stripped_to_base_letter(self):
        # v1 fixture: worksheet name "Zošit 2" (untouched — only header *cells* go through
        # `toAscii`, not sheet names); this exercises the same NFD + combining-mark-strip
        # machinery on a header cell instead.
        assert to_ascii("Zošit") == "Zosit"

    def test_run_of_disallowed_characters_collapses_to_one_underscore(self):
        assert to_ascii("a   b") == "a_b"
        assert to_ascii("a!!!b") == "a_b"

    def test_leading_and_trailing_underscores_are_trimmed(self):
        assert to_ascii("  leading and trailing  ") == "leading_and_trailing"

    def test_dot_and_hyphen_are_preserved(self):
        assert to_ascii("col-1.2") == "col-1.2"

    def test_digits_are_preserved(self):
        assert to_ascii("2024") == "2024"

    def test_empty_string_stays_empty(self):
        assert to_ascii("") == ""

    def test_equals_sign_is_not_a_reserved_character(self):
        # v1 fixture: a worksheet literally named "sheet=4" — `=` isn't ASCII-folded away (it's
        # not alnum/`-`/`.`, so it *would* normally become `_`, but this asserts the header
        # normalizer's behavior on a similarly "special" character for completeness).
        assert to_ascii("sheet=4") == "sheet_4"


class TestNormalizeHeaderRow:
    def test_v1_fixture_col_1_2_3(self):
        assert normalize_header_row(["Col 1", "Col 2", "Col 3"]) == ["Col_1", "Col_2", "Col_3"]

    def test_empty_sheet_single_empty_cell_normalizes_to_no_header(self):
        # Graph still returns a one-cell row (`text: [[""]]`) for a genuinely empty worksheet;
        # v1 special-cases this as "no header at all" rather than a one-column sheet.
        assert normalize_header_row([""]) == []

    def test_zero_cells_normalizes_to_no_header(self):
        assert normalize_header_row([]) == []

    def test_blank_cell_becomes_positional_column_name(self):
        assert normalize_header_row(["a", "", "c"]) == ["a", "column-2", "c"]

    def test_duplicate_names_get_dash_suffixes_in_column_order(self):
        assert normalize_header_row(["id", "id", "id"]) == ["id", "id-1", "id-2"]

    def test_duplicate_after_normalization_still_gets_suffixed(self):
        # "Col!" and "Col?" both normalize to "Col_" before trimming to "Col" — still a collision.
        assert normalize_header_row(["Col", "Col!"]) == ["Col", "Col-1"]

    def test_row_with_multiple_blank_cells_is_not_the_empty_shortcut(self):
        # `len(cells) <= 1` is the empty-sheet shortcut; three blank cells go through normal
        # positional naming instead (v1: `TableHeader::parseColumns`).
        assert normalize_header_row(["", "", ""]) == ["column-1", "column-2", "column-3"]
