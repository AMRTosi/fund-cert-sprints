"""Tests for the sheet duplicator service."""

from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color

from sprint_cert_automation.services.sheet_duplicator import (
    COST_FIRST_ROW,
    COST_LAST_ROW,
    DAY_LETTER_ROW,
    DAY_NUMBER_ROW,
    FIRST_DAY_COL,
    REVENUE_FIRST_ROW,
    REVENUE_LAST_ROW,
    _generate_calendar,
    _is_gray_fill,
    _replace_column_in_formula,
    duplicate_sheet,
    fill_empty_cost_cells,
    remove_gray_fills,
    update_calendar,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _create_template_workbook(path: Path, sheet_name: str = "Template_Mes") -> Path:
    """Create a minimal workbook that mirrors the Template_Mes structure."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    # Row 5: day numbers 1-31
    for day in range(1, 32):
        ws.cell(row=DAY_NUMBER_ROW, column=FIRST_DAY_COL + day - 1, value=day)

    # Row 6: day letters for a generic month (October-like: starts Thursday)
    letters = ["J", "V", "S", "D", "L", "M", "X", "J", "V", "S", "D",
               "L", "M", "X", "J", "V", "S", "D", "L", "M", "X",
               "J", "V", "S", "D", "L", "M", "X", "J", "V", "S"]
    for i, letter in enumerate(letters):
        ws.cell(row=DAY_LETTER_ROW, column=FIRST_DAY_COL + i, value=letter)

    # Header cells (row 6 cols A-K)
    ws.cell(row=6, column=1, value="Target")
    ws.cell(row=6, column=2, value="Nombre y Apellidos")

    # AT3 = 9 (hours config)
    ws.cell(row=3, column=46, value=9)

    # Rows 7-9: sample cost formulas (3 employees)
    for row in range(COST_FIRST_ROW, COST_FIRST_ROW + 3):
        ws.cell(row=row, column=1, value=1)  # Target
        ws.cell(row=row, column=2, value=f"Employee {row - 6}")
        for day in range(1, 32):
            col = FIRST_DAY_COL + day - 1
            col_letter = openpyxl.utils.get_column_letter(col)
            ws.cell(
                row=row, column=col,
                value=f'=IF(OR({col_letter}$6="S", {col_letter}$6="D"), "",IF({col_letter}$6="V", 6.5, $AT$3))',
            )

    # Rows 41-43: sample revenue formulas
    for offset in range(3):
        cost_row = COST_FIRST_ROW + offset
        rev_row = REVENUE_FIRST_ROW + offset
        ws.cell(row=rev_row, column=2, value=f"Employee {offset + 1}")
        for day in range(1, 32):
            col = FIRST_DAY_COL + day - 1
            col_letter = openpyxl.utils.get_column_letter(col)
            ws.cell(
                row=rev_row, column=col,
                value=f'=IF({col_letter}{cost_row}=9,8.5,IF({col_letter}{cost_row}=7,7.5,IF({col_letter}{cost_row}=6.5,IF($AT$3=9,8.5,7.5),"")))',
            )

    wb.save(str(path))
    return path


def _add_gray_cells(path: Path, sheet_name: str, cells: list[tuple[int, int]]) -> None:
    """Add gray fill to specific cells and clear their value."""
    wb = openpyxl.load_workbook(str(path))
    ws = wb[sheet_name]
    gray_fill = PatternFill(
        patternType="solid",
        fgColor=Color(theme=2, tint=-0.0999786370433668),
    )
    for row, col in cells:
        cell = ws.cell(row=row, column=col)
        cell.fill = gray_fill
        cell.value = None
    wb.save(str(path))


# ---------------------------------------------------------------------------
# Calendar generation tests
# ---------------------------------------------------------------------------

class TestGenerateCalendar:
    def test_february_2026_has_28_days(self):
        cal = _generate_calendar(2026, 2)
        assert len(cal) == 28
        # Feb 1 2026 is a Sunday
        assert cal[0] == (1, "D")
        assert cal[1] == (2, "L")
        assert cal[27] == (28, "S")

    def test_february_2028_leap_year_has_29_days(self):
        cal = _generate_calendar(2028, 2)
        assert len(cal) == 29

    def test_october_2026_has_31_days(self):
        cal = _generate_calendar(2026, 10)
        assert len(cal) == 31
        # Oct 1 2026 is a Thursday
        assert cal[0] == (1, "J")

    def test_november_2026_has_30_days(self):
        cal = _generate_calendar(2026, 11)
        assert len(cal) == 30
        # Nov 1 2026 is a Sunday
        assert cal[0] == (1, "D")

    def test_weekday_letters_are_spanish(self):
        cal = _generate_calendar(2026, 10)
        # Oct 5 2026 is Monday
        assert cal[4] == (5, "L")
        # Oct 7 2026 is Wednesday
        assert cal[6] == (7, "X")


# ---------------------------------------------------------------------------
# Calendar update tests
# ---------------------------------------------------------------------------

class TestUpdateCalendar:
    def test_updates_day_numbers_and_letters(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb.active

        num_days = update_calendar(ws, 2026, 11)  # November 2026

        assert num_days == 30

        # Check day 1
        assert ws.cell(row=DAY_NUMBER_ROW, column=FIRST_DAY_COL).value == 1
        # Nov 1 2026 is Sunday
        assert ws.cell(row=DAY_LETTER_ROW, column=FIRST_DAY_COL).value == "D"

        # Check day 30
        col_30 = FIRST_DAY_COL + 29
        assert ws.cell(row=DAY_NUMBER_ROW, column=col_30).value == 30

        # Day 31 should be cleared
        col_31 = FIRST_DAY_COL + 30
        assert ws.cell(row=DAY_NUMBER_ROW, column=col_31).value is None
        assert ws.cell(row=DAY_LETTER_ROW, column=col_31).value is None

    def test_clears_data_for_excess_days_february(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb.active

        num_days = update_calendar(ws, 2026, 2)  # February 2026 = 28 days

        assert num_days == 28

        # Days 29, 30, 31 should be cleared
        for excess_day in [29, 30, 31]:
            col = FIRST_DAY_COL + excess_day - 1
            assert ws.cell(row=DAY_NUMBER_ROW, column=col).value is None
            # Cost rows should be cleared too
            assert ws.cell(row=COST_FIRST_ROW, column=col).value is None

    def test_31_day_month_keeps_all_columns(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb.active

        num_days = update_calendar(ws, 2026, 10)  # October 2026 = 31 days

        assert num_days == 31
        col_31 = FIRST_DAY_COL + 30
        assert ws.cell(row=DAY_NUMBER_ROW, column=col_31).value == 31


# ---------------------------------------------------------------------------
# Gray fill tests
# ---------------------------------------------------------------------------

class TestRemoveGrayFills:
    def test_removes_gray_fills_from_cost_area(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        gray_cells = [
            (COST_FIRST_ROW, FIRST_DAY_COL + 1),  # day 2
            (COST_FIRST_ROW, FIRST_DAY_COL + 4),  # day 5
            (COST_FIRST_ROW + 1, FIRST_DAY_COL + 8),  # day 9
        ]
        _add_gray_cells(wb_path, "Template_Mes", gray_cells)

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb.active

        cleared = remove_gray_fills(ws, 31)

        assert cleared == 3
        for row, col in gray_cells:
            assert not _is_gray_fill(ws.cell(row=row, column=col))

    def test_does_not_touch_non_gray_cells(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb.active

        cleared = remove_gray_fills(ws, 31)

        assert cleared == 0


# ---------------------------------------------------------------------------
# Fill empty cost cells tests
# ---------------------------------------------------------------------------

class TestFillEmptyCostCells:
    def test_fills_empty_weekday_cells_with_formula(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        # Add gray cells (which clears value) then remove fills
        gray_cells = [
            (COST_FIRST_ROW, FIRST_DAY_COL + 4),  # day 5, should be weekday
        ]
        _add_gray_cells(wb_path, "Template_Mes", gray_cells)

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb.active

        # Set day 5 to a weekday letter (Monday)
        ws.cell(row=DAY_LETTER_ROW, column=FIRST_DAY_COL + 4, value="L")

        remove_gray_fills(ws, 31)
        filled = fill_empty_cost_cells(ws, 31)

        assert filled == 1
        cell = ws.cell(row=COST_FIRST_ROW, column=FIRST_DAY_COL + 4)
        assert cell.value is not None
        assert cell.value.startswith("=")

    def test_does_not_fill_weekend_cells(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb.active

        # Manually clear a weekend cell
        ws.cell(row=DAY_LETTER_ROW, column=FIRST_DAY_COL + 2).value = "S"
        ws.cell(row=COST_FIRST_ROW, column=FIRST_DAY_COL + 2).value = None

        filled = fill_empty_cost_cells(ws, 31)

        # Should not fill cells on Saturday
        cell = ws.cell(row=COST_FIRST_ROW, column=FIRST_DAY_COL + 2)
        assert cell.value is None

    def test_does_not_overwrite_existing_formulas(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb.active

        original_value = ws.cell(row=COST_FIRST_ROW, column=FIRST_DAY_COL).value
        filled = fill_empty_cost_cells(ws, 31)

        assert ws.cell(row=COST_FIRST_ROW, column=FIRST_DAY_COL).value == original_value


# ---------------------------------------------------------------------------
# Formula column replacement tests
# ---------------------------------------------------------------------------

class TestReplaceColumnInFormula:
    def test_replaces_simple_column_reference(self):
        formula = '=IF(OR(M$6="S", M$6="D"), "",IF(M$6="V", 6.5, $AT$3))'
        result = _replace_column_in_formula(formula, "M", "N")
        assert result == '=IF(OR(N$6="S", N$6="D"), "",IF(N$6="V", 6.5, $AT$3))'

    def test_replaces_two_letter_column(self):
        formula = '=IF(OR(AO$6="S", AO$6="D"), "",IF(AO$6="V", 6.5, $AT$3))'
        result = _replace_column_in_formula(formula, "AO", "AP")
        assert result == '=IF(OR(AP$6="S", AP$6="D"), "",IF(AP$6="V", 6.5, $AT$3))'

    def test_does_not_replace_dollar_prefixed_refs(self):
        formula = '=IF(M$6="V", 6.5, $AT$3)'
        result = _replace_column_in_formula(formula, "M", "N")
        # $AT$3 should remain unchanged
        assert "$AT$3" in result

    def test_same_column_returns_unchanged(self):
        formula = '=IF(M$6="S", "", 8)'
        result = _replace_column_in_formula(formula, "M", "M")
        assert result == formula


# ---------------------------------------------------------------------------
# Gray fill detection tests
# ---------------------------------------------------------------------------

class TestIsGrayFill:
    def test_detects_gray_fill(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        cell = ws.cell(row=1, column=1)
        cell.fill = PatternFill(
            patternType="solid",
            fgColor=Color(theme=2, tint=-0.0999786370433668),
        )
        assert _is_gray_fill(cell) is True

    def test_rejects_no_fill(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        cell = ws.cell(row=1, column=1)
        assert _is_gray_fill(cell) is False

    def test_rejects_non_gray_solid_fill(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        cell = ws.cell(row=1, column=1)
        cell.fill = PatternFill(patternType="solid", fgColor="FF0000")
        assert _is_gray_fill(cell) is False


# ---------------------------------------------------------------------------
# End-to-end duplicate_sheet tests
# ---------------------------------------------------------------------------

class TestDuplicateSheet:
    def test_dry_run_does_not_modify_file(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)
        original_size = wb_path.stat().st_size

        result = duplicate_sheet(
            workbook_path=wb_path,
            source_sheet_name="Template_Mes",
            new_sheet_name="FY27_dic",
            year=2026,
            month=12,
            dry_run=True,
        )

        assert result.new_sheet == "FY27_dic"
        assert result.days_in_month == 31
        # File should not have been modified
        wb = openpyxl.load_workbook(str(wb_path))
        assert "FY27_dic" not in wb.sheetnames

    def test_creates_new_sheet_with_correct_name(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        result = duplicate_sheet(
            workbook_path=wb_path,
            source_sheet_name="Template_Mes",
            new_sheet_name="FY27_dic",
            year=2026,
            month=12,
        )

        wb = openpyxl.load_workbook(str(wb_path))
        assert "FY27_dic" in wb.sheetnames
        assert result.new_sheet == "FY27_dic"

    def test_new_sheet_positioned_after_source(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        # Add another sheet after Template_Mes
        wb = openpyxl.load_workbook(str(wb_path))
        wb.create_sheet("ExtraSheet")
        wb.save(str(wb_path))

        duplicate_sheet(
            workbook_path=wb_path,
            source_sheet_name="Template_Mes",
            new_sheet_name="FY27_dic",
            year=2026,
            month=12,
        )

        wb = openpyxl.load_workbook(str(wb_path))
        names = wb.sheetnames
        src_idx = names.index("Template_Mes")
        new_idx = names.index("FY27_dic")
        assert new_idx == src_idx + 1

    def test_calendar_updated_for_target_month(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        duplicate_sheet(
            workbook_path=wb_path,
            source_sheet_name="Template_Mes",
            new_sheet_name="FY27_dic",
            year=2026,
            month=12,
        )

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb["FY27_dic"]

        # Dec 1 2026 is Tuesday
        assert ws.cell(row=DAY_NUMBER_ROW, column=FIRST_DAY_COL).value == 1
        assert ws.cell(row=DAY_LETTER_ROW, column=FIRST_DAY_COL).value == "M"

    def test_february_clears_excess_days(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        duplicate_sheet(
            workbook_path=wb_path,
            source_sheet_name="Template_Mes",
            new_sheet_name="FY26_feb",
            year=2026,
            month=2,
        )

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb["FY26_feb"]

        # Day 29 should be cleared
        col_29 = FIRST_DAY_COL + 28
        assert ws.cell(row=DAY_NUMBER_ROW, column=col_29).value is None
        assert ws.cell(row=COST_FIRST_ROW, column=col_29).value is None

    def test_gray_fills_removed_and_formulas_filled(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        # Add gray cells on weekdays
        gray_cells = [
            (COST_FIRST_ROW, FIRST_DAY_COL + 4),  # day 5
        ]
        _add_gray_cells(wb_path, "Template_Mes", gray_cells)

        duplicate_sheet(
            workbook_path=wb_path,
            source_sheet_name="Template_Mes",
            new_sheet_name="FY27_oct",
            year=2026,
            month=10,
        )

        wb = openpyxl.load_workbook(str(wb_path))
        ws = wb["FY27_oct"]

        # Oct 5 2026 is Monday (L) - should have formula filled
        cell = ws.cell(row=COST_FIRST_ROW, column=FIRST_DAY_COL + 4)
        day_letter = ws.cell(row=DAY_LETTER_ROW, column=FIRST_DAY_COL + 4).value
        if day_letter not in ("S", "D"):
            assert cell.value is not None
            assert str(cell.value).startswith("=")

    def test_raises_on_missing_source_sheet(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        with pytest.raises(ValueError, match="not found"):
            duplicate_sheet(
                workbook_path=wb_path,
                source_sheet_name="NonExistent",
                new_sheet_name="FY27_dic",
                year=2026,
                month=12,
            )

    def test_raises_on_existing_target_sheet(self, tmp_path):
        wb_path = tmp_path / "test.xlsx"
        _create_template_workbook(wb_path)

        with pytest.raises(ValueError, match="already exists"):
            duplicate_sheet(
                workbook_path=wb_path,
                source_sheet_name="Template_Mes",
                new_sheet_name="Template_Mes",
                year=2026,
                month=12,
            )


# ---------------------------------------------------------------------------
# CLI integration test for duplicate-sheet
# ---------------------------------------------------------------------------

class TestDuplicateSheetCli:
    def test_cli_duplicate_sheet_dry_run(self, monkeypatch, tmp_path):
        from sprint_cert_automation import cli

        captured = {}

        def fake_duplicate(**kwargs):
            captured.update(kwargs)

            class Result:
                workbook_path = kwargs["forecast_path"]
                source_sheet = kwargs["source_sheet"]
                new_sheet = kwargs["new_sheet"]
                year = kwargs["year"]
                month = kwargs["month"]
                days_in_month = 31

            return Result()

        monkeypatch.setattr(cli, "duplicate_period_sheet", fake_duplicate)

        exit_code = cli.main([
            "duplicate-sheet",
            "--forecast", "forecast.xlsx",
            "--source", "Template_Mes",
            "--target", "FY27_dic",
            "--year", "2026",
            "--month", "12",
            "--dry-run",
        ])

        assert exit_code == 0
        assert captured["source_sheet"] == "Template_Mes"
        assert captured["new_sheet"] == "FY27_dic"
        assert captured["year"] == 2026
        assert captured["month"] == 12
        assert captured["dry_run"] is True
