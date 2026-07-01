from __future__ import annotations

from pathlib import Path


class ExcelComClient:
    """Thin COM wrapper so business logic can stay independent from pywin32."""

    def __init__(self) -> None:
        self._app = None

    def open(self) -> None:
        try:
            import win32com.client  # type: ignore
        except ImportError as exc:
            raise RuntimeError("pywin32 is required for Excel COM automation") from exc

        self._app = win32com.client.Dispatch("Excel.Application")
        self._app.Visible = False
        self._app.DisplayAlerts = False

    def close(self) -> None:
        if self._app is not None:
            self._app.Quit()
            self._app = None

    def copy_workbook(self, template_path: Path, output_path: Path) -> None:
        if self._app is None:
            raise RuntimeError("ExcelComClient is not open")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb = self._app.Workbooks.Open(str(template_path.resolve()))
        try:
            wb.SaveAs(str(output_path.resolve()))
        finally:
            wb.Close(SaveChanges=True)

    def open_workbook(self, workbook_path: Path):
        if self._app is None:
            raise RuntimeError("ExcelComClient is not open")
        return self._app.Workbooks.Open(str(workbook_path.resolve()))

    def export_workbook_to_pdf(self, workbook, output_path: Path) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            selected_sheets = workbook.Windows(1).SelectedSheets
            try:
                selected_sheet_names = [
                    selected_sheets.Item(index).Name for index in range(1, selected_sheets.Count + 1)
                ]
            except Exception:
                selected_sheet_names = [selected_sheets.Name]

            workbook.Worksheets(selected_sheet_names).Select()
            workbook.ActiveSheet.ExportAsFixedFormat(0, str(output_path.resolve()))
        except Exception:
            workbook.ExportAsFixedFormat(0, str(output_path.resolve()))
