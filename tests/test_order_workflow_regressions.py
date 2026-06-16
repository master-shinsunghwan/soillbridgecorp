from __future__ import annotations

import importlib.util
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side


ROOT = Path(__file__).resolve().parents[1]
MODULE_SETS = [
    ROOT / "scripts",
    ROOT / "_workhub_zip_inspect" / "scripts",
]
TEMPLATE_PATHS = [
    ROOT / "templates" / "vehicle_receipt_template.xlsx",
    ROOT / "_workhub_zip_inspect" / "templates" / "vehicle_receipt_template.xlsx",
]


def load_module(module_path: Path):
    spec = importlib.util.spec_from_file_location(module_path.stem, module_path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def create_delivery_workbook(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Supply"
    worksheet.append(
        [
            "\uc21c\uc11c",
            "\uc8fc\ubb38\uc77c",
            "\uc8fc\ubb38\ubc88\ud638",
            "\uc218\ub839\uc790",
            "\uc218\ub839\uc790\uc5f0\ub77d\ucc98",
            "\uc81c \ud488 \uba85",
            "\uc218\ub7c9",
            "\uc0c1 \uc138 \uc8fc \uc18c",
        ]
    )
    worksheet.append(
        [
            1,
            "2026-06-15 09:30:46",
            "WSO260615-000000001",
            "\ud64d\uae38\ub3d9",
            "010-1111-2222",
            "\ud14c\uc2a4\ud2b8 \uc0c1\ud488",
            2,
            "\uc11c\uc6b8\uc2dc \ud14c\uc2a4\ud2b8\ub85c 1",
        ]
    )
    workbook.save(path)


def create_invoice_workbook(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Invoices"
    worksheet.append(
        [
            "\uc218\ud558\uc778\uba85",
            "\uc1a1\uc7a5\ubc88\ud638",
            "\uc8fc\ubb38\ubc88\ud638",
            "\uc218\ud558\uc778\uae30\ubcf8\uc8fc\uc18c",
        ]
    )
    worksheet.append(["\ud64d\uae38\ub3d9", "1234567890", "ORDER-1", "\uc11c\uc6b8"])
    worksheet.append(["\ud64d\uae38\ub3d9", "1234567890", "ORDER-1", "\uc11c\uc6b8"])
    workbook.save(path)


def create_lotte_source_workbook(path: Path, source_headers: list[str]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Source"
    worksheet.append(source_headers)
    worksheet.append([f"테스트{idx}" for idx, _ in enumerate(source_headers, start=1)])
    workbook.save(path)


def create_lotte_template_workbook(path: Path, template_headers: list[str]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Template"
    worksheet.append(template_headers)
    worksheet.append(["" for _ in template_headers])

    left_border = Side(style="thin", color="000000")
    for column in range(2, 6):
        cell = worksheet.cell(2, column)
        cell.border = Border(left=left_border)

    workbook.save(path)


class OrderWorkflowRegressionTests(unittest.TestCase):
    def test_delivery_summary_accepts_receiver_headers_in_all_script_copies(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            workbook_path = Path(tmp) / "delivery.xlsx"
            create_delivery_workbook(workbook_path)

            for scripts_dir in MODULE_SETS:
                with self.subTest(scripts_dir=scripts_dir):
                    module = load_module(scripts_dir / "delivery_text_summary.py")
                    summary, sheet_names = module.summarize_workbook(workbook_path)

                    self.assertEqual(sheet_names, ["Supply"])
                    self.assertIn("\ud14c\uc2a4\ud2b8 \uc0c1\ud488 - 2\uac1c (1\uac74)", summary)

    def test_invoice_export_preserves_duplicate_invoice_numbers_in_all_script_copies(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            workbook_path = Path(tmp) / "invoice.xlsx"
            create_invoice_workbook(workbook_path)

            for scripts_dir in MODULE_SETS:
                with self.subTest(scripts_dir=scripts_dir):
                    module = load_module(scripts_dir / "invoice_number_exporter.py")
                    rows = module.extract_invoice_rows(workbook_path)

                    self.assertEqual(rows, [("\ud64d\uae38\ub3d9", "1234567890 / 1234567890")])

    def test_lotte_order_export_removes_left_borders_from_columns_b_to_e(self) -> None:
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as tmp:
            temp_dir = Path(tmp)

            for scripts_dir in MODULE_SETS:
                with self.subTest(scripts_dir=scripts_dir):
                    module = load_module(scripts_dir / "lotte_order_form_converter.py")
                    source_path = temp_dir / f"source-{scripts_dir.name}.xlsx"
                    template_path = temp_dir / f"template-{scripts_dir.name}.xlsx"
                    output_dir = temp_dir / f"output-{scripts_dir.name}"

                    create_lotte_source_workbook(source_path, list(module.SOURCE_TO_TEMPLATE.keys()))
                    create_lotte_template_workbook(
                        template_path,
                        list(module.SOURCE_TO_TEMPLATE.values()),
                    )

                    output_path = module.convert_lotte_order_form(
                        source_path,
                        template_path,
                        output_dir,
                        output_date=date(2026, 6, 16),
                    )
                    workbook = load_workbook(output_path)
                    worksheet = workbook.active

                    for column in range(2, 6):
                        self.assertIsNone(
                            worksheet.cell(2, column).border.left.style,
                            f"{worksheet.cell(2, column).coordinate} should not have a left border",
                        )

    def test_vehicle_receipt_template_formatting_matches_print_requirements(self) -> None:
        red_values = {"FFFF0000", "00FF0000", "FF0000"}

        for template_path in TEMPLATE_PATHS:
            with self.subTest(template_path=template_path):
                workbook = load_workbook(template_path)
                worksheet = workbook.active

                for row in range(8, 39):
                    color = worksheet[f"J{row}"].font.color
                    rgb = str(color.rgb).upper() if color and color.type == "rgb" else None
                    self.assertNotIn(rgb, red_values, f"J{row} should not be red")

                b40_color = worksheet["B40"].font.color
                b40_rgb = (
                    str(b40_color.rgb).upper()
                    if b40_color and b40_color.type == "rgb"
                    else None
                )
                self.assertIn(b40_rgb, red_values)

                for row in range(1, 45):
                    self.assertIsNone(
                        worksheet[f"A{row}"].border.left.style,
                        f"A{row} should not have a left border",
                    )


if __name__ == "__main__":
    unittest.main()
