"""Run with python -m unittest; optional sample: PAS_SAMPLE_PDF environment variable."""
import ast
import io
import os
import re
import unittest
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pandas as pd

ROOT = Path(__file__).parent
COLUMNS = ["Site", "Order No", "Contract No", "Fleet / Stock No", "Description", "Qty", "On Hire Date", "Weekly Rate", "Net Weekly"]


def load_functions(filename, source=None):
    tree = ast.parse(source or (ROOT / filename).read_text(encoding="utf-8-sig"))
    # Load extraction logic without starting the Streamlit interface.
    nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef)
             and not n.name.startswith("render_")]
    scope = dict(globals(), OUTPUT_COLUMNS=COLUMNS)
    exec(compile(ast.Module(body=nodes, type_ignores=[]), filename, "exec"), scope)
    return scope


class ExtractionTests(unittest.TestCase):
    def test_existing_formats(self):
        fixtures = [
            ("parse_ambrose_live_hire", "0000001234 PAS002 Order No: P123/H456 Delivery Address Yard A\nABC01 Storage container 1 25.00 0% 25.00 N/A 01/02/2026 Weekly"),
            ("parse_mcs_items_on_hire", "Contract No: C123 PO Number: P123/H456\nSite Name: Yard A\n01-Feb-2026 ABC01 Storage container 1"),
            ("parse_equipment_custom", "Site: Yard A\nCAT ABC01 SERIAL01 1 Storage container 01/02/2026 C123 P123/H456 25.00/w"),
            ("parse_current_hires_anchored", "Yard A\nCAT ABC01 SERIAL01 Storage container 1 01/02/2026 HC123 P123/H456 25.00 week"),
            ("parse_generic_order_blocks", "Contract No: C123 PO Number: P123/H456\nSite Name: Yard A\n01/02/2026 ABC01 Storage container 1"),
        ]
        for filename in ["app.py", "app_pas_vendor_pdf_converter.py"]:
            scope = load_functions(filename)
            for parser, text in fixtures:
                with self.subTest(file=filename, parser=parser):
                    rows = scope[parser](text)
                    self.assertEqual(len(rows), 1)
                    self.assertEqual(rows[0]["Order No"], "P123/H456")
                    self.assertEqual(rows[0]["On Hire Date"], "01/02/2026")

    @unittest.skipUnless(os.environ.get("PAS_SAMPLE_PDF"), "Set PAS_SAMPLE_PDF to the What's Out Where sample")
    def test_whats_out_where_sample(self):
        for filename in ["app.py", "app_pas_vendor_pdf_converter.py"]:
            scope = load_functions(filename)
            with open(os.environ["PAS_SAMPLE_PDF"], "rb") as sample:
                frame, warnings = scope["smart_extract_vendor_pdf"](sample)
                self.assertEqual(sample.tell(), 0)
            self.assertEqual(len(frame), 23)
            self.assertEqual(len(frame["Fleet / Stock No"].unique()), 23)
            row = frame[frame["Fleet / Stock No"] == "3057/10-8/S"].iloc[0]
            self.assertEqual(row["Description"], "10FT X 8FT STORAGE CONTAINER with racking + chain rack")
            self.assertEqual(row["On Hire Date"], "29/07/2026")
            self.assertEqual(row["Contract No"], "0000003148")
            wrapped = frame[frame["Fleet / Stock No"] == "2777/10-8/S"].iloc[0]
            self.assertEqual(wrapped["Order No"], "P149/H6404/Peel Hall")
            self.assertTrue((frame["Weekly Rate"] == "").all())
            self.assertTrue((frame["Net Weekly"] == "").all())
            self.assertTrue((frame["Site"] == "").all())


if __name__ == "__main__":
    unittest.main()
