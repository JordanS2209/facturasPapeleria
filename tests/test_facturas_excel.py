import unittest

from facturas_excel import (
    InvoiceItem,
    build_invoice_row,
    normalize_product_name,
    parse_invoice_text,
    update_comparison_sheet,
)
from openpyxl import Workbook


class FacturasExcelTest(unittest.TestCase):
    def test_build_invoice_row_calculates_iva(self):
        item = InvoiceItem(cantidad=2, descripcion="Lapiz", precio_base=10, observaciones=[])
        row = build_invoice_row(item, prices_include_iva=False)

        self.assertEqual(row[2], 10.0)
        self.assertEqual(row[3], 11.5)
        self.assertEqual(row[4], 20.0)
        self.assertEqual(row[5], 3.0)
        self.assertEqual(row[6], 23.0)

    def test_parse_invoice_marks_missing_fields(self):
        items, issues = parse_invoice_text("Producto sin datos")

        self.assertEqual(len(items), 1)
        self.assertTrue(items[0].observaciones)
        self.assertTrue(any("cantidad" in issue.lower() for issue in issues))

    def test_comparison_sheet_statuses(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Facturas"
        ws.append(["Cantidad", "Descripción del producto"])
        ws.append([2, "Lapiz"])
        ws.append([1, "Cuaderno"])

        received = {
            normalize_product_name("Lapiz"): 2,
            normalize_product_name("Cuaderno"): 1,
        }
        requested = [("Lapiz", 2), ("Cuaderno", 2), ("Borrador", 1)]

        update_comparison_sheet(wb, requested=requested, received=received)
        comp = wb["Comparación"]

        rows = list(comp.iter_rows(min_row=2, values_only=True))
        self.assertEqual(rows[0][4], "Completo")
        self.assertEqual(rows[1][4], "Parcial")
        self.assertEqual(rows[2][4], "Faltante")


if __name__ == "__main__":
    unittest.main()
