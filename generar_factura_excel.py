from pathlib import Path

from openpyxl import Workbook

OUTPUT_FILE = Path(__file__).resolve().parent / "factura_con_iva.xlsx"
IVA = 0.15

# Datos visibles en el contexto de conversación.
# No se completan más líneas para evitar inventar datos.
FACTURA_ROWS = [
    {
        "cantidad": 30,
        "descripcion": "C.A. Embalaje Strong Cristal 48x40",
        "precio_sin_iva": 0.32,
        "valor_total_sin_iva": 9.66,
        "observaciones": "REVISAR: valor total no coincide exactamente con cantidad x precio unitario.",
    },
    {
        "cantidad": 25,
        "descripcion": "Membrete Lamina Creativ X15",
        "precio_sin_iva": 0.08,
        "valor_total_sin_iva": 2.07,
        "observaciones": "REVISAR: verificar ortografía/lectura exacta del producto en la foto.",
    },
    {
        "cantidad": 10,
        "descripcion": "Clip Cromado Lancer 50mm X50",
        "precio_sin_iva": 0.24,
        "valor_total_sin_iva": 2.47,
        "observaciones": "",
    },
]

TOTALES_IMPRESOS = {
    "subtotal": 287.39,
    "iva": 43.11,
    "total": 330.50,
}


def build_factura_sheet(workbook: Workbook) -> None:
    ws = workbook.active
    ws.title = "Factura"

    headers = [
        "Cantidad",
        "Descripción del producto",
        "Precio unitario sin IVA",
        "Precio unitario con IVA",
        "Valor total sin IVA",
        "IVA del 15%",
        "Total con IVA",
        "Observaciones",
    ]
    ws.append(headers)

    for row_index, item in enumerate(FACTURA_ROWS, start=2):
        ws.cell(row=row_index, column=1, value=item["cantidad"])
        ws.cell(row=row_index, column=2, value=item["descripcion"])
        ws.cell(row=row_index, column=3, value=item["precio_sin_iva"])
        ws.cell(row=row_index, column=4, value=f"=ROUND(C{row_index}*(1+{IVA}),2)")
        ws.cell(row=row_index, column=5, value=item["valor_total_sin_iva"])
        ws.cell(row=row_index, column=6, value=f"=ROUND(E{row_index}*{IVA},2)")
        ws.cell(row=row_index, column=7, value=f"=ROUND(E{row_index}+F{row_index},2)")
        ws.cell(row=row_index, column=8, value=item["observaciones"])

    total_row = len(FACTURA_ROWS) + 2
    ws.cell(row=total_row, column=4, value="Subtotal")
    ws.cell(row=total_row, column=5, value=f"=ROUND(SUM(E2:E{total_row-1}),2)")

    ws.cell(row=total_row + 1, column=4, value="IVA total")
    ws.cell(row=total_row + 1, column=5, value=f"=ROUND(SUM(F2:F{total_row-1}),2)")

    ws.cell(row=total_row + 2, column=4, value="Total general")
    ws.cell(row=total_row + 2, column=5, value=f"=ROUND(SUM(G2:G{total_row-1}),2)")

    ws.cell(row=1, column=10, value="Totales impresos en factura")
    ws.cell(row=2, column=9, value="Subtotal impreso")
    ws.cell(row=2, column=10, value=TOTALES_IMPRESOS["subtotal"])
    ws.cell(row=3, column=9, value="IVA impreso")
    ws.cell(row=3, column=10, value=TOTALES_IMPRESOS["iva"])
    ws.cell(row=4, column=9, value="Total impreso")
    ws.cell(row=4, column=10, value=TOTALES_IMPRESOS["total"])

    ws.cell(row=2, column=11, value="Diferencia")
    ws.cell(row=2, column=12, value=f"=ROUND(E{total_row}-J2,2)")
    ws.cell(row=3, column=12, value=f"=ROUND(E{total_row+1}-J3,2)")
    ws.cell(row=4, column=12, value=f"=ROUND(E{total_row+2}-J4,2)")

    ws.cell(
        row=total_row + 4,
        column=1,
        value="REVISAR: los totales calculados no coinciden con los impresos; faltan líneas por transcribir o validar en las fotos.",
    )

    widths = {
        "A": 12,
        "B": 42,
        "C": 24,
        "D": 24,
        "E": 20,
        "F": 14,
        "G": 16,
        "H": 64,
        "I": 20,
        "J": 20,
        "K": 12,
        "L": 14,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in range(2, total_row + 3):
        for col in (3, 4, 5, 6, 7, 10, 12):
            ws.cell(row=row, column=col).number_format = "0.00"


def build_comparacion_sheet(workbook: Workbook) -> None:
    ws = workbook.create_sheet(title="Comparación")
    ws.append(
        [
            "Producto solicitado",
            "Cantidad solicitada",
            "Cantidad recibida",
            "Diferencia",
            "Estado",
            "Observaciones",
        ]
    )
    ws.append(
        [
            "",
            "",
            "",
            "",
            "",
            "Se necesita la lista del pedido original para detectar faltantes.",
        ]
    )

    widths = {
        "A": 40,
        "B": 20,
        "C": 20,
        "D": 14,
        "E": 14,
        "F": 60,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def main() -> None:
    wb = Workbook()
    build_factura_sheet(wb)
    build_comparacion_sheet(wb)
    wb.save(OUTPUT_FILE)
    print(f"Archivo generado: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
