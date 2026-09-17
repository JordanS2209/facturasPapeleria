from openpyxl import Workbook
from openpyxl.styles import Font

OUTPUT_FILE = "factura_con_iva.xlsx"

# Datos legibles de las fotos compartidas en la conversación.
PRODUCTOS = [
    {
        "cantidad": 30,
        "descripcion": "C.A. Embalaje Strong Cristal 48x40",
        "precio_unitario_sin_iva": 0.32,
        "valor_total_sin_iva": 9.66,
        "observaciones": "",
    },
    {
        "cantidad": 25,
        "descripcion": "Membrete Lamina Creativ X15",
        "precio_unitario_sin_iva": 0.08,
        "valor_total_sin_iva": 2.07,
        "observaciones": "REVISAR: descripción/transcripción poco legible en foto.",
    },
    {
        "cantidad": 10,
        "descripcion": "Clip Cromado Lancer 50mm X50",
        "precio_unitario_sin_iva": 0.24,
        "valor_total_sin_iva": 2.47,
        "observaciones": "",
    },
]

TOTALES_IMPRESOS = {
    "subtotal": 287.39,
    "iva": 43.11,
    "total": 330.50,
}


def crear_hoja_factura(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "Factura"

    encabezados = [
        "Cantidad",
        "Descripción del producto",
        "Precio unitario sin IVA",
        "Precio unitario con IVA",
        "Valor total sin IVA",
        "IVA del 15%",
        "Total con IVA",
        "Observaciones",
    ]

    ws.append(encabezados)
    for col in "ABCDEFGH":
        ws[f"{col}1"].font = Font(bold=True)

    fila_inicio = 2
    for i, p in enumerate(PRODUCTOS, start=fila_inicio):
        ws[f"A{i}"] = p["cantidad"]
        ws[f"B{i}"] = p["descripcion"]
        ws[f"C{i}"] = p["precio_unitario_sin_iva"]
        ws[f"D{i}"] = f"=ROUND(C{i}*1.15,2)"
        ws[f"E{i}"] = p["valor_total_sin_iva"]
        ws[f"F{i}"] = f"=ROUND(E{i}*0.15,2)"
        ws[f"G{i}"] = f"=ROUND(E{i}+F{i},2)"
        ws[f"H{i}"] = p["observaciones"]

    fila_fin = fila_inicio + len(PRODUCTOS) - 1
    fila_subtotal = fila_fin + 2

    ws[f"D{fila_subtotal}"] = "Subtotal"
    ws[f"E{fila_subtotal}"] = f"=ROUND(SUM(E{fila_inicio}:E{fila_fin}),2)"

    ws[f"D{fila_subtotal + 1}"] = "IVA total"
    ws[f"F{fila_subtotal + 1}"] = f"=ROUND(SUM(F{fila_inicio}:F{fila_fin}),2)"

    ws[f"D{fila_subtotal + 2}"] = "Total general"
    ws[f"G{fila_subtotal + 2}"] = f"=ROUND(SUM(G{fila_inicio}:G{fila_fin}),2)"

    ws[f"D{fila_subtotal + 4}"] = "Totales impresos (factura)"
    ws[f"E{fila_subtotal + 4}"] = TOTALES_IMPRESOS["subtotal"]
    ws[f"F{fila_subtotal + 4}"] = TOTALES_IMPRESOS["iva"]
    ws[f"G{fila_subtotal + 4}"] = TOTALES_IMPRESOS["total"]

    ws[f"D{fila_subtotal + 5}"] = "Diferencia (calculado - impreso)"
    ws[f"E{fila_subtotal + 5}"] = f"=ROUND(E{fila_subtotal}-E{fila_subtotal + 4},2)"
    ws[f"F{fila_subtotal + 5}"] = f"=ROUND(F{fila_subtotal + 1}-F{fila_subtotal + 4},2)"
    ws[f"G{fila_subtotal + 5}"] = f"=ROUND(G{fila_subtotal + 2}-G{fila_subtotal + 4},2)"
    ws[f"H{fila_subtotal + 5}"] = "REVISAR: el total calculado difiere del impreso; faltan líneas no legibles/visibles en la transcripción actual."

    for col in "DEFG":
        for row in range(2, fila_subtotal + 6):
            ws[f"{col}{row}"].number_format = "0.00"

    anchos = {
        "A": 10,
        "B": 45,
        "C": 22,
        "D": 22,
        "E": 20,
        "F": 14,
        "G": 16,
        "H": 70,
    }
    for col, width in anchos.items():
        ws.column_dimensions[col].width = width


def crear_hoja_comparacion(wb: Workbook) -> None:
    ws = wb.create_sheet("Comparación")
    encabezados = [
        "Producto solicitado",
        "Cantidad solicitada",
        "Cantidad recibida",
        "Diferencia",
        "Estado",
        "Observaciones",
    ]
    ws.append(encabezados)
    for col in "ABCDEF":
        ws[f"{col}1"].font = Font(bold=True)

    ws["A2"] = "(Pendiente de completar)"
    ws["F2"] = "Se requiere la lista/foto original del pedido para calcular faltantes."

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 65


def main() -> None:
    wb = Workbook()
    crear_hoja_factura(wb)
    crear_hoja_comparacion(wb)
    wb.save(OUTPUT_FILE)


if __name__ == "__main__":
    main()
