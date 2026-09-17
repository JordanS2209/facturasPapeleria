# facturasPapeleria

Este repositorio sirve para organizar facturas de papelería en Excel.

## Archivo generado
- `factura_con_iva.xlsx`
  - Hoja `Factura` con cálculo de IVA 15% por producto y totales.
  - Hoja `Comparación` preparada para validar faltantes contra pedido original.

## Uso rápido
1. Instala dependencia:
   ```bash
   pip install openpyxl
   ```
2. Genera/actualiza el Excel:
   ```bash
   python facturas_excel.py
   ```

## Recomendación para futuras facturas
- Sube fotos claras y completas de cada factura.
- Si quieres detectar faltantes, también sube la lista o foto del pedido original para poder comparar cantidades.
