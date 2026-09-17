# facturasPapeleria

Este repositorio genera un archivo Excel inicial `factura_con_iva.xlsx` para capturar facturas con cálculo de IVA (15%) y una hoja de comparación contra pedido.

## Requisitos

- Python 3
- `openpyxl`

Instalación:

```bash
pip install openpyxl
```

## Generar el Excel

Desde la raíz del repositorio:

```bash
python generar_factura_excel.py
```

Se generará/actualizará:

- `factura_con_iva.xlsx`

## Uso para futuras facturas

1. Reemplaza en `generar_factura_excel.py` los productos por los datos visibles de las fotos.
2. Si un dato no es legible, deja `REVISAR` en Observaciones (no inventar datos).
3. Verifica la sección de comparación de totales calculados vs. totales impresos.
4. Completa `COMPARACION_ROWS` con la lista solicitada para que la hoja `Comparación` marque completos, parciales o faltantes.
5. Si la factura usa cajas/paquetes, conserva esa unidad y aclara en Observaciones cualquier diferencia de unidad.
