# facturasPapeleria

Automatiza la captura de productos de facturas (fotos o texto) hacia un Excel `factura_con_iva.xlsx`.

## Uso rápido

1. Instala dependencias:
   ```bash
   pip install -r requirements.txt
   ```
2. Procesa una o varias facturas:
   ```bash
   python facturas_excel.py factura1.jpg factura2.jpg
   ```
3. (Opcional) Compara contra un pedido y genera la hoja `Comparación`:
   ```bash
   python facturas_excel.py factura1.jpg --pedido pedido.csv
   ```

Si un dato no se puede leer, se marca como `REVISAR` y también se imprime en consola la parte no legible.
