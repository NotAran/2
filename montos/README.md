# Registro de montos

Página web para ingresar montos de dinero (ingresos y gastos). Los datos se guardan
en una base de datos SQLite interna (`montos.db`), la página calcula los totales y
muestra los resúmenes, y todo se puede descargar como Excel.

## Qué hace

- **Cargar montos**: tipo (gasto / ingreso), monto, fecha, categoría y descripción.
  El monto acepta `1234.56`, `1.234,56`, `$ 500`, `9.850`, etc.
- **Cálculos**: total de ingresos, total de gastos, balance, % de ahorro,
  gasto promedio y mayor gasto.
- **Resúmenes**: por categoría (con % del gasto total) y por mes.
- **Filtro por período**: elegir un mes para ver solo esos datos.
- **Exportar a Excel** (`.xlsx`) con 4 hojas: *Movimientos*, *Resumen*,
  *Por categoría* y *Por mes*. Las hojas de resumen usan fórmulas (`SUMIF`, etc.),
  así que si se editan los movimientos en Excel, los totales se recalculan.
- Eliminar movimientos.

Los montos se guardan en centavos (enteros) para evitar errores de redondeo.

## Cómo ejecutarlo

```bash
cd montos
pip install -r requirements.txt
python app.py
```

Abrir <http://localhost:5000>.

Variables de entorno opcionales:

| Variable     | Para qué                                   | Por defecto          |
|--------------|--------------------------------------------|----------------------|
| `PORT`       | Puerto del servidor                        | `5000`               |
| `MONTOS_DB`  | Ruta del archivo de la base de datos       | `montos/montos.db`   |
| `SECRET_KEY` | Clave para los mensajes de la sesión       | (cambiarla en producción) |

## Tests

```bash
pip install pytest
python -m pytest
```
