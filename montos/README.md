# Planilla de montos

Página web con usuarios donde se cargan montos de dinero en una planilla de
filas (cuentas) y columnas. Los datos se guardan en una base de datos SQLite
interna (`montos.db`), la página calcula los totales y todo se puede descargar
en Excel.

## Qué hace

- **Admin**: la primera vez que se abre la página pide crear la cuenta del admin.
- **Usuarios**: el admin crea usuarios (todos los que quiera), les cambia la
  contraseña o los elimina.
- **Planilla**: filas iniciales Mercado, Naranja, Ualá, Personal, Brubank, Lemon,
  Prex, Arq y Binance; columnas iniciales Ingresos (suma) y Gastos (resta).
  El admin puede agregar o quitar filas y columnas. Cada columna puede sumar o
  restar en el total.
- **Permisos**: el admin elige qué filas y qué columnas ve y carga cada usuario.
  El admin ve y carga todo.
- **Cálculos**: total por fila, total por columna y total general.
- **Excel**: descarga la planilla visible con los totales como fórmulas.
- Cada celda guarda quién la modificó por última vez y cuándo (se ve al pasar
  el mouse por encima).

Los montos se guardan en centavos (enteros) para evitar errores de redondeo.
Se aceptan formatos como `1.234,56`, `1234.56`, `$ 500`, `9.850` o `-500`.

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
| `SECRET_KEY` | Clave para firmar las sesiones             | se genera sola en `.secret_key` |

## Tests

```bash
pip install pytest
python -m pytest
```
