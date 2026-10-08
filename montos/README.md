# Cajas por turno

Página web con usuarios para llevar las cajas de cada turno. Los datos se guardan
en una base de datos SQLite interna (`montos.db`), la página hace los cálculos y
todo se puede descargar en Excel.

## Cómo funciona

- **3 cajas por día**, en este orden: Noche → Mañana → Tarde (y después la Noche
  del día siguiente). Se elige la caja con la fecha y los botones de turno, o con
  "Caja anterior / Caja siguiente".
- Cada caja es una planilla de **cuentas** (filas: Mercado, Naranja, Ualá,
  Personal, Brubank, Lemon, Prex, Arq, Binance) por **personas** (columnas:
  Paco, Antonio, Ibra; son solo columnas, no usuarios) y se divide en:
  - **Turno anterior**: un solo número, el total con el que cerró la caja
    anterior (todas las cuentas y columnas, acumulado). Se calcula solo.
  - **Turno actual**: los montos que se cargan en cada columna, con el total
    por cuenta y por columna.
  - **Total caja** = turno anterior + turno actual. Pasa a ser el turno anterior
    de la caja siguiente.
- **Otros datos del turno**: Depósito, Retiro, Bono, Saldo y Bajada, un monto de cada uno por caja.
  Se guardan con la caja pero no entran en las sumas.
- Si se corrige una caja vieja, los saldos de las cajas siguientes se actualizan solos.
- **Últimas cajas**: lista con el movimiento de cada turno y el saldo al cierre.
- **Excel**: descarga la caja elegida (con los totales como fórmulas) y el historial.

## Usuarios y admin

- La primera vez que se abre la página pide crear la cuenta del **admin**.
- El admin crea usuarios (todos los que quiera), les cambia la contraseña o los elimina.
- El admin agrega o quita filas y columnas. Cada columna puede sumar o restar.
  Si se quita una fila o columna que ya tiene montos, se oculta pero sus montos
  siguen contando en los saldos (se puede reactivar).
- El admin elige qué filas y columnas ve y carga cada usuario. Los totales de
  la caja (turno anterior, turno actual y total) son siempre de toda la caja.
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
