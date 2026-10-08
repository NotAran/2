# Cajas por turno

Página web con usuarios para llevar las cajas de cada turno. Los datos se guardan
en una base de datos SQLite interna (`montos.db`), la página hace los cálculos y
todo se puede descargar en Excel.

## Cómo funciona

- **3 cajas por día**, en este orden:
  - Noche: 23:00 a 07:00 (arranca a las 23:00 del día anterior)
  - Mañana: 07:00 a 15:00
  - Tarde: 15:00 a 23:00

  Al entrar se abre la caja del turno en curso (hora de Argentina; se puede
  cambiar con la variable `MONTOS_TZ`). Se cambia de caja con la fecha, los
  botones de turno o "Caja anterior / Caja siguiente".
- **Cierre del día**: a las 23:00, cuando termina la tarde, las 3 cajas de ese día
  quedan cerradas. Solo el admin puede modificarlas después.
- Cada caja es una planilla de **cuentas** (filas: Mercado, Naranja, Ualá,
  Personal, Brubank, Lemon, Prex, Arq, Binance) por **personas** (columnas:
  Paco, Antonio, Ibra; son solo columnas, no usuarios) y se divide en:
  - **Turno actual**: se suma cada columna (su total se ve en la fila "Total")
    y después se suman los totales de las columnas.
  - **Turno anterior**: el turno actual de la caja inmediatamente anterior
    (0 si esa caja no tuvo montos). Se calcula solo.
- **Cuenta de la caja** (arriba de todo):
  - Resultado = turno anterior − (turno actual + bajada)
  - Final = resultado − saldo
- **Otros datos del turno**: Depósito, Retiro, Bono, Saldo y Bajada, un monto de
  cada uno por caja. No suman en el turno actual. El saldo se calcula solo:
  depósito − retiro.
- Si se corrige una caja, el turno anterior de la siguiente se actualiza solo.
- **Últimas cajas**: lista con turno anterior, turno actual, resultado y final de cada caja.
- **Excel**: descarga la caja elegida (con los totales como fórmulas) y el historial.

## Usuarios y niveles

- La primera vez que se abre la página pide crear la cuenta del **admin**.
- Niveles: **admin**, **encargado** y **cajero**. Por ahora encargado y cajero
  pueden hacer lo mismo: ven y cargan las filas y columnas que el admin les asigna.
- El admin crea usuarios (todos los que quiera), elige su nivel, les cambia el
  nombre o la contraseña, o los elimina. No puede cambiar su propio nivel.
- Cada usuario elige **su nombre** (el que aparece en la página) en "Mi cuenta",
  donde también puede cambiar su contraseña. Para ingresar se usa el usuario.
- El admin agrega o quita filas y columnas. Cada columna puede sumar o restar.
  Si se quita una fila o columna que ya tiene montos, se oculta pero sus montos
  se conservan (se puede reactivar).
- Los totales de la caja son siempre de toda la caja, aunque el usuario vea una parte.
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
