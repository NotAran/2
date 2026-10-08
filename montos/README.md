# Cajas por turno

Página web con usuarios para llevar las cajas de cada turno. Los datos se guardan
en una base de datos SQLite interna (`montos.db`), la página hace los cálculos y
todo se puede descargar en Excel.

## Cómo funciona

- **Cajas**: hay varias cajas (al empezar, Caja 1 y Caja 2). Se eligen con las
  pestañas de arriba. Cada caja tiene sus propias columnas (personas), montos,
  panel, observaciones y cuentas: una caja no ve ni altera los cálculos de otra.
- **3 turnos por día**, en este orden:
  - Noche: 23:00 a 07:00 (arranca a las 23:00 del día anterior)
  - Mañana: 07:00 a 15:00
  - Tarde: 15:00 a 23:00

  Al entrar se abre el turno en curso (hora de Argentina; se puede cambiar con
  la variable `MONTOS_TZ`). Se cambia de turno con la fecha, los botones de
  turno o "Turno anterior / Turno siguiente".
- **Cierre del día**: a las 23:00, cuando termina la tarde, los turnos de ese día
  quedan cerrados. Solo el admin puede modificarlos después.
- Cada caja y turno es una planilla de **billeteras** (filas: Mercado, Naranja,
  Ualá, Personal, Brubank, Lemon, Prex, Arq, Binance; las mismas para todas las
  cajas) por **columnas** de personas de esa caja (al empezar, la Caja 1 tiene
  Paco, Antonio e Ibra, y la Caja 2 ninguna).
  - **Turno actual**: se suma cada columna (su total se ve en la fila "Total")
    y después se suman los totales de las columnas.
  - **Turno anterior**: el turno actual de la misma caja en el turno
    inmediatamente anterior (0 si no tuvo montos). Se calcula solo.
- **Cuentas de la caja** (debajo de la grilla, encima del "Panel"):
  - Resultado = turno anterior − (turno actual + bajada)
  - Final = resultado + saldo
  - Total final = final − bono
- **Panel**: Depósito, Retiro, Bono, Saldo y Bajada, un monto de cada uno por
  caja y turno. No suman en el turno actual. El saldo se calcula solo:
  depósito − retiro.
- **Observaciones**: un texto libre por caja y turno, debajo del panel.
- **Últimos turnos**: lista con turno anterior, turno actual, resultado, final y
  total final de cada turno de la caja.
- **Excel**: descarga la caja y el turno elegidos (con los totales como
  fórmulas) y el historial de esa caja.

## Usuarios y niveles

- La primera vez que se abre la página pide crear la cuenta del **admin**.
- Niveles: **admin**, **encargado** y **cajero**. Por ahora encargado y cajero
  pueden hacer lo mismo.
- El admin asigna a cada usuario **qué cajas ve y edita**. Un usuario no ve las
  cajas que no tiene asignadas. El admin ve y edita todas.
- El admin crea usuarios (todos los que quiera), elige su nivel, les cambia el
  nombre o la contraseña, o los elimina. No puede cambiar su propio nivel.
- Cada usuario elige **su nombre** (el que aparece en la página) en "Mi cuenta",
  donde también puede cambiar su contraseña. Para ingresar se usa el usuario.
- El admin crea y renombra cajas, agrega o quita billeteras y columnas de cada
  caja, y cambia su orden con las flechas ↑ ↓. Cada columna puede sumar o
  restar. Si se quita una billetera o columna que ya tiene montos, se oculta
  pero sus montos se conservan (se puede reactivar).
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
