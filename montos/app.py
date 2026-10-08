"""Registro de montos de dinero.

Una pequeña aplicación web donde el usuario ingresa montos (ingresos y gastos).
Los datos se guardan en una base de datos SQLite interna, se calculan totales y
resúmenes, y se pueden exportar a un archivo Excel (.xlsx).
"""

import io
import os
import sqlite3
from datetime import date
from decimal import Decimal, InvalidOperation

from flask import Flask, abort, flash, g, redirect, render_template, request, send_file, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TIPOS = ("ingreso", "gasto")
FORMATO_MONEDA = '"$"#,##0.00;[Red]-"$"#,##0.00'

app = Flask(__name__)
app.config.from_mapping(
    SECRET_KEY=os.environ.get("SECRET_KEY", "cambiar-esta-clave"),
    DATABASE=os.environ.get("MONTOS_DB", os.path.join(BASE_DIR, "montos.db")),
)


# --- Base de datos -----------------------------------------------------------

ESQUEMA = """
CREATE TABLE IF NOT EXISTS movimientos (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha       TEXT    NOT NULL,              -- AAAA-MM-DD
    descripcion TEXT    NOT NULL DEFAULT '',
    categoria   TEXT    NOT NULL DEFAULT 'General',
    tipo        TEXT    NOT NULL CHECK (tipo IN ('ingreso', 'gasto')),
    centavos    INTEGER NOT NULL CHECK (centavos > 0),  -- monto en centavos, sin decimales flotantes
    creado      TEXT    NOT NULL DEFAULT CURRENT_TIMESTAMP
);
"""


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(app.config["DATABASE"])
        g.db.row_factory = sqlite3.Row
        g.db.executescript(ESQUEMA)
    return g.db


@app.teardown_appcontext
def cerrar_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


# --- Utilidades --------------------------------------------------------------

def parsear_monto(texto):
    """Convierte texto como '1.234,56', '1234.56' o '$ 500' a centavos (int).

    Si aparecen punto y coma, el último que aparece es el separador decimal.
    Si aparece uno solo, es de miles cuando se repite o lo siguen exactamente
    3 dígitos ('9.850' = 9850, '1,000' = 1000); si no, es decimal ('12,5').
    """
    t = (texto or "").strip().replace("$", "").replace(" ", "")
    if "," in t and "." in t:
        decimal_sep = "," if t.rfind(",") > t.rfind(".") else "."
        miles_sep = "." if decimal_sep == "," else ","
        t = t.replace(miles_sep, "").replace(decimal_sep, ".")
    elif "," in t or "." in t:
        sep = "," if "," in t else "."
        entero, _, dec = t.rpartition(sep)
        if t.count(sep) > 1 or len(dec) == 3:
            t = t.replace(sep, "")
        else:
            t = f"{entero}.{dec}"
    try:
        valor = Decimal(t)
    except InvalidOperation:
        raise ValueError("El monto no es un número válido.")
    if not valor.is_finite() or valor <= 0:
        raise ValueError("El monto debe ser mayor que cero.")
    centavos = (valor * 100).quantize(Decimal("1"))
    if centavos != valor * 100:
        raise ValueError("El monto admite como máximo 2 decimales.")
    return int(centavos)


@app.template_filter("dinero")
def formato_dinero(centavos):
    signo = "-" if centavos < 0 else ""
    entero, resto = divmod(abs(int(centavos)), 100)
    return f"{signo}${entero:,}.{resto:02d}".replace(",", "X").replace(".", ",").replace("X", ".")


@app.template_filter("pct")
def formato_pct(valor):
    return f"{valor:.1f} %".replace(".", ",")


def mes_valido(mes):
    return bool(mes) and len(mes) == 7 and mes[4] == "-" and mes.replace("-", "").isdigit()


def calcular(db, mes=None):
    """Devuelve los movimientos y todos los cálculos (opcionalmente de un mes AAAA-MM)."""
    filtro, params = ("WHERE substr(fecha, 1, 7) = ?", (mes,)) if mes else ("", ())

    movimientos = db.execute(
        f"SELECT * FROM movimientos {filtro} ORDER BY fecha DESC, id DESC", params
    ).fetchall()

    t = db.execute(
        f"""SELECT
              COALESCE(SUM(CASE WHEN tipo='ingreso' THEN centavos END), 0) AS ingresos,
              COALESCE(SUM(CASE WHEN tipo='gasto'   THEN centavos END), 0) AS gastos,
              COUNT(*) AS cantidad,
              COALESCE(MAX(CASE WHEN tipo='gasto' THEN centavos END), 0) AS mayor_gasto
            FROM movimientos {filtro}""",
        params,
    ).fetchone()
    cant_gastos = sum(1 for m in movimientos if m["tipo"] == "gasto")
    totales = {
        "ingresos": t["ingresos"],
        "gastos": t["gastos"],
        "balance": t["ingresos"] - t["gastos"],
        "cantidad": t["cantidad"],
        "mayor_gasto": t["mayor_gasto"],
        "gasto_promedio": round(t["gastos"] / cant_gastos) if cant_gastos else 0,
        "ahorro_pct": round(100 * (t["ingresos"] - t["gastos"]) / t["ingresos"], 1) if t["ingresos"] else None,
    }

    por_categoria = db.execute(
        f"""SELECT categoria,
              COALESCE(SUM(CASE WHEN tipo='ingreso' THEN centavos END), 0) AS ingresos,
              COALESCE(SUM(CASE WHEN tipo='gasto'   THEN centavos END), 0) AS gastos,
              COUNT(*) AS cantidad
            FROM movimientos {filtro}
            GROUP BY categoria ORDER BY gastos DESC, ingresos DESC, categoria""",
        params,
    ).fetchall()

    por_mes = db.execute(
        """SELECT substr(fecha, 1, 7) AS mes,
              COALESCE(SUM(CASE WHEN tipo='ingreso' THEN centavos END), 0) AS ingresos,
              COALESCE(SUM(CASE WHEN tipo='gasto'   THEN centavos END), 0) AS gastos
            FROM movimientos GROUP BY mes ORDER BY mes DESC"""
    ).fetchall()

    categorias = [r[0] for r in db.execute("SELECT DISTINCT categoria FROM movimientos ORDER BY categoria")]
    return movimientos, totales, por_categoria, por_mes, categorias


# --- Rutas -------------------------------------------------------------------

@app.get("/")
def index():
    mes = request.args.get("mes", "")
    if not mes_valido(mes):
        mes = ""
    movimientos, totales, por_categoria, por_mes, categorias = calcular(get_db(), mes or None)
    return render_template(
        "index.html",
        movimientos=movimientos,
        totales=totales,
        por_categoria=por_categoria,
        por_mes=por_mes,
        categorias=categorias,
        mes=mes,
        hoy=date.today().isoformat(),
    )


@app.post("/agregar")
def agregar():
    f = request.form
    tipo = f.get("tipo", "gasto")
    fecha = f.get("fecha") or date.today().isoformat()
    try:
        if tipo not in TIPOS:
            raise ValueError("Tipo de movimiento inválido.")
        try:
            date.fromisoformat(fecha)
        except ValueError:
            raise ValueError("Fecha inválida.")
        centavos = parsear_monto(f.get("monto"))
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("index", mes=request.args.get("mes", "")))

    db = get_db()
    db.execute(
        "INSERT INTO movimientos (fecha, descripcion, categoria, tipo, centavos) VALUES (?, ?, ?, ?, ?)",
        (
            fecha,
            f.get("descripcion", "").strip()[:200],
            (f.get("categoria", "").strip() or "General")[:60],
            tipo,
            centavos,
        ),
    )
    db.commit()
    flash(f"{tipo.capitalize()} de {formato_dinero(centavos)} guardado.", "ok")
    return redirect(url_for("index", mes=request.args.get("mes", "")))


@app.post("/eliminar/<int:mov_id>")
def eliminar(mov_id):
    db = get_db()
    if db.execute("DELETE FROM movimientos WHERE id = ?", (mov_id,)).rowcount == 0:
        abort(404)
    db.commit()
    flash("Movimiento eliminado.", "ok")
    return redirect(url_for("index", mes=request.args.get("mes", "")))


@app.get("/exportar.xlsx")
def exportar():
    """Genera un Excel con los movimientos y una hoja de resumen con fórmulas."""
    mes = request.args.get("mes", "")
    if not mes_valido(mes):
        mes = ""
    movimientos, _, por_categoria, por_mes, _ = calcular(get_db(), mes or None)

    wb = Workbook()
    negrita = Font(bold=True, color="FFFFFF")
    relleno = PatternFill("solid", fgColor="2F5597")

    def encabezado(ws, columnas, anchos):
        ws.append(columnas)
        for i, ancho in enumerate(anchos, start=1):
            celda = ws.cell(row=1, column=i)
            celda.font, celda.fill = negrita, relleno
            ws.column_dimensions[get_column_letter(i)].width = ancho
        ws.freeze_panes = "A2"

    # Hoja 1: movimientos (gastos con signo negativo en la columna "Neto")
    ws = wb.active
    ws.title = "Movimientos"
    encabezado(ws, ["Fecha", "Descripción", "Categoría", "Tipo", "Monto", "Neto"], [12, 34, 18, 10, 14, 14])
    for m in reversed(movimientos):  # orden cronológico
        fila = ws.max_row + 1
        ws.append([date.fromisoformat(m["fecha"]), m["descripcion"], m["categoria"], m["tipo"], m["centavos"] / 100])
        ws.cell(row=fila, column=6, value=f'=IF(D{fila}="gasto",-E{fila},E{fila})')
        ws.cell(row=fila, column=1).number_format = "yyyy-mm-dd"
    ultima = max(ws.max_row, 2)
    for fila in ws.iter_rows(min_row=2, min_col=5, max_col=6):
        for c in fila:
            c.number_format = FORMATO_MONEDA
    ws.auto_filter.ref = f"A1:F{ws.max_row}"

    rango_tipo, rango_monto, rango_cat = f"Movimientos!D2:D{ultima}", f"Movimientos!E2:E{ultima}", f"Movimientos!C2:C{ultima}"

    # Hoja 2: resumen con fórmulas (se recalcula si se editan los movimientos en Excel)
    rs = wb.create_sheet("Resumen")
    encabezado(rs, ["Concepto", "Valor"], [26, 16])
    filas = [
        ("Total ingresos", f'=SUMIF({rango_tipo},"ingreso",{rango_monto})'),
        ("Total gastos", f'=SUMIF({rango_tipo},"gasto",{rango_monto})'),
        ("Balance", "=B2-B3"),
        ("Cantidad de movimientos", f"=COUNTA({rango_tipo})"),
        ("Gasto promedio", f'=IFERROR(AVERAGEIF({rango_tipo},"gasto",{rango_monto}),0)'),
        ("Mayor gasto", f'=SUMPRODUCT(MAX(({rango_tipo}="gasto")*{rango_monto}))'),
        ("% de ahorro", "=IFERROR(B4/B2,0)"),
    ]
    for concepto, formula in filas:
        rs.append([concepto, formula])
    for r in range(2, 8):
        rs.cell(row=r, column=2).number_format = FORMATO_MONEDA
    rs["B5"].number_format = "0"
    rs["B8"].number_format = "0.0%"
    if mes:
        rs["D1"] = f"Período: {mes}"

    # Hoja 3: por categoría
    cs = wb.create_sheet("Por categoría")
    encabezado(cs, ["Categoría", "Ingresos", "Gastos", "Neto"], [20, 14, 14, 14])
    for i, c in enumerate(por_categoria, start=2):
        cs.append([
            c["categoria"],
            f'=SUMIFS({rango_monto},{rango_cat},A{i},{rango_tipo},"ingreso")',
            f'=SUMIFS({rango_monto},{rango_cat},A{i},{rango_tipo},"gasto")',
            f"=B{i}-C{i}",
        ])
        for col in range(2, 5):
            cs.cell(row=i, column=col).number_format = FORMATO_MONEDA

    # Hoja 4: por mes (siempre todo el historial)
    ms = wb.create_sheet("Por mes")
    encabezado(ms, ["Mes", "Ingresos", "Gastos", "Balance"], [12, 14, 14, 14])
    for i, r in enumerate(por_mes, start=2):
        ms.append([r["mes"], r["ingresos"] / 100, r["gastos"] / 100, f"=B{i}-C{i}"])
        for col in range(2, 5):
            ms.cell(row=i, column=col).number_format = FORMATO_MONEDA

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    nombre = f"montos_{mes or date.today().isoformat()}.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=nombre,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=os.environ.get("DEBUG") == "1")
