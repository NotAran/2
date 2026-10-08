"""Planilla de montos con usuarios.

El admin define las filas (cuentas) y las columnas (conceptos) de la planilla,
crea usuarios y les asigna qué filas y columnas pueden ver y cargar. Los montos
se guardan en una base de datos SQLite interna, la página calcula los totales y
todo se puede descargar en Excel.
"""

import io
import os
import secrets
import sqlite3
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from flask import Flask, abort, flash, g, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VERSION_ESQUEMA = 2
FILAS_INICIALES = ("Mercado", "Naranja", "Ualá", "Personal", "Brubank", "Lemon", "Prex", "Arq", "Binance")
COLUMNAS_INICIALES = (("Ingresos", 1), ("Gastos", -1))  # signo: suma o resta en el total
FORMATO_MONEDA = '"$"#,##0.00;[Red]-"$"#,##0.00'


def clave_secreta():
    """Usa SECRET_KEY si está definida; si no, genera una y la guarda junto a la app."""
    if os.environ.get("SECRET_KEY"):
        return os.environ["SECRET_KEY"]
    ruta = os.path.join(BASE_DIR, ".secret_key")
    if not os.path.exists(ruta):
        with open(ruta, "w") as f:
            f.write(secrets.token_hex(32))
    with open(ruta) as f:
        return f.read().strip()


app = Flask(__name__)
app.config.from_mapping(
    SECRET_KEY=clave_secreta(),
    DATABASE=os.environ.get("MONTOS_DB", os.path.join(BASE_DIR, "montos.db")),
)


# --- Base de datos -----------------------------------------------------------

ESQUEMA = """
CREATE TABLE IF NOT EXISTS usuarios (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre     TEXT    NOT NULL UNIQUE COLLATE NOCASE,
    clave_hash TEXT    NOT NULL,
    es_admin   INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS filas (
    id     INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT    NOT NULL UNIQUE COLLATE NOCASE,
    orden  INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS columnas (
    id     INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT    NOT NULL UNIQUE COLLATE NOCASE,
    signo  INTEGER NOT NULL DEFAULT 1 CHECK (signo IN (1, -1)),
    orden  INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS celdas (
    fila_id         INTEGER NOT NULL REFERENCES filas(id)    ON DELETE CASCADE,
    columna_id      INTEGER NOT NULL REFERENCES columnas(id) ON DELETE CASCADE,
    centavos        INTEGER NOT NULL,  -- monto en centavos, sin decimales flotantes
    actualizado_por TEXT,
    actualizado     TEXT,
    PRIMARY KEY (fila_id, columna_id)
);
CREATE TABLE IF NOT EXISTS permisos_filas (
    usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
    fila_id    INTEGER NOT NULL REFERENCES filas(id)    ON DELETE CASCADE,
    PRIMARY KEY (usuario_id, fila_id)
);
CREATE TABLE IF NOT EXISTS permisos_columnas (
    usuario_id INTEGER NOT NULL REFERENCES usuarios(id) ON DELETE CASCADE,
    columna_id INTEGER NOT NULL REFERENCES columnas(id) ON DELETE CASCADE,
    PRIMARY KEY (usuario_id, columna_id)
);
"""


def inicializar(db):
    """Crea las tablas y carga las filas y columnas iniciales (solo la primera vez)."""
    db.executescript(ESQUEMA)
    db.executemany(
        "INSERT OR IGNORE INTO filas (nombre, orden) VALUES (?, ?)",
        [(nombre, i) for i, nombre in enumerate(FILAS_INICIALES)],
    )
    db.executemany(
        "INSERT OR IGNORE INTO columnas (nombre, signo, orden) VALUES (?, ?, ?)",
        [(nombre, signo, i) for i, (nombre, signo) in enumerate(COLUMNAS_INICIALES)],
    )
    db.execute(f"PRAGMA user_version = {VERSION_ESQUEMA}")
    db.commit()


def get_db():
    if "db" not in g:
        db = sqlite3.connect(app.config["DATABASE"])
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys = ON")
        if db.execute("PRAGMA user_version").fetchone()[0] < VERSION_ESQUEMA:
            inicializar(db)
        g.db = db
    return g.db


@app.teardown_appcontext
def cerrar_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


# --- Utilidades --------------------------------------------------------------

def parsear_monto(texto):
    """Convierte texto como '1.234,56', '1234.56', '-500' o '$ 500' a centavos (int).

    Si aparecen punto y coma, el último que aparece es el separador decimal.
    Si aparece uno solo, es de miles cuando se repite o lo siguen exactamente
    3 dígitos ('9.850' = 9850, '1,000' = 1000); si no, es decimal ('12,5').
    """
    t = (texto or "").strip().replace("$", "").replace(" ", "")
    negativo = t.startswith("-")
    t = t.removeprefix("-")
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
        raise ValueError("no es un número válido")
    if not valor.is_finite() or valor < 0:
        raise ValueError("no es un número válido")
    centavos = (valor * 100).quantize(Decimal("1"))
    if centavos != valor * 100:
        raise ValueError("admite como máximo 2 decimales")
    return -int(centavos) if negativo else int(centavos)


@app.template_filter("numero")
def formato_numero(centavos):
    """123456 -> '1.234,56' (formato argentino, sin símbolo)."""
    signo = "-" if centavos < 0 else ""
    entero, resto = divmod(abs(int(centavos)), 100)
    return f"{signo}{entero:,}".replace(",", ".") + f",{resto:02d}"


@app.template_filter("dinero")
def formato_dinero(centavos):
    numero = formato_numero(centavos)
    return f"-${numero[1:]}" if numero.startswith("-") else f"${numero}"


def validar_usuario(nombre, clave):
    if not 3 <= len(nombre) <= 30 or not nombre.replace("_", "").replace(".", "").isalnum():
        return "El usuario debe tener entre 3 y 30 letras o números (sin espacios)."
    return validar_clave(clave)


def validar_clave(clave):
    if len(clave) < 6:
        return "La contraseña debe tener al menos 6 caracteres."
    return None


# --- Sesión, permisos y protección CSRF ----------------------------------------

def usuario_actual():
    if "usuario" not in g:
        uid = session.get("usuario_id")
        g.usuario = get_db().execute("SELECT * FROM usuarios WHERE id = ?", (uid,)).fetchone() if uid else None
    return g.usuario


def hay_usuarios():
    return get_db().execute("SELECT 1 FROM usuarios LIMIT 1").fetchone() is not None


@app.before_request
def proteger():
    endpoint = request.endpoint or ""
    if endpoint == "static":
        return None
    token = session.setdefault("csrf", secrets.token_hex(16))
    if request.method == "POST":
        if not secrets.compare_digest(token, request.form.get("csrf", "")):
            abort(400, "El formulario venció. Recargá la página e intentá de nuevo.")
    if not hay_usuarios():
        return None if endpoint == "setup" else redirect(url_for("setup"))
    if endpoint in ("login", "setup"):
        return None
    if usuario_actual() is None:
        return redirect(url_for("login"))
    if endpoint.startswith("admin") and not usuario_actual()["es_admin"]:
        abort(403)
    return None


@app.context_processor
def contexto():
    return {"csrf": session.get("csrf", ""), "usuario": usuario_actual()}


# --- Planilla ------------------------------------------------------------------

def visibles(db, usuario):
    """Filas y columnas que ve el usuario: todas si es admin, si no las asignadas."""
    if usuario["es_admin"]:
        filas = db.execute("SELECT * FROM filas ORDER BY orden, id").fetchall()
        columnas = db.execute("SELECT * FROM columnas ORDER BY orden, id").fetchall()
    else:
        filas = db.execute(
            """SELECT f.* FROM filas f JOIN permisos_filas p ON p.fila_id = f.id
               WHERE p.usuario_id = ? ORDER BY f.orden, f.id""",
            (usuario["id"],),
        ).fetchall()
        columnas = db.execute(
            """SELECT c.* FROM columnas c JOIN permisos_columnas p ON p.columna_id = c.id
               WHERE p.usuario_id = ? ORDER BY c.orden, c.id""",
            (usuario["id"],),
        ).fetchall()
    return filas, columnas


def armar_planilla(db, usuario):
    """Datos de la planilla visible y sus totales.

    Total de fila = suma de las columnas con signo + menos las columnas con signo −.
    Total de columna = suma de la columna.
    """
    filas, columnas = visibles(db, usuario)
    celdas = {(c["fila_id"], c["columna_id"]): c for c in db.execute("SELECT * FROM celdas")}

    def valor(f, c):
        celda = celdas.get((f["id"], c["id"]))
        return celda["centavos"] if celda else 0

    total_fila = {f["id"]: sum(valor(f, c) * c["signo"] for c in columnas) for f in filas}
    total_columna = {c["id"]: sum(valor(f, c) for f in filas) for c in columnas}
    return {
        "filas": filas,
        "columnas": columnas,
        "celdas": celdas,
        "total_fila": total_fila,
        "total_columna": total_columna,
        "total": sum(total_fila.values()),
    }


@app.get("/")
def planilla():
    return render_template("planilla.html", **armar_planilla(get_db(), usuario_actual()))


@app.post("/guardar")
def guardar():
    db = get_db()
    usuario = usuario_actual()
    filas, columnas = visibles(db, usuario)  # solo se guardan celdas que el usuario puede ver
    actuales = {(c["fila_id"], c["columna_id"]): c["centavos"] for c in db.execute("SELECT * FROM celdas")}
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    errores, cambios = [], 0

    for f in filas:
        for c in columnas:
            campo = f"c_{f['id']}_{c['id']}"
            if campo not in request.form:
                continue
            texto = request.form[campo].strip()
            clave = (f["id"], c["id"])
            if not texto:
                if clave in actuales:
                    db.execute("DELETE FROM celdas WHERE fila_id = ? AND columna_id = ?", clave)
                    cambios += 1
                continue
            try:
                centavos = parsear_monto(texto)
            except ValueError as e:
                errores.append(f"{f['nombre']} / {c['nombre']}: «{texto}» {e}.")
                continue
            if actuales.get(clave) == centavos:
                continue
            db.execute(
                """INSERT INTO celdas (fila_id, columna_id, centavos, actualizado_por, actualizado)
                   VALUES (?, ?, ?, ?, ?)
                   ON CONFLICT (fila_id, columna_id) DO UPDATE SET
                     centavos = excluded.centavos,
                     actualizado_por = excluded.actualizado_por,
                     actualizado = excluded.actualizado""",
                (*clave, centavos, usuario["nombre"], ahora),
            )
            cambios += 1

    db.commit()
    for error in errores:
        flash(error, "error")
    if cambios:
        flash(f"Se guardaron {cambios} cambio(s).", "ok")
    elif not errores:
        flash("No había cambios para guardar.", "ok")
    return redirect(url_for("planilla"))


@app.get("/exportar.xlsx")
def exportar():
    """Excel con la planilla visible; los totales son fórmulas de Excel."""
    datos = armar_planilla(get_db(), usuario_actual())
    filas, columnas, celdas = datos["filas"], datos["columnas"], datos["celdas"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Planilla"
    encabezados = ["Cuenta"] + [f"{c['nombre']} ({'+' if c['signo'] > 0 else '−'})" for c in columnas] + ["Total"]
    ws.append(encabezados)
    for i in range(1, len(encabezados) + 1):
        celda = ws.cell(row=1, column=i)
        celda.font, celda.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="2F5597")
        ws.column_dimensions[get_column_letter(i)].width = 16 if i > 1 else 18
    ws.freeze_panes = "B2"

    col_total = len(columnas) + 2
    letra = {c["id"]: get_column_letter(j) for j, c in enumerate(columnas, start=2)}
    for i, f in enumerate(filas, start=2):
        ws.cell(row=i, column=1, value=f["nombre"])
        for c in columnas:
            celda = celdas.get((f["id"], c["id"]))
            ws[f"{letra[c['id']]}{i}"] = celda["centavos"] / 100 if celda else None
        terminos = "".join(f"{'+' if c['signo'] > 0 else '-'}{letra[c['id']]}{i}" for c in columnas)
        ws.cell(row=i, column=col_total, value=f"={terminos or 0}")

    fila_total = len(filas) + 2
    ws.cell(row=fila_total, column=1, value="Total").font = Font(bold=True)
    for j in range(2, col_total + 1):
        l = get_column_letter(j)
        ws.cell(row=fila_total, column=j, value=f"=SUM({l}2:{l}{fila_total - 1})" if filas else 0).font = Font(bold=True)
    for fila in ws.iter_rows(min_row=2, min_col=2, max_col=col_total):
        for c in fila:
            c.number_format = FORMATO_MONEDA

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"planilla_{date.today().isoformat()}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# --- Login ---------------------------------------------------------------------

@app.route("/setup", methods=["GET", "POST"])
def setup():
    """Primera vez: crear la cuenta del admin."""
    if hay_usuarios():
        return redirect(url_for("login"))
    if request.method == "POST":
        nombre, clave = request.form.get("nombre", "").strip(), request.form.get("clave", "")
        error = validar_usuario(nombre, clave)
        if not error:
            db = get_db()
            cur = db.execute(
                "INSERT INTO usuarios (nombre, clave_hash, es_admin) VALUES (?, ?, 1)",
                (nombre, generate_password_hash(clave)),
            )
            db.commit()
            session.clear()
            session["usuario_id"] = cur.lastrowid
            flash(f"Listo, {nombre}: sos el admin.", "ok")
            return redirect(url_for("planilla"))
        flash(error, "error")
    return render_template("setup.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        nombre, clave = request.form.get("nombre", "").strip(), request.form.get("clave", "")
        u = get_db().execute("SELECT * FROM usuarios WHERE nombre = ?", (nombre,)).fetchone()
        if u and check_password_hash(u["clave_hash"], clave):
            session.clear()
            session["usuario_id"] = u["id"]
            return redirect(url_for("planilla"))
        flash("Usuario o contraseña incorrectos.", "error")
    return render_template("login.html")


@app.post("/salir")
def salir():
    session.clear()
    return redirect(url_for("login"))


# --- Panel de admin -------------------------------------------------------------

@app.get("/admin")
def admin():
    db = get_db()
    permisos = {"filas": {}, "columnas": {}}
    for tabla, campo in (("filas", "fila_id"), ("columnas", "columna_id")):
        for p in db.execute(f"SELECT * FROM permisos_{tabla}"):
            permisos[tabla].setdefault(p["usuario_id"], set()).add(p[campo])
    return render_template(
        "admin.html",
        usuarios=db.execute("SELECT * FROM usuarios ORDER BY es_admin DESC, nombre").fetchall(),
        filas=db.execute("SELECT * FROM filas ORDER BY orden, id").fetchall(),
        columnas=db.execute("SELECT * FROM columnas ORDER BY orden, id").fetchall(),
        permisos=permisos,
    )


def volver_admin(seccion):
    return redirect(url_for("admin") + f"#{seccion}")


@app.post("/admin/usuarios")
def admin_crear_usuario():
    nombre, clave = request.form.get("nombre", "").strip(), request.form.get("clave", "")
    error = validar_usuario(nombre, clave)
    if not error:
        db = get_db()
        try:
            db.execute(
                "INSERT INTO usuarios (nombre, clave_hash) VALUES (?, ?)", (nombre, generate_password_hash(clave))
            )
            db.commit()
            flash(f"Usuario «{nombre}» creado. Ahora asignale filas y columnas.", "ok")
        except sqlite3.IntegrityError:
            error = "Ya existe un usuario con ese nombre."
    if error:
        flash(error, "error")
    return volver_admin("usuarios")


@app.post("/admin/usuarios/<int:uid>/eliminar")
def admin_eliminar_usuario(uid):
    db = get_db()
    if db.execute("DELETE FROM usuarios WHERE id = ? AND es_admin = 0", (uid,)).rowcount == 0:
        abort(404)
    db.commit()
    flash("Usuario eliminado.", "ok")
    return volver_admin("usuarios")


@app.post("/admin/usuarios/<int:uid>/clave")
def admin_cambiar_clave(uid):
    clave = request.form.get("clave", "")
    error = validar_clave(clave)
    if error:
        flash(error, "error")
        return volver_admin("usuarios")
    db = get_db()
    if db.execute("UPDATE usuarios SET clave_hash = ? WHERE id = ?", (generate_password_hash(clave), uid)).rowcount == 0:
        abort(404)
    db.commit()
    flash("Contraseña actualizada.", "ok")
    return volver_admin("usuarios")


@app.post("/admin/usuarios/<int:uid>/permisos")
def admin_permisos(uid):
    db = get_db()
    if not db.execute("SELECT 1 FROM usuarios WHERE id = ? AND es_admin = 0", (uid,)).fetchone():
        abort(404)
    for tabla, campo in (("filas", "fila_id"), ("columnas", "columna_id")):
        db.execute(f"DELETE FROM permisos_{tabla} WHERE usuario_id = ?", (uid,))
        db.executemany(
            f"INSERT INTO permisos_{tabla} (usuario_id, {campo}) SELECT ?, id FROM {tabla} WHERE id = ?",
            [(uid, item_id) for item_id in request.form.getlist(tabla, type=int)],
        )
    db.commit()
    flash("Permisos guardados.", "ok")
    return volver_admin(f"permisos-{uid}")


@app.post("/admin/<any(filas, columnas):tabla>")
def admin_agregar(tabla):
    nombre = request.form.get("nombre", "").strip()[:40]
    if not nombre:
        flash("Escribí un nombre.", "error")
        return volver_admin(tabla)
    db = get_db()
    orden = db.execute(f"SELECT COALESCE(MAX(orden), -1) + 1 FROM {tabla}").fetchone()[0]
    try:
        if tabla == "filas":
            db.execute("INSERT INTO filas (nombre, orden) VALUES (?, ?)", (nombre, orden))
        else:
            signo = -1 if request.form.get("signo") == "-1" else 1
            db.execute("INSERT INTO columnas (nombre, signo, orden) VALUES (?, ?, ?)", (nombre, signo, orden))
        db.commit()
        flash(f"«{nombre}» agregada. Acordate de asignarla a los usuarios que la tengan que ver.", "ok")
    except sqlite3.IntegrityError:
        flash(f"Ya existe una {tabla[:-1]} llamada «{nombre}».", "error")
    return volver_admin(tabla)


@app.post("/admin/<any(filas, columnas):tabla>/<int:item_id>/eliminar")
def admin_eliminar(tabla, item_id):
    db = get_db()
    if db.execute(f"DELETE FROM {tabla} WHERE id = ?", (item_id,)).rowcount == 0:
        abort(404)
    db.commit()
    flash("Eliminada, junto con sus montos.", "ok")
    return volver_admin(tabla)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=os.environ.get("DEBUG") == "1")
