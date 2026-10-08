"""Cajas por turno.

Cada día tiene 3 cajas (turnos noche, mañana y tarde). Cada caja es una planilla
de cuentas (filas) por personas (columnas) y se divide en:

- Turno anterior: un solo total, con el que cerró la caja anterior (se calcula solo).
- Turno actual: los montos que se cargan en las columnas durante el turno.

Total caja = turno anterior + turno actual, y pasa a ser el turno anterior de
la caja siguiente. El admin define filas y columnas, crea usuarios y les asigna
qué filas y columnas ven y cargan. Los datos se guardan en una base SQLite
interna y se pueden descargar en Excel.
"""

import io
import os
import secrets
import sqlite3
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation

from flask import Flask, abort, flash, g, redirect, render_template, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
VERSION_ESQUEMA = 5
FILAS_INICIALES = ("Mercado", "Naranja", "Ualá", "Personal", "Brubank", "Lemon", "Prex", "Arq", "Binance")
COLUMNAS_INICIALES = (("Paco", 1), ("Antonio", 1), ("Ibra", 1))  # signo: suma o resta en el total
TURNOS = (("noche", "Noche"), ("manana", "Mañana"), ("tarde", "Tarde"))  # orden de las cajas en el día
TURNO_POR_SLUG = {slug: i for i, (slug, _) in enumerate(TURNOS)}
# Datos de cada caja que no suman en la caja, en el orden en que se muestran.
# El saldo no se carga: se calcula como depósito − retiro.
EXTRAS = (("deposito", "Depósito"), ("retiro", "Retiro"), ("bono", "Bono"), ("saldo", "Saldo"), ("bajada", "Bajada"))
EXTRAS_CALCULADOS = {"saldo"}
DIAS = ("lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo")
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
    orden  INTEGER NOT NULL DEFAULT 0,
    activo INTEGER NOT NULL DEFAULT 1  -- 0 = quitada pero con historial
);
CREATE TABLE IF NOT EXISTS columnas (
    id     INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT    NOT NULL UNIQUE COLLATE NOCASE,
    signo  INTEGER NOT NULL DEFAULT 1 CHECK (signo IN (1, -1)),
    orden  INTEGER NOT NULL DEFAULT 0,
    activo INTEGER NOT NULL DEFAULT 1  -- 0 = quitada pero con historial
);
CREATE TABLE IF NOT EXISTS celdas (
    fecha           TEXT    NOT NULL,  -- AAAA-MM-DD
    turno           INTEGER NOT NULL CHECK (turno IN (0, 1, 2)),  -- índice en TURNOS
    fila_id         INTEGER NOT NULL REFERENCES filas(id)    ON DELETE CASCADE,
    columna_id      INTEGER NOT NULL REFERENCES columnas(id) ON DELETE CASCADE,
    centavos        INTEGER NOT NULL,  -- monto en centavos, sin decimales flotantes
    actualizado_por TEXT,
    actualizado     TEXT,
    PRIMARY KEY (fecha, turno, fila_id, columna_id)
);
CREATE TABLE IF NOT EXISTS extras (
    fecha           TEXT    NOT NULL,
    turno           INTEGER NOT NULL CHECK (turno IN (0, 1, 2)),
    campo           TEXT    NOT NULL,  -- clave en EXTRAS
    centavos        INTEGER NOT NULL,
    actualizado_por TEXT,
    actualizado     TEXT,
    PRIMARY KEY (fecha, turno, campo)
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
    """Crea o actualiza las tablas y carga las filas y columnas iniciales si no hay."""
    if db.execute("PRAGMA user_version").fetchone()[0] == 2:
        db.execute("DROP TABLE IF EXISTS celdas")  # versión anterior: montos sin caja ni turno
    db.executescript(ESQUEMA)
    db.execute("DELETE FROM extras WHERE campo = 'saldo'")  # desde la versión 5 el saldo se calcula
    for tabla in ("filas", "columnas"):
        if "activo" not in {r["name"] for r in db.execute(f"PRAGMA table_info({tabla})")}:
            db.execute(f"ALTER TABLE {tabla} ADD COLUMN activo INTEGER NOT NULL DEFAULT 1")
    if not db.execute("SELECT 1 FROM filas").fetchone():
        db.executemany(
            "INSERT INTO filas (nombre, orden) VALUES (?, ?)", [(n, i) for i, n in enumerate(FILAS_INICIALES)]
        )
    if not db.execute("SELECT 1 FROM columnas").fetchone():
        db.executemany(
            "INSERT INTO columnas (nombre, signo, orden) VALUES (?, ?, ?)",
            [(n, signo, i) for i, (n, signo) in enumerate(COLUMNAS_INICIALES)],
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


# --- Cajas -------------------------------------------------------------------
# Una caja se identifica por (fecha, turno), con turno = índice en TURNOS.

def mover(caja, pasos):
    """Caja que está `pasos` turnos después (o antes, si es negativo)."""
    fecha, turno = caja
    n = turno + pasos
    return fecha + timedelta(days=n // len(TURNOS)), n % len(TURNOS)


@app.template_global()
def nombre_caja(caja):
    fecha, turno = caja
    return f"{DIAS[fecha.weekday()].capitalize()} {fecha:%d/%m/%Y} – {TURNOS[turno][1]}"


@app.template_global()
def url_caja(caja, endpoint="planilla"):
    fecha, turno = caja
    return url_for(endpoint, fecha=fecha.isoformat(), turno=TURNOS[turno][0])


@app.template_global()
def turno_slug(caja):
    return TURNOS[caja[1]][0]


def ultima_caja():
    r = get_db().execute("SELECT fecha, turno FROM celdas ORDER BY fecha DESC, turno DESC LIMIT 1").fetchone()
    return (date.fromisoformat(r["fecha"]), r["turno"]) if r else (date.today(), 0)


def caja_pedida(estricto=False):
    """Caja de los parámetros fecha y turno; si faltan, la última con datos."""
    try:
        return date.fromisoformat(request.values.get("fecha", "")), TURNO_POR_SLUG[request.values.get("turno", "")]
    except (ValueError, KeyError):
        if estricto:
            abort(400, "Caja inválida.")
        return ultima_caja()


def permitidas(db, usuario):
    """Todas las filas y columnas (también las quitadas) que le corresponden al usuario."""
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


def armar_caja(db, usuario, caja):
    """Datos y cálculos de una caja.

    Los totales son siempre de toda la caja (todas las filas y columnas):
    - Turno anterior: total con el que cerró la caja anterior, es decir la suma
      (con signo) de todos los montos de todas las cajas previas.
    - Turno actual: suma (con signo) de los montos de esta caja.
    - Total caja = turno anterior + turno actual.
    La grilla muestra las filas y columnas asignadas al usuario: las activas, y las
    quitadas solo si tienen montos en esta caja. Los extras (depósito, retiro, bono,
    bajada) no entran en las sumas de la caja; saldo = depósito − retiro y
    total + bajada = total caja + bajada (no pasa a la caja siguiente).
    """
    fecha, turno = caja[0].isoformat(), caja[1]
    filas_ok, columnas_ok = permitidas(db, usuario)
    celdas = {
        (r["fila_id"], r["columna_id"]): r
        for r in db.execute("SELECT * FROM celdas WHERE fecha = ? AND turno = ?", (fecha, turno))
    }
    filas = [f for f in filas_ok if f["activo"] or any(k[0] == f["id"] for k in celdas)]
    columnas = [c for c in columnas_ok if c["activo"] or any(k[1] == c["id"] for k in celdas)]

    anterior, actual = db.execute(
        """SELECT
             COALESCE(SUM(CASE WHEN ce.fecha < ? OR (ce.fecha = ? AND ce.turno < ?) THEN ce.centavos * co.signo END), 0),
             COALESCE(SUM(CASE WHEN ce.fecha = ? AND ce.turno = ? THEN ce.centavos * co.signo END), 0)
           FROM celdas ce JOIN columnas co ON co.id = ce.columna_id""",
        (fecha, fecha, turno, fecha, turno),
    ).fetchone()

    def valor(f, c):
        celda = celdas.get((f["id"], c["id"]))
        return celda["centavos"] if celda else 0

    extras = {r["campo"]: r for r in db.execute("SELECT * FROM extras WHERE fecha = ? AND turno = ?", (fecha, turno))}

    def monto_extra(campo):
        return extras[campo]["centavos"] if campo in extras else 0

    total_fila = {f["id"]: sum(valor(f, c) * c["signo"] for c in columnas) for f in filas}
    return {
        "caja": caja,
        "filas": filas,
        "columnas": columnas,
        "celdas": celdas,
        "total_fila": total_fila,
        "total_columna": {c["id"]: sum(valor(f, c) for f in filas) for c in columnas},
        "total_grilla": sum(total_fila.values()),
        "totales": {
            "anterior": anterior,
            "actual": actual,
            "caja": anterior + actual,
            "con_bajada": anterior + actual + monto_extra("bajada"),
        },
        "extras": extras,
        "saldo": monto_extra("deposito") - monto_extra("retiro"),
    }


def historial(db):
    """Una entrada por caja con montos: turno actual y total al cierre (más nuevas primero)."""
    saldo, filas = 0, []
    for r in db.execute(
        """SELECT ce.fecha, ce.turno, SUM(ce.centavos * co.signo) AS movimiento
           FROM celdas ce JOIN columnas co ON co.id = ce.columna_id
           GROUP BY ce.fecha, ce.turno ORDER BY ce.fecha, ce.turno"""
    ):
        saldo += r["movimiento"]
        filas.append({"caja": (date.fromisoformat(r["fecha"]), r["turno"]), "actual": r["movimiento"], "cierre": saldo})
    return filas[::-1]


@app.get("/")
def planilla():
    db = get_db()
    caja = caja_pedida()
    return render_template(
        "planilla.html",
        **armar_caja(db, usuario_actual(), caja),
        caja_previa=mover(caja, -1),
        caja_siguiente=mover(caja, 1),
        turnos=[((caja[0], i), nombre) for i, (_, nombre) in enumerate(TURNOS)],
        historial=historial(db)[:15],
        campos_extras=EXTRAS,
    )


def leer_cambio(texto, actual, etiqueta, errores):
    """Interpreta lo escrito en un campo. Devuelve (hay_cambio, centavos); centavos None = borrar."""
    texto = texto.strip()
    if not texto:
        return actual is not None, None
    try:
        centavos = parsear_monto(texto)
    except ValueError as e:
        errores.append(f"{etiqueta}: «{texto}» {e}.")
        return False, None
    return not (actual and actual["centavos"] == centavos), centavos


@app.post("/guardar")
def guardar():
    db = get_db()
    usuario = usuario_actual()
    caja = caja_pedida(estricto=True)
    datos = armar_caja(db, usuario, caja)  # solo se guardan celdas que el usuario puede ver
    fecha, turno = caja[0].isoformat(), caja[1]
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M")
    errores, cambios = [], 0

    for f in datos["filas"]:
        for c in datos["columnas"]:
            campo = f"c_{f['id']}_{c['id']}"
            if campo not in request.form:
                continue
            clave = (fecha, turno, f["id"], c["id"])
            hay_cambio, centavos = leer_cambio(
                request.form[campo], datos["celdas"].get((f["id"], c["id"])), f"{f['nombre']} / {c['nombre']}", errores
            )
            if not hay_cambio:
                continue
            if centavos is None:
                db.execute("DELETE FROM celdas WHERE fecha = ? AND turno = ? AND fila_id = ? AND columna_id = ?", clave)
            else:
                db.execute(
                    """INSERT INTO celdas (fecha, turno, fila_id, columna_id, centavos, actualizado_por, actualizado)
                       VALUES (?, ?, ?, ?, ?, ?, ?)
                       ON CONFLICT (fecha, turno, fila_id, columna_id) DO UPDATE SET
                         centavos = excluded.centavos,
                         actualizado_por = excluded.actualizado_por,
                         actualizado = excluded.actualizado""",
                    (*clave, centavos, usuario["nombre"], ahora),
                )
            cambios += 1

    for campo, nombre in EXTRAS:
        if campo in EXTRAS_CALCULADOS or f"x_{campo}" not in request.form:
            continue
        hay_cambio, centavos = leer_cambio(request.form[f"x_{campo}"], datos["extras"].get(campo), nombre, errores)
        if not hay_cambio:
            continue
        if centavos is None:
            db.execute("DELETE FROM extras WHERE fecha = ? AND turno = ? AND campo = ?", (fecha, turno, campo))
        else:
            db.execute(
                """INSERT INTO extras (fecha, turno, campo, centavos, actualizado_por, actualizado)
                   VALUES (?, ?, ?, ?, ?, ?)
                   ON CONFLICT (fecha, turno, campo) DO UPDATE SET
                     centavos = excluded.centavos,
                     actualizado_por = excluded.actualizado_por,
                     actualizado = excluded.actualizado""",
                (fecha, turno, campo, centavos, usuario["nombre"], ahora),
            )
        cambios += 1

    db.commit()
    for error in errores:
        flash(error, "error")
    if cambios:
        flash(f"Caja {nombre_caja(caja)}: se guardaron {cambios} cambio(s).", "ok")
    elif not errores:
        flash("No había cambios para guardar.", "ok")
    return redirect(url_caja(caja))


def estilo_encabezado(ws, fila, cantidad):
    for i in range(1, cantidad + 1):
        celda = ws.cell(row=fila, column=i)
        celda.font, celda.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="2F5597")


@app.get("/exportar.xlsx")
def exportar():
    """Excel con la caja elegida (totales como fórmulas) y el historial de cajas."""
    db = get_db()
    caja = caja_pedida()
    datos = armar_caja(db, usuario_actual(), caja)
    filas, columnas, celdas, totales = datos["filas"], datos["columnas"], datos["celdas"], datos["totales"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Caja"
    ws["A1"] = f"Caja: {nombre_caja(caja)}"
    ws["A1"].font = Font(bold=True, size=13)
    encabezados = ["Cuenta"] + [f"{c['nombre']}{'' if c['signo'] > 0 else ' (−)'}" for c in columnas] + ["Total"]
    ws.append([])
    ws.append(encabezados)
    estilo_encabezado(ws, 3, len(encabezados))
    for i in range(1, len(encabezados) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16 if i > 1 else 18
    ws.freeze_panes = "B4"

    col_total = len(columnas) + 2
    letra = {c["id"]: get_column_letter(j) for j, c in enumerate(columnas, start=2)}
    for i, f in enumerate(filas, start=4):
        ws.cell(row=i, column=1, value=f["nombre"])
        for c in columnas:
            celda = celdas.get((f["id"], c["id"]))
            ws[f"{letra[c['id']]}{i}"] = celda["centavos"] / 100 if celda else None
        terminos = "".join(f"{'+' if c['signo'] > 0 else '-'}{letra[c['id']]}{i}" for c in columnas)
        ws.cell(row=i, column=col_total, value=f"={terminos or 0}")

    fila_total = len(filas) + 4
    l_total = get_column_letter(col_total)
    ws.cell(row=fila_total, column=1, value="Total").font = Font(bold=True)
    for j in range(2, col_total + 1):
        l = get_column_letter(j)
        ws.cell(row=fila_total, column=j, value=f"=SUM({l}4:{l}{fila_total - 1})" if filas else 0).font = Font(
            bold=True
        )

    # Resumen de la caja. Si el usuario ve toda la caja, el turno actual es la fórmula del total de la grilla.
    r = fila_total + 2
    ws.cell(row=r, column=1, value="Turno anterior")
    ws.cell(row=r, column=2, value=totales["anterior"] / 100)
    ws.cell(row=r + 1, column=1, value="Turno actual")
    completa = datos["total_grilla"] == totales["actual"]
    ws.cell(row=r + 1, column=2, value=f"={l_total}{fila_total}" if completa else totales["actual"] / 100)
    ws.cell(row=r + 2, column=1, value="Total caja").font = Font(bold=True)
    ws.cell(row=r + 2, column=2, value=f"=B{r}+B{r + 1}").font = Font(bold=True)
    ws.cell(row=r + 5, column=1, value="Otros datos (no suman en la caja)").font = Font(bold=True)
    fila_extra = {campo: k for k, (campo, _) in enumerate(EXTRAS, start=r + 6)}
    for campo, nombre in EXTRAS:
        extra = datos["extras"].get(campo)
        ws.cell(row=fila_extra[campo], column=1, value=nombre)
        ws.cell(row=fila_extra[campo], column=2, value=extra["centavos"] / 100 if extra else None)
    ws.cell(row=fila_extra["saldo"], column=2, value=f"=B{fila_extra['deposito']}-B{fila_extra['retiro']}")
    ws.cell(row=r + 3, column=1, value="Total + bajada")
    ws.cell(row=r + 3, column=2, value=f"=B{r + 2}+B{fila_extra['bajada']}")
    for fila in ws.iter_rows(min_row=4, min_col=2, max_col=col_total):
        for c in fila:
            c.number_format = FORMATO_MONEDA

    hs = wb.create_sheet("Historial")
    hs.append(["Fecha", "Turno", "Turno actual", "Total al cierre"])
    estilo_encabezado(hs, 1, 4)
    for i, ancho in enumerate((12, 10, 16, 16), start=1):
        hs.column_dimensions[get_column_letter(i)].width = ancho
    for h in reversed(historial(db)):
        hs.append([h["caja"][0], TURNOS[h["caja"][1]][1], h["actual"] / 100, h["cierre"] / 100])
        hs.cell(row=hs.max_row, column=1).number_format = "dd/mm/yyyy"
        for col in (3, 4):
            hs.cell(row=hs.max_row, column=col).number_format = FORMATO_MONEDA

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"caja_{caja[0].isoformat()}_{TURNOS[caja[1]][0]}.xlsx",
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
    listas = {
        f"{tabla}{sufijo}": db.execute(f"SELECT * FROM {tabla} WHERE activo = ? ORDER BY orden, id", (activo,)).fetchall()
        for tabla in ("filas", "columnas")
        for sufijo, activo in (("", 1), ("_quitadas", 0))
    }
    return render_template(
        "admin.html",
        usuarios=db.execute("SELECT * FROM usuarios ORDER BY es_admin DESC, nombre").fetchall(),
        permisos=permisos,
        **listas,
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
        # el formulario solo muestra las activas: las quitadas conservan su permiso (y su historial)
        db.execute(
            f"DELETE FROM permisos_{tabla} WHERE usuario_id = ? AND {campo} IN (SELECT id FROM {tabla} WHERE activo = 1)",
            (uid,),
        )
        db.executemany(
            f"INSERT INTO permisos_{tabla} (usuario_id, {campo}) SELECT ?, id FROM {tabla} WHERE id = ? AND activo = 1",
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
    existente = db.execute(f"SELECT * FROM {tabla} WHERE nombre = ?", (nombre,)).fetchone()
    if existente and existente["activo"]:
        flash(f"Ya existe una {tabla[:-1]} llamada «{existente['nombre']}».", "error")
        return volver_admin(tabla)
    if existente:
        db.execute(f"UPDATE {tabla} SET activo = 1, orden = ? WHERE id = ?", (orden, existente["id"]))
        flash(f"«{existente['nombre']}» se volvió a agregar, con su historial.", "ok")
    elif tabla == "filas":
        db.execute("INSERT INTO filas (nombre, orden) VALUES (?, ?)", (nombre, orden))
        flash(f"«{nombre}» agregada. Acordate de asignarla a los usuarios que la tengan que ver.", "ok")
    else:
        signo = -1 if request.form.get("signo") == "-1" else 1
        db.execute("INSERT INTO columnas (nombre, signo, orden) VALUES (?, ?, ?)", (nombre, signo, orden))
        flash(f"«{nombre}» agregada. Acordate de asignarla a los usuarios que la tengan que ver.", "ok")
    db.commit()
    return volver_admin(tabla)


@app.post("/admin/<any(filas, columnas):tabla>/<int:item_id>/eliminar")
def admin_eliminar(tabla, item_id):
    """Quita una fila o columna. Si ya tiene montos, se oculta para no cambiar los saldos."""
    db = get_db()
    item = db.execute(f"SELECT * FROM {tabla} WHERE id = ?", (item_id,)).fetchone()
    if not item:
        abort(404)
    campo = "fila_id" if tabla == "filas" else "columna_id"
    if db.execute(f"SELECT 1 FROM celdas WHERE {campo} = ? LIMIT 1", (item_id,)).fetchone():
        db.execute(f"UPDATE {tabla} SET activo = 0 WHERE id = ?", (item_id,))
        flash(f"«{item['nombre']}» se quitó. Sus montos de cajas anteriores se conservan para no cambiar los saldos.", "ok")
    else:
        db.execute(f"DELETE FROM {tabla} WHERE id = ?", (item_id,))
        flash(f"«{item['nombre']}» se eliminó.", "ok")
    db.commit()
    return volver_admin(tabla)


@app.post("/admin/<any(filas, columnas):tabla>/<int:item_id>/reactivar")
def admin_reactivar(tabla, item_id):
    db = get_db()
    orden = db.execute(f"SELECT COALESCE(MAX(orden), -1) + 1 FROM {tabla}").fetchone()[0]
    if db.execute(f"UPDATE {tabla} SET activo = 1, orden = ? WHERE id = ?", (orden, item_id)).rowcount == 0:
        abort(404)
    db.commit()
    flash("Se volvió a agregar, con su historial.", "ok")
    return volver_admin(tabla)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=os.environ.get("DEBUG") == "1")
