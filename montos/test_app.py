import io
from datetime import date

import pytest
from openpyxl import load_workbook

import app as montos


@pytest.fixture
def client(tmp_path):
    montos.app.config.update(TESTING=True, DATABASE=str(tmp_path / "test.db"))
    with montos.app.test_client() as c:
        yield c


def post(client, url, **datos):
    """POST con el token CSRF de la sesión."""
    client.get("/login")  # asegura que la sesión tenga token
    with client.session_transaction() as s:
        datos["csrf"] = s["csrf"]
    return client.post(url, data=datos, follow_redirects=True)


def crear_admin(client):
    return post(client, "/setup", nombre="admin", clave="secreto1")


def ingresar(client, nombre, clave):
    post(client, "/salir")
    return post(client, "/login", nombre=nombre, clave=clave)


def cargar(client, fecha, turno, **celdas):
    """Guarda montos en una caja. celdas: {"Fila__Columna": "monto"}."""
    f, c = ids("filas"), ids("columnas")
    datos = {}
    for clave, monto in celdas.items():
        fila, columna = clave.split("__")
        datos[f"c_{f[fila]}_{c[columna]}"] = monto
    return post(client, "/guardar", fecha=fecha, turno=turno, **datos)


def caja(fecha, turno, usuario="admin"):
    with montos.app.app_context():
        db = montos.get_db()
        u = db.execute("SELECT * FROM usuarios WHERE nombre = ?", (usuario,)).fetchone()
        return montos.armar_caja(db, u, (date.fromisoformat(fecha), montos.TURNO_POR_SLUG[turno]))


def ids(tabla):
    with montos.app.app_context():
        return {r["nombre"]: r["id"] for r in montos.get_db().execute(f"SELECT id, nombre FROM {tabla} ORDER BY id")}


# --- Formato de montos ---------------------------------------------------------

@pytest.mark.parametrize("texto,centavos", [
    ("1234.56", 123456), ("1.234,56", 123456), ("1,234.56", 123456),
    ("$ 500", 50000), ("12,5", 1250), ("1,000", 100000), ("0.01", 1),
    ("9.850", 985000), ("1.234.567", 123456700), ("45.780,50", 4578050), ("3.5", 350),
    ("-500", -50000), ("-1.234,56", -123456), ("0", 0),
])
def test_parsear_monto(texto, centavos):
    assert montos.parsear_monto(texto) == centavos


@pytest.mark.parametrize("texto", ["", "abc", "--5", "1.2345", "nan", "inf"])
def test_parsear_monto_invalido(texto):
    with pytest.raises(ValueError):
        montos.parsear_monto(texto)


def test_formatos():
    assert montos.formato_dinero(123456789) == "$1.234.567,89"
    assert montos.formato_dinero(-50) == "-$0,50"
    # lo que se muestra en la celda se vuelve a leer igual
    for centavos in (0, 5, 123400, -98765432):
        assert montos.parsear_monto(montos.formato_numero(centavos)) == centavos


# --- Login y admin -----------------------------------------------------------

def test_primera_vez_pide_crear_admin(client):
    assert client.get("/").headers["Location"].endswith("/setup")
    r = crear_admin(client)
    assert "sos el admin" in r.get_data(as_text=True)
    # el setup ya no está disponible
    assert client.get("/setup").headers["Location"].endswith("/login")


def test_filas_y_columnas_iniciales(client):
    crear_admin(client)
    assert list(ids("filas")) == list(montos.FILAS_INICIALES)
    assert list(ids("columnas")) == ["Paco", "Antonio", "Ibra"]


def test_login(client):
    crear_admin(client)
    post(client, "/salir")
    assert client.get("/").headers["Location"].endswith("/login")
    assert "incorrectos" in ingresar(client, "admin", "mal").get_data(as_text=True)
    assert "Cajas por turno" in ingresar(client, "admin", "secreto1").get_data(as_text=True)


def test_post_sin_csrf_es_rechazado(client):
    crear_admin(client)
    assert client.post("/admin/filas", data={"nombre": "X"}).status_code == 400


def test_admin_agrega_y_quita_filas_y_columnas(client):
    crear_admin(client)
    post(client, "/admin/filas", nombre="Efectivo")
    post(client, "/admin/columnas", nombre="Retiros", signo="-1")
    assert "Efectivo" in ids("filas") and "Retiros" in ids("columnas")
    assert "Ya existe" in post(client, "/admin/filas", nombre="efectivo").get_data(as_text=True)

    # sin montos: se elimina de verdad
    post(client, f"/admin/filas/{ids('filas')['Efectivo']}/eliminar")
    assert "Efectivo" not in ids("filas")

    # con montos: se oculta y sus montos siguen contando en los saldos
    cargar(client, "2026-10-08", "noche", Mercado__Paco="1000", Mercado__Retiros="300")
    post(client, f"/admin/columnas/{ids('columnas')['Retiros']}/eliminar")
    with montos.app.app_context():
        r = montos.get_db().execute("SELECT activo FROM columnas WHERE nombre = 'Retiros'").fetchone()
    assert r["activo"] == 0
    d = caja("2026-10-08", "manana")
    assert "Retiros" not in [c["nombre"] for c in d["columnas"]]
    assert d["anterior"][ids("filas")["Mercado"]] == 70000
    # en la caja donde tiene montos se sigue viendo
    assert "Retiros" in [c["nombre"] for c in caja("2026-10-08", "noche")["columnas"]]

    # volver a agregarla con el mismo nombre la reactiva
    post(client, "/admin/columnas", nombre="retiros")
    assert "Retiros" in [c["nombre"] for c in caja("2026-10-08", "manana")["columnas"]]


def test_usuario_ve_y_carga_solo_lo_asignado(client):
    crear_admin(client)
    post(client, "/admin/usuarios", nombre="paco", clave="clave123")
    post(client, "/admin/usuarios", nombre="ana", clave="clave456")
    f, c = ids("filas"), ids("columnas")
    uid = ids("usuarios")["paco"]
    post(client, f"/admin/usuarios/{uid}/permisos", filas=[f["Mercado"], f["Lemon"]], columnas=[c["Paco"]])

    ingresar(client, "paco", "clave123")
    pagina = client.get("/?fecha=2026-10-08&turno=noche").get_data(as_text=True)
    assert "Mercado" in pagina and "Lemon" in pagina and "Brubank" not in pagina
    assert "Antonio" not in pagina
    assert client.get("/admin").status_code == 403

    cargar(client, "2026-10-08", "noche", Mercado__Paco="1.500", Brubank__Paco="999", Mercado__Antonio="999")
    with montos.app.app_context():
        celdas = montos.get_db().execute("SELECT * FROM celdas").fetchall()
    assert [(r["fila_id"], r["columna_id"], r["centavos"], r["actualizado_por"]) for r in celdas] == [
        (f["Mercado"], c["Paco"], 150000, "paco")  # lo no asignado se ignora
    ]

    # un usuario sin asignaciones no ve nada
    ingresar(client, "ana", "clave456")
    assert "todavía no te asignó" in client.get("/").get_data(as_text=True)


def test_eliminar_usuario_y_no_al_admin(client):
    crear_admin(client)
    post(client, "/admin/usuarios", nombre="juan", clave="clave123")
    post(client, f"/admin/usuarios/{ids('usuarios')['juan']}/eliminar")
    assert "juan" not in ids("usuarios")
    assert post(client, f"/admin/usuarios/{ids('usuarios')['admin']}/eliminar").status_code == 404


# --- Cajas y cálculos ----------------------------------------------------------

def test_mover_entre_cajas():
    d = date(2026, 10, 8)
    assert montos.mover((d, 0), 1) == (d, 1)
    assert montos.mover((d, 2), 1) == (date(2026, 10, 9), 0)   # tarde -> noche del día siguiente
    assert montos.mover((d, 0), -1) == (date(2026, 10, 7), 2)  # noche -> tarde del día anterior
    assert montos.nombre_caja((d, 1)) == "Jueves 08/10/2026 – Mañana"


def test_turno_anterior_arrastra_el_cierre_de_la_caja_previa(client):
    crear_admin(client)
    m = ids("filas")["Mercado"]
    cargar(client, "2026-10-08", "noche", Mercado__Paco="1000", Mercado__Antonio="500", Lemon__Ibra="200")
    noche = caja("2026-10-08", "noche")
    assert noche["anterior"][m] == 0
    assert noche["actual"][m] == 150000
    assert noche["totales"] == {"anterior": 0, "actual": 170000, "caja": 170000}

    cargar(client, "2026-10-08", "manana", Mercado__Ibra="-300")
    manana = caja("2026-10-08", "manana")
    assert manana["anterior"][m] == 150000          # cierre de la noche
    assert manana["total_fila"][m] == 120000
    assert manana["totales"]["caja"] == 140000

    # la tarde no tuvo movimientos; la noche siguiente arrastra igual
    siguiente = caja("2026-10-09", "noche")
    assert siguiente["anterior"][m] == 120000 and siguiente["totales"]["actual"] == 0

    # corregir una caja vieja actualiza los saldos de las siguientes
    cargar(client, "2026-10-08", "noche", Mercado__Paco="2000")
    assert caja("2026-10-09", "noche")["anterior"][m] == 220000

    pagina = client.get("/?fecha=2026-10-08&turno=manana").get_data(as_text=True)
    assert "$2.400,00" in pagina and "Jueves 08/10/2026 – Mañana" in pagina


def test_columna_que_resta(client):
    crear_admin(client)
    post(client, "/admin/columnas", nombre="Retiros", signo="-1")
    cargar(client, "2026-10-08", "noche", Mercado__Paco="1000", Mercado__Retiros="250,50")
    assert caja("2026-10-08", "noche")["total_fila"][ids("filas")["Mercado"]] == 74950


def test_sin_parametros_muestra_la_ultima_caja_con_datos(client):
    crear_admin(client)
    cargar(client, "2026-10-05", "tarde", Mercado__Paco="10")
    assert "Lunes 05/10/2026 – Tarde" in client.get("/").get_data(as_text=True)


def test_vaciar_celda_la_borra(client):
    crear_admin(client)
    cargar(client, "2026-10-08", "noche", Mercado__Paco="10", Lemon__Paco="20")
    cargar(client, "2026-10-08", "noche", Lemon__Paco="")
    with montos.app.app_context():
        assert montos.get_db().execute("SELECT COUNT(*) FROM celdas").fetchone()[0] == 1


def test_error_de_monto_se_informa(client):
    crear_admin(client)
    r = cargar(client, "2026-10-08", "noche", Prex__Ibra="abc")
    assert "Prex / Ibra" in r.get_data(as_text=True)


def test_guardar_sin_caja_valida(client):
    crear_admin(client)
    assert post(client, "/guardar", fecha="mal", turno="noche").status_code == 400


def test_exportar_excel(client):
    crear_admin(client)
    cargar(client, "2026-10-08", "noche", Mercado__Paco="1000")
    cargar(client, "2026-10-08", "manana", Mercado__Antonio="250")
    r = client.get("/exportar.xlsx?fecha=2026-10-08&turno=manana")
    assert r.status_code == 200
    wb = load_workbook(io.BytesIO(r.data))
    ws = wb["Caja"]
    assert ws["A1"].value == "Caja: Jueves 08/10/2026 – Mañana"
    assert [ws.cell(row=3, column=i).value for i in range(1, 8)] == [
        "Cuenta", "Turno anterior", "Paco", "Antonio", "Ibra", "Turno actual", "Total caja"
    ]
    assert ws["A4"].value == "Mercado" and ws["B4"].value == 1000 and ws["D4"].value == 250
    assert ws["F4"].value == "=+C4+D4+E4" and ws["G4"].value == "=B4+F4"
    assert ws["A13"].value == "Total" and ws["G13"].value == "=SUM(G4:G12)"
    hs = wb["Historial"]
    assert [c.value for c in hs[3]][1:] == ["Mañana", 250, 1250]
