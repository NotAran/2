import io

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
    assert list(ids("columnas")) == ["Ingresos", "Gastos"]


def test_login(client):
    crear_admin(client)
    post(client, "/salir")
    assert client.get("/").headers["Location"].endswith("/login")
    assert "incorrectos" in ingresar(client, "admin", "mal").get_data(as_text=True)
    assert "Planilla" in ingresar(client, "admin", "secreto1").get_data(as_text=True)


def test_post_sin_csrf_es_rechazado(client):
    crear_admin(client)
    assert client.post("/admin/filas", data={"nombre": "X"}).status_code == 400


def test_admin_agrega_y_quita_filas_y_columnas(client):
    crear_admin(client)
    post(client, "/admin/filas", nombre="Efectivo")
    post(client, "/admin/columnas", nombre="Ahorro", signo="-1")
    assert "Efectivo" in ids("filas") and "Ahorro" in ids("columnas")
    assert "Ya existe" in post(client, "/admin/filas", nombre="efectivo").get_data(as_text=True)

    post(client, "/guardar", **{f"c_{ids('filas')['Efectivo']}_{ids('columnas')['Ingresos']}": "100"})
    post(client, f"/admin/filas/{ids('filas')['Efectivo']}/eliminar")
    post(client, f"/admin/columnas/{ids('columnas')['Ahorro']}/eliminar")
    assert "Efectivo" not in ids("filas") and "Ahorro" not in ids("columnas")
    with montos.app.app_context():
        assert montos.get_db().execute("SELECT COUNT(*) FROM celdas").fetchone()[0] == 0  # se borran en cascada


def test_usuario_ve_y_carga_solo_lo_asignado(client):
    crear_admin(client)
    post(client, "/admin/usuarios", nombre="juan", clave="clave123")
    post(client, "/admin/usuarios", nombre="ana", clave="clave456")
    f, c = ids("filas"), ids("columnas")
    uid = ids("usuarios")["juan"]
    post(client, f"/admin/usuarios/{uid}/permisos", filas=[f["Mercado"], f["Lemon"]], columnas=[c["Ingresos"]])

    ingresar(client, "juan", "clave123")
    pagina = client.get("/").get_data(as_text=True)
    assert "Mercado" in pagina and "Lemon" in pagina and "Brubank" not in pagina
    assert "Gastos" not in pagina
    assert client.get("/admin").status_code == 403

    post(client, "/guardar", **{
        f"c_{f['Mercado']}_{c['Ingresos']}": "1.500",
        f"c_{f['Brubank']}_{c['Ingresos']}": "999",  # no asignada: se ignora
        f"c_{f['Mercado']}_{c['Gastos']}": "999",    # no asignada: se ignora
    })
    with montos.app.app_context():
        celdas = montos.get_db().execute("SELECT * FROM celdas").fetchall()
    assert [(r["fila_id"], r["columna_id"], r["centavos"], r["actualizado_por"]) for r in celdas] == [
        (f["Mercado"], c["Ingresos"], 150000, "juan")
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


# --- Planilla y cálculos -----------------------------------------------------

def test_totales_con_signo(client):
    crear_admin(client)
    f, c = ids("filas"), ids("columnas")
    post(client, "/guardar", **{
        f"c_{f['Mercado']}_{c['Ingresos']}": "1000",
        f"c_{f['Mercado']}_{c['Gastos']}": "250,50",
        f"c_{f['Lemon']}_{c['Ingresos']}": "500",
        f"c_{f['Prex']}_{c['Gastos']}": "abc",
    })
    with montos.app.app_context():
        db = montos.get_db()
        d = montos.armar_planilla(db, db.execute("SELECT * FROM usuarios").fetchone())
    assert d["total_fila"][f["Mercado"]] == 74950
    assert d["total_columna"][c["Ingresos"]] == 150000
    assert d["total_columna"][c["Gastos"]] == 25050
    assert d["total"] == 124950

    pagina = client.get("/").get_data(as_text=True)
    assert "$1.249,50" in pagina

    # dejar la celda vacía la borra
    post(client, "/guardar", **{f"c_{f['Lemon']}_{c['Ingresos']}": ""})
    with montos.app.app_context():
        assert montos.get_db().execute("SELECT COUNT(*) FROM celdas").fetchone()[0] == 2


def test_error_de_monto_se_informa(client):
    crear_admin(client)
    f, c = ids("filas"), ids("columnas")
    r = post(client, "/guardar", **{f"c_{f['Prex']}_{c['Gastos']}": "abc"})
    assert "Prex / Gastos" in r.get_data(as_text=True)


def test_exportar_excel(client):
    crear_admin(client)
    f, c = ids("filas"), ids("columnas")
    post(client, "/guardar", **{f"c_{f['Mercado']}_{c['Ingresos']}": "1000", f"c_{f['Mercado']}_{c['Gastos']}": "250"})
    r = client.get("/exportar.xlsx")
    assert r.status_code == 200
    ws = load_workbook(io.BytesIO(r.data))["Planilla"]
    assert [ws.cell(row=1, column=i).value for i in range(1, 5)] == ["Cuenta", "Ingresos (+)", "Gastos (−)", "Total"]
    assert ws["A2"].value == "Mercado" and ws["B2"].value == 1000 and ws["C2"].value == 250
    assert ws["D2"].value == "=+B2-C2"
    assert ws["A11"].value == "Total" and ws["B11"].value == "=SUM(B2:B10)"
