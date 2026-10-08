import io

import pytest
from openpyxl import load_workbook

import app as montos


@pytest.fixture
def client(tmp_path):
    montos.app.config.update(TESTING=True, DATABASE=str(tmp_path / "test.db"))
    with montos.app.test_client() as c:
        yield c


def agregar(client, **datos):
    base = {"tipo": "gasto", "fecha": "2026-10-01", "categoria": "Comida", "descripcion": "", "monto": "10"}
    base.update(datos)
    return client.post("/agregar", data=base, follow_redirects=True)


@pytest.mark.parametrize("texto,centavos", [
    ("1234.56", 123456), ("1.234,56", 123456), ("1,234.56", 123456),
    ("$ 500", 50000), ("12,5", 1250), ("1,000", 100000), ("0.01", 1),
    ("9.850", 985000), ("1.234.567", 123456700), ("45.780,50", 4578050), ("3.5", 350),
])
def test_parsear_monto(texto, centavos):
    assert montos.parsear_monto(texto) == centavos


@pytest.mark.parametrize("texto", ["", "abc", "0", "-5", "1.2345", "nan", "inf"])
def test_parsear_monto_invalido(texto):
    with pytest.raises(ValueError):
        montos.parsear_monto(texto)


def test_formato_dinero():
    assert montos.formato_dinero(123456789) == "$1.234.567,89"
    assert montos.formato_dinero(-50) == "-$0,50"


def test_calculos(client):
    agregar(client, tipo="ingreso", monto="1000", categoria="Sueldo")
    agregar(client, monto="250,50", categoria="Comida")
    agregar(client, monto="100", categoria="Transporte")
    agregar(client, monto="50", fecha="2026-09-15", categoria="Comida")

    with montos.app.app_context():
        movs, tot, cats, meses, _ = montos.calcular(montos.get_db())
        assert tot["ingresos"] == 100000
        assert tot["gastos"] == 40050
        assert tot["balance"] == 59950
        assert tot["cantidad"] == 4
        assert tot["mayor_gasto"] == 25050
        assert tot["gasto_promedio"] == 13350
        assert tot["ahorro_pct"] == 60.0
        assert {c["categoria"]: c["gastos"] for c in cats}["Comida"] == 30050
        assert [m["mes"] for m in meses] == ["2026-10", "2026-09"]

        _, tot_sep, _, _, _ = montos.calcular(montos.get_db(), "2026-09")
        assert tot_sep["gastos"] == 5000 and tot_sep["cantidad"] == 1

    pagina = client.get("/").get_data(as_text=True)
    assert "$599,50" in pagina and "Sueldo" in pagina


def test_monto_invalido_no_se_guarda(client):
    r = agregar(client, monto="hola")
    assert "no es un número válido" in r.get_data(as_text=True)
    r = agregar(client, fecha="2026-13-40")
    assert "Fecha inválida" in r.get_data(as_text=True)
    with montos.app.app_context():
        assert montos.calcular(montos.get_db())[1]["cantidad"] == 0


def test_eliminar(client):
    agregar(client)
    client.post("/eliminar/1")
    assert client.post("/eliminar/1").status_code == 404


def test_exportar_excel(client):
    agregar(client, tipo="ingreso", monto="1000", categoria="Sueldo")
    agregar(client, monto="250", categoria="Comida")
    r = client.get("/exportar.xlsx")
    assert r.status_code == 200
    wb = load_workbook(io.BytesIO(r.data))
    assert wb.sheetnames == ["Movimientos", "Resumen", "Por categoría", "Por mes"]
    ws = wb["Movimientos"]
    assert ws["E2"].value == 1000 and ws["E3"].value == 250
    assert ws["F3"].value == '=IF(D3="gasto",-E3,E3)'
    assert wb["Resumen"]["B2"].value.startswith("=SUMIF(")
    assert wb["Por mes"]["B2"].value == 1000
