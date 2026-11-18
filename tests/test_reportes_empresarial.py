from unittest.mock import MagicMock, patch
from types import SimpleNamespace
import pytest
from fastapi.testclient import TestClient

from app.main import app
from app.core.dependencies import get_current_user
from app.database.connection import get_db
from app.models.role import Role

client = TestClient(app)

ADMIN = 1


def mock_db():
    db = MagicMock()
    db.get.side_effect = lambda model, _id: SimpleNamespace(id_PK=ADMIN, nombre="Administrador") if model == Role else None
    return db


@pytest.fixture(autouse=True)
def limpiar():
    app.dependency_overrides[get_db] = mock_db
    yield
    app.dependency_overrides.clear()


def test_guard_bloquea_plan_profesional():
    app.dependency_overrides[get_current_user] = lambda: SimpleNamespace(
        id_PK=1, rol_id_FK=ADMIN, clinica_id_FK=10, estado="activo"
    )
    with patch("app.services.suscripcion_service.obtener_clinica_con_plan") as mock_plan:
        mock_plan.return_value = (
            SimpleNamespace(id_PK=10, activa=True),
            SimpleNamespace(id_PK=2, nombre="Profesional"),
        )
        res = client.get("/api/reportes/empresarial/datos")
        assert res.status_code == 403
        assert "Plan Empresarial" in res.json()["detail"]


def test_guard_permite_plan_empresarial_y_devuelve_datos():
    app.dependency_overrides[get_current_user] = lambda: SimpleNamespace(
        id_PK=1,
        rol_id_FK=ADMIN,
        clinica_id_FK=10,
        estado="activo",
        clinica=SimpleNamespace(nombre_comercial="Clínica Premium"),
    )
    with patch("app.services.suscripcion_service.obtener_clinica_con_plan") as mock_plan, \
         patch("app.routers.reportes.ReporteService") as MockSrv:
        mock_plan.return_value = (
            SimpleNamespace(id_PK=10, activa=True),
            SimpleNamespace(id_PK=3, nombre="Empresarial"),
        )
        srv_inst = MockSrv.return_value
        srv_inst.resumen_general.return_value = {
            "desde": "2026-10-01",
            "hasta": "2026-10-31",
            "total_citas": 50,
            "citas_atendidas": 45,
            "citas_canceladas": 5,
            "total_salas": 4,
            "horas_ocupadas_totales": 32.5,
            "espera_promedio_minutos": 14.2,
            "total_atenciones_espera": 40,
            "total_medicamentos_stock_bajo": 1,
            "total_recetas": 38,
            "detalle_salas": [],
            "detalle_doctores": [],
            "detalle_espera": [],
            "detalle_inventario": [],
            "detalle_recetas": [],
        }
        srv_inst.calcular_periodo_anterior.return_value = ("2026-08-31", "2026-09-30")
        srv_inst.obtener_comparativa_metricas.return_value = [
            {"indicador": "Citas totales", "valor_actual": 50, "valor_anterior": 40, "variacion_pct": 25.0, "tendencia": "up", "favorable": True}
        ]

        res = client.get("/api/reportes/empresarial/datos")
        assert res.status_code == 200
        data = res.json()
        assert "insights_ia" in data
        assert "score_eficiencia" in data["insights_ia"]
        assert "comparativa" in data
        assert len(data["comparativa"]) > 0


def test_consultor_ia_empresarial():
    app.dependency_overrides[get_current_user] = lambda: SimpleNamespace(
        id_PK=1,
        rol_id_FK=ADMIN,
        clinica_id_FK=10,
        estado="activo",
        clinica=SimpleNamespace(nombre_comercial="Clínica Premium"),
    )
    with patch("app.services.suscripcion_service.obtener_clinica_con_plan") as mock_plan, \
         patch("app.routers.reportes.ReporteService") as MockSrv:
        mock_plan.return_value = (
            SimpleNamespace(id_PK=10, activa=True),
            SimpleNamespace(id_PK=3, nombre="Empresarial"),
        )
        srv_inst = MockSrv.return_value
        srv_inst.resumen_general.return_value = {
            "desde": "2026-10-01",
            "hasta": "2026-10-31",
            "total_citas": 20,
            "citas_atendidas": 18,
            "citas_canceladas": 2,
            "total_salas": 3,
            "horas_ocupadas_totales": 15.0,
            "espera_promedio_minutos": 10.0,
            "total_atenciones_espera": 18,
            "total_medicamentos_stock_bajo": 0,
            "total_recetas": 15,
        }
        srv_inst.calcular_periodo_anterior.return_value = ("2026-08-31", "2026-09-30")
        srv_inst.obtener_comparativa_metricas.return_value = []

        res = client.post(
            "/api/reportes/empresarial/consultor-ia",
            json={"pregunta": "¿Cómo están los tiempos de espera?"},
        )
        assert res.status_code == 200
        assert "respuesta" in res.json()


def test_exportar_reporte_empresarial_pdf_y_excel():
    app.dependency_overrides[get_current_user] = lambda: SimpleNamespace(
        id_PK=1,
        rol_id_FK=ADMIN,
        clinica_id_FK=10,
        estado="activo",
        clinica=SimpleNamespace(nombre_comercial="Clínica Premium"),
    )
    with patch("app.services.suscripcion_service.obtener_clinica_con_plan") as mock_plan, \
         patch("app.routers.reportes.ReporteService") as MockSrv:
        mock_plan.return_value = (
            SimpleNamespace(id_PK=10, activa=True),
            SimpleNamespace(id_PK=3, nombre="Empresarial"),
        )
        srv_inst = MockSrv.return_value
        srv_inst.resumen_general.return_value = {
            "desde": "2026-10-01",
            "hasta": "2026-10-31",
            "total_citas": 10,
            "citas_atendidas": 8,
            "citas_canceladas": 2,
            "total_salas": 2,
            "horas_ocupadas_totales": 8.0,
            "espera_promedio_minutos": 12.0,
            "total_atenciones_espera": 8,
            "total_medicamentos_stock_bajo": 0,
            "total_recetas": 7,
            "detalle_salas": [],
            "detalle_doctores": [],
            "detalle_espera": [],
            "detalle_inventario": [],
            "detalle_recetas": [],
        }
        srv_inst.calcular_periodo_anterior.return_value = ("2026-08-31", "2026-09-30")
        srv_inst.obtener_comparativa_metricas.return_value = []
        srv_inst.citas_por_doctor.return_value = {
            "desde": "2026-10-01",
            "hasta": "2026-10-31",
            "total_citas": 10,
            "detalle": [{"doctor_id": 1, "doctor": "Dr. Carlos", "total": 10, "por_estado": {"atendida": 10}}],
        }

        # PDF
        res_pdf = client.get("/api/reportes/empresarial/citas/exportar?formato=pdf")
        assert res_pdf.status_code == 200
        assert res_pdf.content.startswith(b"%PDF")

        # Excel
        res_excel = client.get("/api/reportes/empresarial/citas/exportar?formato=excel")
        assert res_excel.status_code == 200
        assert res_excel.content.startswith(b"PK")
