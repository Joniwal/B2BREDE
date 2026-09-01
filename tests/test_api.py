from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from app import create_app
from sharepoint_client import ExcelDataStore, FIELDS


@pytest.fixture()
def client(tmp_path: Path):
    wb = Workbook()
    wb.active.title = "Dados"
    wb.active.append(["_ITEM_ID", *FIELDS])
    wb.save(tmp_path / "test.xlsx")
    wb.close()
    store = ExcelDataStore(tmp_path / "test.xlsx")
    store.create_item(
        {
            **{field: "" for field in FIELDS},
            "IDCLIENTE": "TESTE-001",
            "CLIENTE": "Cliente Teste",
            "CIDADE": "Campinas",
            "STATUS": "Agendado",
            "EXECUTADOPOR": "Equipe A",
            "DATAAGENDAMENTO": "2026-07-20",
        }
    )
    app = create_app(
        {"TESTING": True, "LOG_DIR": str(tmp_path / "logs")},
        data_store=store,
    )
    return app.test_client()


def test_health_and_page(client):
    assert client.get("/").status_code == 200
    health = client.get("/api/health").get_json()
    assert health["success"] is True
    assert health["data"]["mode"] == "Excel local"


def test_list_filter_sort_and_dashboard(client):
    response = client.get(
        "/api/items?cidade=Campinas&page=1&page_size=10&sort=CLIENTE:asc"
    )
    assert response.status_code == 200
    body = response.get_json()["data"]
    assert body["pagination"]["total"] == 1
    assert body["items"][0]["IDCLIENTE"] == "TESTE-001"

    dashboard = client.get("/api/dashboard?status=Agendado").get_json()["data"]
    assert dashboard["kpis"]["total"] == 1
    assert dashboard["kpis"]["scheduled"] == 1


def test_full_crud(client):
    payload = {field: "" for field in FIELDS}
    payload.update(
        IDCLIENTE="TESTE-002",
        CLIENTE="Novo Cliente",
        CIDADE="Santos",
        STATUS="Pendente",
        DATAAGENDAMENTO="2026-07-21",
    )
    created_response = client.post("/api/items", json=payload)
    assert created_response.status_code == 201
    created = created_response.get_json()["data"]

    updated_response = client.patch(
        f"/api/items/{created['id']}", json={"STATUS": "Concluído"}
    )
    assert updated_response.status_code == 200
    assert updated_response.get_json()["data"]["STATUS"] == "Concluído"

    assert client.delete(f"/api/items/{created['id']}").status_code == 200
    assert client.get(f"/api/items/{created['id']}").status_code == 404


def test_validation_errors_are_json(client):
    invalid = client.post(
        "/api/items",
        json={"IDCLIENTE": "X", "CLIENTE": "Teste", "DATAAGENDAMENTO": "20/07/2026"},
    )
    assert invalid.status_code == 422
    assert invalid.get_json()["error"]["code"] == "VALIDATION_ERROR"

    bad_sort = client.get("/api/items?sort=CAMPO_INEXISTENTE:asc")
    assert bad_sort.status_code == 422
    assert bad_sort.get_json()["error"]["code"] == "INVALID_SORT"
