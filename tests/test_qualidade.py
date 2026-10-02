import unittest
from copy import deepcopy
from pathlib import Path
from unittest.mock import patch
from urllib.parse import unquote

import api
from app import create_app
from excel_client import DataClient, DataClientError, FIELDS


class FakeQualityClient:
    def __init__(self):
        self.rows = [
            self._row("Q-001", "AÇÃO DE QUALIDADE", "Concluído", "GPON"),
            self._row("Q-002", "ACAO DE QUALIDADE", "AGENDADO", "ERB"),
            self._row("R-001", "REPARO", "Concluído", "GPON"),
        ]
        self.last_filters = None

    @staticmethod
    def _row(item_id, activity, status, technology):
        row = {field: "" for field in FIELDS}
        row.update({
            "IDCLIENTE": item_id,
            "CLIENTE": f"CLIENTE {item_id}",
            "ATIVIDADE": activity,
            "STATUS": status,
            "TECNOLOGIA": technology,
            "DATAAGENDAMENTO": "2026-09-29",
        })
        return row

    @staticmethod
    def _normalized(value):
        return api._normalize_text(value)

    def _filtered(self, filters):
        self.last_filters = dict(filters or {})
        rows = self.rows
        activity = (filters or {}).get("atividade")
        if activity:
            rows = [row for row in rows if self._normalized(row["ATIVIDADE"]) == self._normalized(activity)]
        technology = (filters or {}).get("tecnologia")
        if technology:
            rows = [row for row in rows if row["TECNOLOGIA"] == technology]
        return [deepcopy(row) for row in rows]

    def list_items(self, filters=None, page=1, page_size=20, sort=None):
        rows = self._filtered(filters)
        start = (page - 1) * page_size
        return {"items": rows[start:start + page_size], "total": len(rows), "page": page, "page_size": page_size}

    def export_items(self, filters=None, sort=None):
        return self._filtered(filters)

    def get_item(self, item_id):
        item_id = unquote(item_id)
        for row in self.rows:
            if row["IDCLIENTE"] == item_id:
                return deepcopy(row)
        raise DataClientError("Registro não encontrado.", status_code=404)

    def create_item(self, payload):
        row = {field: payload.get(field, "") for field in FIELDS}
        self.rows.append(row)
        return deepcopy(row)

    def update_item(self, item_id, payload):
        current = self.get_item(item_id)
        for row in self.rows:
            if row["IDCLIENTE"] == current["IDCLIENTE"]:
                row.update(payload)
                return deepcopy(row)
        raise AssertionError("registro deveria existir")

    def delete_item(self, item_id):
        current = self.get_item(item_id)
        self.rows = [row for row in self.rows if row["IDCLIENTE"] != current["IDCLIENTE"]]
        return {"deleted": item_id}

    @staticmethod
    def status_arquivo():
        return {"caminho": "C:/OneDrive/REDE_B2B.xlsx"}


class QualidadeTests(unittest.TestCase):
    def setUp(self):
        self.client = FakeQualityClient()
        self.patch = patch.object(api, "data_client", self.client)
        self.patch.start()
        app = create_app()
        app.config["TESTING"] = True
        self.http = app.test_client()

    def tearDown(self):
        self.patch.stop()

    def test_page_and_navigation_are_available(self):
        response = self.http.get("/qualidade")
        self.assertEqual(response.status_code, 200)
        html = response.get_data(as_text=True)
        self.assertIn("Ação de Qualidade", html)
        self.assertIn('id="qChartConclusao"', html)
        self.assertIn('id="qualidadeModal"', html)
        for route in ("/", "/dashboard", "/ativacao", "/reparo"):
            page = self.http.get(route)
            self.assertEqual(page.status_code, 200)
            self.assertIn('href="/qualidade"', page.get_data(as_text=True))
        self.assertNotIn('id="quickStatusFilter"', self.http.get("/").get_data(as_text=True))

    def test_api_always_scopes_records_and_dashboard_to_quality(self):
        listed = self.http.get("/api/qualidade/records?tecnologia=GPON").get_json()["data"]
        self.assertEqual(listed["total"], 1)
        self.assertEqual(listed["items"][0]["IDCLIENTE"], "Q-001")
        self.assertEqual(self.client.last_filters["atividade"], "AÇÃO DE QUALIDADE")

        dashboard = self.http.get("/api/qualidade/dashboard").get_json()["data"]
        self.assertEqual(dashboard["total"], 2)
        self.assertEqual(dashboard["concluidas"], 1)
        self.assertEqual(dashboard["nao_concluidas"], 1)

    def test_create_and_update_force_quality_activity(self):
        created = self.http.post("/api/qualidade/records", json={
            "IDCLIENTE": "Q-003", "CLIENTE": "NOVO", "ATIVIDADE": "REPARO",
        })
        self.assertEqual(created.status_code, 201)
        self.assertEqual(created.get_json()["data"]["ATIVIDADE"], "AÇÃO DE QUALIDADE")

        changed = self.http.patch("/api/qualidade/records/Q-001", json={"ATIVIDADE": "REPARO", "STATUS": "PCC"})
        self.assertEqual(changed.status_code, 200)
        self.assertEqual(changed.get_json()["data"]["ATIVIDADE"], "AÇÃO DE QUALIDADE")
        self.assertEqual(self.http.patch("/api/qualidade/records/R-001", json={"STATUS": "PCC"}).status_code, 404)

    def test_data_client_activity_filter_ignores_accents(self):
        rows = [
            {"ATIVIDADE": "AÇÃO DE QUALIDADE", "IDCLIENTE": "1"},
            {"ATIVIDADE": "REPARO", "IDCLIENTE": "2"},
        ]
        filtered = DataClient()._apply_filters(rows, {"atividade": "ACAO DE QUALIDADE"})
        self.assertEqual([row["IDCLIENTE"] for row in filtered], ["1"])


if __name__ == "__main__":
    unittest.main()
