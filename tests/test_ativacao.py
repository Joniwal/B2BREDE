import os
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

from ativacao_client import AtivacaoClient, EXCEL_HEADERS
from excel_client import DataClientError


class AtivacaoClientTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.path = Path(self.temporary.name) / "ATIVACAO.xlsx"
        AtivacaoClient._caminho_cache = None
        self.client = AtivacaoClient(self.path)

    def tearDown(self):
        AtivacaoClient._caminho_cache = None
        self.temporary.cleanup()

    @staticmethod
    def payload(**changes):
        values = {
            "cliente": "CLIENTE TESTE",
            "cidade": "CURITIBA",
            "servico": "IP DEDICADO",
            "tecnologia": "GPON",
            "empresa": "VIVO",
            "status": "OK",
            "data_agendamento": date.today().isoformat(),
            "data_execucao": date.today().isoformat(),
            "no_mes": "SIM",
            "tecnico": "TECNICO TESTE",
            "faturado": "SIM",
            "com_rfs": "NÃO",
        }
        values.update(changes)
        return values

    def test_creates_expected_existing_workbook_schema(self):
        self.client.list()
        workbook = load_workbook(self.path)
        sheet = workbook["ATIVACAO"]
        self.assertEqual([sheet.cell(1, col).value for col in range(1, 14)], EXCEL_HEADERS)
        self.assertIn("ATIVACAO", sheet.tables)
        self.assertIn("CIDADES", workbook.sheetnames)
        self.assertEqual(workbook["CIDADES"]["A1"].value, "CIDADES")
        workbook.close()

    def test_create_three_edit_view_and_delete(self):
        first = self.client.create(self.payload(cliente="ALFA"))
        second = self.client.create(self.payload(cliente="BETA", empresa="RS TELECOM", status="NOK"))
        third = self.client.create(self.payload(cliente="GAMA", servico="SIP"))
        self.assertEqual([first["id"], second["id"], third["id"]], [1, 2, 3])
        self.assertEqual(self.client.list()["total"], 3)
        self.assertEqual(self.client.get(2)["cliente"], "BETA")

        updated = self.client.update(2, {"status": "OK", "faturado": "NÃO"})
        self.assertEqual(updated["status"], "OK")
        self.assertEqual(updated["faturado"], "NÃO")
        self.assertEqual(updated["cidade"], "CURITIBA")

        self.assertEqual(self.client.delete(1), {"id": 1, "deleted": True})
        self.assertEqual(self.client.list()["total"], 2)
        with self.assertRaises(DataClientError):
            self.client.get(1)

    def test_manual_id_can_be_created_and_changed_without_duplicates(self):
        created = self.client.create(self.payload(id=2224567, cliente="ID MANUAL"))
        self.assertEqual(created["id"], 2224567)

        with self.assertRaisesRegex(DataClientError, "Já existe uma atividade com o ID 2224567"):
            self.client.create(self.payload(id=2224567, cliente="ID REPETIDO"))

        updated = self.client.update(2224567, {"id": 2224568})
        self.assertEqual(updated["id"], 2224568)
        with self.assertRaises(DataClientError):
            self.client.get(2224567)

        with self.assertRaisesRegex(DataClientError, "ID inválido"):
            self.client.create(self.payload(id="12A"))
        with self.assertRaisesRegex(DataClientError, "Campo ID obrigatório"):
            self.client.create(self.payload(id=""))

    def test_filters_by_execution_and_schedule_dates(self):
        today = date.today()
        yesterday = today - timedelta(days=1)
        self.client.create(self.payload(cliente="EXECUTA HOJE", data_agendamento=yesterday.isoformat()))
        self.client.create(self.payload(cliente="AGENDA HOJE", data_execucao=yesterday.isoformat()))

        execution = self.client.list({"data_execucao": today.isoformat()})
        schedule = self.client.list({"data_agendamento": today.isoformat()})
        self.assertEqual([item["cliente"] for item in execution["items"]], ["EXECUTA HOJE"])
        self.assertEqual([item["cliente"] for item in schedule["items"]], ["AGENDA HOJE"])

    def test_dashboard_requested_distributions(self):
        today = date.today().isoformat()
        self.client.create(self.payload(servico="IP DEDICADO", tecnologia="GPON", faturado="SIM", com_rfs="SIM"))
        self.client.create(self.payload(servico="SIP", tecnologia="ERB", status="NOK", faturado="NÃO", com_rfs="NÃO"))
        dashboard = self.client.dashboard({"data_execucao": today})
        self.assertEqual(dashboard["kpis"]["total"], 2)
        self.assertEqual(dashboard["por_status"]["values"], [1, 1])
        self.assertEqual(dashboard["por_tecnologia"]["values"], [1, 1])
        self.assertEqual(dashboard["por_faturado"]["values"], [1, 1])
        self.assertEqual(dashboard["por_rfs"]["values"], [1, 1])
        self.assertEqual(dashboard["por_data_execucao"], {"labels": [today], "values": [2]})
        self.assertEqual(set(dashboard["por_servico"]["labels"]), {"IP DEDICADO", "SIP"})

    def test_validation_and_dynamic_options(self):
        self.client.list()
        workbook = load_workbook(self.path)
        workbook["CIDADES"].append(["São José dos Pinhais"])
        workbook["CIDADES"].append(["Ponta Grossa"])
        workbook.save(self.path)
        workbook.close()
        self.client.create(self.payload(cliente="FUST", cidade="PINHAIS", servico="SIP", tecnico="EDSON"))
        options = self.client.options()
        self.assertIn("FUST", options["clientes"])
        self.assertIn("PINHAIS", options["cidades"])
        self.assertIn("SÃO JOSÉ DOS PINHAIS", options["cidades"])
        self.assertIn("PONTA GROSSA", options["cidades"])
        self.assertIn("SIP", options["servicos"])
        self.assertEqual(
            options["tecnicos"],
            [
                "ALEXSANDRO NUNES DA SILVA",
                "ANDRE DOS SANTOS CARVALHO",
                "EDSON LUIS DE CRISTO",
                "EMERSON FAVERO RODRIGUES",
                "LUCAS SILVA ANDRADE",
                "MARCOS ROBERTO HOLTMAN",
                "ERENILSON SANT'ANA",
            ],
        )
        with self.assertRaisesRegex(DataClientError, "Campo Cliente obrigatório"):
            self.client.create(self.payload(cliente=""))
        with self.assertRaisesRegex(DataClientError, "Data de Agendamento inválida"):
            self.client.create(self.payload(data_agendamento="31/02/2026"))

    def test_automatic_discovery_finds_file_on_desktop(self):
        desktop = Path(self.temporary.name) / "Desktop"
        desktop.mkdir()
        expected = desktop / "ATIVACAO.xlsx"
        expected.write_bytes(b"arquivo-portatil")

        with (
            patch.dict(
                os.environ,
                {"ATIVACAO_PATH": "", "ATIVACAO_FILENAME": "ATIVACAO.xlsx"},
                clear=False,
            ),
            patch.object(AtivacaoClient, "_candidate_directories", return_value=[desktop]),
            patch(
                "ativacao_client.DataClient._resolver_caminho_excel",
                side_effect=DataClientError("Planilha principal indisponível", 500),
            ),
            patch("ativacao_client._localizar_excel_no_onedrive", return_value=(None, [])),
        ):
            self.assertEqual(AtivacaoClient()._resolve_path(), expected)

    def test_automatic_discovery_rechecks_when_cached_file_was_moved(self):
        old_file = Path(self.temporary.name) / "antigo" / "ATIVACAO.xlsx"
        new_directory = Path(self.temporary.name) / "novo"
        new_directory.mkdir()
        new_file = new_directory / "ATIVACAO.xlsx"
        new_file.write_bytes(b"arquivo-movido")
        AtivacaoClient._caminho_cache = str(old_file)

        with (
            patch.dict(
                os.environ,
                {"ATIVACAO_PATH": "", "ATIVACAO_FILENAME": "ATIVACAO.xlsx"},
                clear=False,
            ),
            patch.object(
                AtivacaoClient, "_candidate_directories", return_value=[new_directory]
            ),
            patch(
                "ativacao_client.DataClient._resolver_caminho_excel",
                side_effect=DataClientError("Planilha principal indisponível", 500),
            ),
            patch("ativacao_client._localizar_excel_no_onedrive", return_value=(None, [])),
        ):
            self.assertEqual(AtivacaoClient()._resolve_path(), new_file)

    def test_automatic_mode_does_not_create_an_empty_file_when_missing(self):
        empty_directory = Path(self.temporary.name) / "vazio"
        empty_directory.mkdir()

        with (
            patch.dict(
                os.environ,
                {"ATIVACAO_PATH": "", "ATIVACAO_FILENAME": "ATIVACAO.xlsx"},
                clear=False,
            ),
            patch.object(
                AtivacaoClient, "_candidate_directories", return_value=[empty_directory]
            ),
            patch(
                "ativacao_client.DataClient._resolver_caminho_excel",
                side_effect=DataClientError("Planilha principal indisponível", 500),
            ),
            patch("ativacao_client._localizar_excel_no_onedrive", return_value=(None, [])),
        ):
            with self.assertRaisesRegex(DataClientError, "não foi encontrado"):
                AtivacaoClient()._ensure_workbook()
            self.assertFalse((empty_directory / "ATIVACAO.xlsx").exists())


if __name__ == "__main__":
    unittest.main()
