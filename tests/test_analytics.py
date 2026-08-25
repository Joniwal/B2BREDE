# -*- coding: utf-8 -*-
import unittest
from datetime import datetime
from unittest.mock import patch

from excel_client import DataClient


class FixedDateTime(datetime):
    @classmethod
    def today(cls):
        return cls(2026, 8, 25, 12, 0, 0)


class AnalyticsTests(unittest.TestCase):
    def setUp(self):
        self.rows = [
            {
                "STATUS": " CONCLUÍDO ",
                "DATACONCLUSAO": "2026-08-24",
                "DATAAGENDAMENTO": "2026-07-10",
                "EXECUTADOPOR": "Ana",
                "CLIENTE": "Cliente A",
                "OBSERVACAO": "Aguardando liberação do cliente",
                "TIPOCABO": "AS80",
                "METRAGEM": "120",
            },
            {
                "STATUS": "CONCLUIDO",
                "DATACONCLUSAO": "2026-08-24",
                "DATAAGENDAMENTO": "2026-08-02",
                "EXECUTADOPOR": "Ana",
                "CLIENTE": "Cliente B",
            },
            {
                "STATUS": "Concluído",
                "DATACONCLUSAO": "2026-08-23",
                "DATAAGENDAMENTO": "2026-08-01",
                "EXECUTADOPOR": "Bruno",
                "CLIENTE": "Cliente C",
            },
            {
                "STATUS": "NOVO",
                "DATACONCLUSAO": "2026-08-24",
                "DATAAGENDAMENTO": "2026-08-03",
                "EXECUTADOPOR": "Carlos",
                "CLIENTE": "Cliente D",
            },
            {
                "STATUS": "CONCLUIDO",
                "DATACONCLUSAO": "2026-07-31",
                "DATAAGENDAMENTO": "2026-08-04",
                "EXECUTADOPOR": "Dora",
                "CLIENTE": "Cliente E",
            },
        ]
        self.client = DataClient()
        self.client._excel_read_all = lambda: self.rows

    @patch("excel_client.datetime", FixedDateTime)
    def test_dashboard_card_counts_only_previous_day(self):
        data = self.client.dashboard_aggregates()

        self.assertEqual(data["kpis"]["concluidos_label"], "Concluídos ontem")
        self.assertEqual(data["kpis"]["concluidos_total"], 2)

    @patch("excel_client.datetime", FixedDateTime)
    def test_dashboard_executor_counts_only_current_month_conclusions(self):
        data = self.client.dashboard_aggregates()

        self.assertEqual(data["por_executadopor"]["labels"], ["Ana", "Bruno"])
        self.assertEqual(data["por_executadopor"]["data"], [2, 1])

    def test_analytics_timeline_and_executor_share_completed_base(self):
        data = self.client.analytics(ano=2026, mes=8)

        self.assertEqual(data["concluidos_timeline"]["data"], [3])
        self.assertEqual(data["por_executado_por"]["labels"], ["Ana", "Bruno"])
        self.assertEqual(data["por_executado_por"]["data"], [2, 1])

    def test_executor_export_contains_only_completed_rows_by_completion_date(self):
        rows = self.client.analytics_export(
            ano=2026,
            mes=8,
            grafico="por_executado_por",
        )

        self.assertEqual(len(rows), 3)
        self.assertTrue(all(row["EXECUTADOPOR"] in {"Ana", "Bruno"} for row in rows))

    def test_items_by_date_returns_observation_instead_of_cable_data(self):
        rows = self.client.items_by_date("2026-08-24", date_field="DATACONCLUSAO")

        self.assertEqual(rows[0]["OBSERVACAO"], "Aguardando liberação do cliente")
        self.assertNotIn("TIPOCABO", rows[0])
        self.assertNotIn("METRAGEM", rows[0])


if __name__ == "__main__":
    unittest.main()
