# -*- coding: utf-8 -*-
import os
import unittest
from datetime import date, datetime
from unittest.mock import patch

from excel_client import DataClient, _dia_util_anterior


CALENDAR_ENV = {
    "BUSINESS_HOLIDAY_SUBDIV": "SP",
    "BUSINESS_HOLIDAY_INCLUDE_OPTIONAL": "false",
    "BUSINESS_HOLIDAYS": "",
}


class MondayDateTime(datetime):
    @classmethod
    def today(cls):
        return cls(2026, 8, 31, 12, 0, 0)


class BusinessDayTests(unittest.TestCase):
    @patch.dict(os.environ, CALENDAR_ENV)
    def test_skips_saturday_and_sunday(self):
        self.assertEqual(_dia_util_anterior(date(2026, 8, 31)), date(2026, 8, 28))

    @patch.dict(os.environ, CALENDAR_ENV)
    def test_skips_national_holiday(self):
        # 07/09/2026 é segunda-feira e feriado da Independência.
        self.assertEqual(_dia_util_anterior(date(2026, 9, 8)), date(2026, 9, 4))

    @patch.dict(
        os.environ,
        {**CALENDAR_ENV, "BUSINESS_HOLIDAYS": "2026-08-28"},
    )
    def test_skips_custom_holiday(self):
        self.assertEqual(_dia_util_anterior(date(2026, 8, 31)), date(2026, 8, 27))

    @patch("excel_client.datetime", MondayDateTime)
    @patch.dict(os.environ, CALENDAR_ENV)
    def test_dashboard_summary_uses_previous_business_day(self):
        client = DataClient()
        client._excel_read_all = lambda: [
            {
                "STATUS": "CONCLUIDO",
                "DATACONCLUSAO": "2026-08-28",
                "DATAAGENDAMENTO": "2026-08-20",
            },
            {
                "STATUS": "PCC",
                "USUARIO": "CINTIA - 30/08/2026, 18:49",
                "DATAAGENDAMENTO": "2026-08-28",
            },
            {
                "STATUS": "CANCELADO",
                "DATAAGENDAMENTO": "2026-08-28",
            },
            {
                "STATUS": "AGENDADO",
                "DATAAGENDAMENTO": "28/08/2026",
            },
            {
                "STATUS": "INICIADO NAO CONCLUIDO",
                "DATAAGENDAMENTO": "2026-08-28 09:00:00",
            },
            # Não conta: conclusão ocorreu em outra data, mesmo que a data de
            # agendamento seja o dia útil anterior.
            {
                "STATUS": "CONCLUIDO",
                "DATACONCLUSAO": "2026-08-27",
                "DATAAGENDAMENTO": "2026-08-28",
            },
            # Não conta: para PCC e demais status vale DATAAGENDAMENTO, não a
            # data de auditoria registrada em USUARIO.
            {
                "STATUS": "PCC",
                "USUARIO": "CINTIA - 28/08/2026, 10:00",
                "DATAAGENDAMENTO": "2026-08-30",
            },
        ]

        summary = client.dashboard_aggregates()["resumo_dia_anterior"]

        self.assertEqual(summary["data"], "2026-08-28")
        self.assertEqual(
            summary["labels"],
            [
                "Concluídos",
                "PCC",
                "Cancelados",
                "Agendados",
                "Iniciados não concluídos",
                "Total",
            ],
        )
        self.assertEqual(summary["data_valores"], [1, 1, 1, 1, 1, 5])


if __name__ == "__main__":
    unittest.main()
