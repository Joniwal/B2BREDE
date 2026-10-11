# -*- coding: utf-8 -*-
"""Composição dos indicadores executivos do Dashboard B2B."""

from __future__ import annotations

import unicodedata
from collections import Counter
from calendar import monthrange
from datetime import date, datetime

from excel_client import DataClientError, _parse_date
from backlog_client import _date as backlog_date, _text as backlog_text


def _norm(value):
    text = str(value or "").strip().casefold()
    return "".join(char for char in unicodedata.normalize("NFKD", text) if not unicodedata.combining(char))


def _iso_date(value):
    parsed = _parse_date(value)
    try:
        return date.fromisoformat(parsed) if parsed else None
    except ValueError:
        return None


def resolve_period(ano=None, mes=None, data_inicio=None, data_fim=None):
    if data_inicio or data_fim:
        start = _iso_date(data_inicio or data_fim)
        end = _iso_date(data_fim or data_inicio)
        if not start or not end:
            raise DataClientError("Período inválido.", status_code=400)
        return (end, start) if start > end else (start, end)
    today = date.today()
    year = int(ano) if ano else today.year
    if mes:
        month = int(mes)
        return date(year, month, 1), date(year, month, monthrange(year, month)[1])
    if ano:
        return date(year, 1, 1), date(year, 12, 31)
    return date(today.year, today.month, 1), date(today.year, today.month, monthrange(today.year, today.month)[1])


class DashboardService:
    def __init__(self, data_client, ativacao_client, reparo_client, backlog_client, analitico_client):
        self.data = data_client
        self.ativacao = ativacao_client
        self.reparo = reparo_client
        self.backlog = backlog_client
        self.analitico = analitico_client

    @staticmethod
    def _safe(label, callback, default):
        try:
            return callback(), None
        except DataClientError as exc:
            return default, {"fonte": label, "mensagem": exc.message}

    def options(self):
        technologies, eps, wallets, warnings = set(), set(), set(), []
        loaders = [
            ("Rede", lambda: sorted({str(row.get("TECNOLOGIA") or "").strip() for row in self.data._excel_read_all() if str(row.get("TECNOLOGIA") or "").strip()})),
            ("Ativação", lambda: self.ativacao.options().get("tecnologias", [])),
            ("Reparo", lambda: self.reparo.options().get("tecnologia", [])),
            ("Analítico", lambda: self.analitico.options().get("tecnologias", [])),
        ]
        for label, loader in loaders:
            values, warning = self._safe(label, loader, [])
            technologies.update(values)
            if warning:
                warnings.append(warning)
        repair_options, warning = self._safe("Reparo", self.reparo.options, {})
        eps.update(repair_options.get("eps", []))
        if warning:
            warnings.append(warning)
        backlog_options, warning = self._safe("Backlog", self.backlog.options, {})
        wallets.update(backlog_options.get("carteiras", []))
        technologies.update(backlog_options.get("tecnologias", []))
        if warning:
            warnings.append(warning)
        return {
            "tecnologias": sorted(technologies, key=str.casefold),
            "eps": sorted(eps, key=str.casefold),
            "carteiras": sorted(wallets, key=str.casefold),
            "avisos": warnings,
        }

    def build(self, *, base, ano=None, mes=None, data_inicio=None, data_fim=None,
              tecnologias=None, eps=None, carteiras=None):
        start, end = resolve_period(ano, mes, data_inicio, data_fim)
        technologies = tecnologias or []
        selected_tech = {_norm(value) for value in technologies}
        warnings = []

        network_rows, warning = self._safe("Rede", self.data._excel_read_all, [])
        if warning:
            warnings.append(warning)
        if selected_tech:
            network_rows = [row for row in network_rows if _norm(row.get("TECNOLOGIA")) in selected_tech]

        def in_period(row, field):
            value = _iso_date(row.get(field))
            return bool(value and start <= value <= end)

        completed = [row for row in network_rows if _norm(row.get("STATUS")) == "concluido" and in_period(row, "DATACONCLUSAO")]
        quality_completed = [row for row in completed if _norm(row.get("ATIVIDADE")) == "acao de qualidade"]
        rede_completed = [row for row in completed if _norm(row.get("ATIVIDADE")) != "acao de qualidade"]

        activation_rows, warning = self._safe("Ativação", self.ativacao._all, [])
        if warning:
            warnings.append(warning)
        if selected_tech:
            activation_rows = [row for row in activation_rows if _norm(row.get("tecnologia")) in selected_tech]
        activation_total = sum(
            1 for row in activation_rows
            if (value := _iso_date(row.get("data_execucao"))) and start <= value <= end
        )

        repair_filters = {"de": start.isoformat(), "ate": end.isoformat()}
        if technologies:
            repair_filters["tecnologia"] = technologies
        if eps:
            repair_filters["eps"] = eps
        repair, warning = self._safe("Reparo", lambda: self.reparo.dashboard(repair_filters), {"motivos": {"labels": [], "values": []}})
        if warning:
            warnings.append(warning)

        backlog_filters = {}
        if technologies:
            backlog_filters["tecnologia"] = technologies
        if carteiras:
            backlog_filters["carteira"] = carteiras
        backlog, warning = self._safe("Backlog", lambda: self.backlog.dashboard(backlog_filters), {"carteiras": []})
        if warning:
            warnings.append(warning)

        # O resumo por carteira representa o retrato atual do arquivo; apenas o RFS
        # é recortado pelo mês/período selecionado.
        def monthly_rfs():
            _headers, rows = self.backlog._read()
            if carteiras:
                rows = self.backlog._filter(rows, {"carteira": carteiras})
            if technologies:
                rows = self.backlog._filter(rows, {"tecnologia": technologies})
            counts = Counter()
            for row in rows:
                value = backlog_date(self.backlog._value(row, "Data_RFS"))
                if value and start <= value <= end:
                    counts[value] += 1
            labels = sorted(counts)
            return {"labels": [item.isoformat() for item in labels], "values": [counts[item] for item in labels], "total": sum(counts.values())}

        rfs_month, warning = self._safe("Backlog", monthly_rfs, {"labels": [], "values": [], "total": 0})
        if warning:
            warnings.append(warning)

        analytical, warning = self._safe(
            "Analítico",
            lambda: self.analitico.dashboard(start=start, end=end, technologies=technologies),
            {"bds_por_mes": {"labels": [], "values": []}, "armarios_criticos": [], "erbs_criticas": [], "r30_tratativas": [], "clientes_criticos": [], "tmr_medio_horas": 0, "total_periodo": 0},
        )
        if warning:
            warnings.append(warning)

        base.update({
            "periodo_executivo": {"inicio": start.isoformat(), "fim": end.isoformat()},
            "kpis_executivos": {
                "ativacoes": activation_total,
                "rede_concluidas": len(rede_completed),
                "qualidade_concluidas": len(quality_completed),
                "rfs": rfs_month["total"],
                "tmr_analitico": analytical.get("tmr_medio_horas", 0),
            },
            "reparo_motivos": repair.get("motivos", {"labels": [], "values": []}),
            "backlog_resumo": backlog.get("carteiras", []),
            "rfs_mes": rfs_month,
            "analitico": analytical,
            "avisos_fontes": warnings,
        })
        return base
