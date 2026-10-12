/* ==========================================================================
   REDEB2B — analises.js
   Lógica da página de Análises: filtro de período (ano/mês ou últimos 6
   meses por padrão), busca dos dados agregados e renderização dos gráficos.
   ========================================================================== */

/* INICIO ROTULOS FIXOS
 * Embutido nas duas páginas para funcionar mesmo com o HTML antigo em cache.
 * Mantenha este bloco igual em main.js e analises.js.
 */
(() => {
  if (typeof Chart === "undefined") return;
  const formatter = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 1 });

  const chartValueLabels = {
    id: "valueLabels",
    defaults: {
      color: "#333",
      backgroundColor: "rgba(226, 226, 226, 0.95)",
      fontSize: 11,
      offset: 4,
      showZero: true,
    },
    afterDatasetsDraw(chart, _args, options) {
      if (!chart.chartArea) return;
      const ctx = chart.ctx;
      const fontSize = options.fontSize;
      const labelHeight = fontSize + 6;
      ctx.save();
      ctx.font = `600 ${fontSize}px 'Segoe UI', Arial, sans-serif`;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";

      chart.data.datasets.forEach((_dataset, datasetIndex) => {
        if (!chart.isDatasetVisible(datasetIndex)) return;
        const meta = chart.getDatasetMeta(datasetIndex);
        if (meta.type !== "bar" && meta.type !== "line") return;

        meta.data.forEach((element, index) => {
          if (element.skip || element.hidden || !chart.getDataVisibility(index)) return;
          const parsed = meta.controller.getParsed(index);
          const valueAxis = meta.vScale?.axis || "y";
          const rawValue = parsed?.[valueAxis];
          if (rawValue === null || rawValue === undefined) return;
          const value = Number(rawValue);
          if (!Number.isFinite(value) || (value === 0 && !options.showZero)) return;
          if (!Number.isFinite(element.x) || !Number.isFinite(element.y)) return;

          const isPositive = value >= 0;
          const isHorizontalBar = meta.type === "bar" && valueAxis === "x";
          const label = formatter.format(value);
          const labelWidth = ctx.measureText(label).width + 8;
          let x = element.x - labelWidth / 2;
          let y = isPositive
            ? element.y - options.offset - labelHeight
            : element.y + options.offset;

          if (isHorizontalBar) {
            x = isPositive
              ? element.x + options.offset
              : element.x - options.offset - labelWidth;
            y = element.y - labelHeight / 2;
          }

          // Mantém a caixa inteira dentro do canvas, inclusive nos extremos.
          x = Math.max(2, Math.min(x, chart.width - labelWidth - 2));
          y = Math.max(2, Math.min(y, chart.height - labelHeight - 2));
          ctx.fillStyle = options.backgroundColor;
          ctx.fillRect(x, y, labelWidth, labelHeight);
          ctx.fillStyle = options.color;
          ctx.fillText(label, x + labelWidth / 2, y + labelHeight / 2);
        });
      });

      ctx.restore();
    },
  };

  Chart.register(chartValueLabels);
})();
/* FIM ROTULOS FIXOS */

const analiseState = {
  charts: {},
  analitico: {},
};

document.addEventListener("DOMContentLoaded", () => {
  popularSelectAno();
  bindEventosAnalise();
  // Indicadores e opções são buscados juntos; os cards não ficam esperando
  // a montagem das listas de filtros vindas das várias planilhas.
  carregarAnalise();
  carregarOpcoesDashboard();
});

async function carregarOpcoesDashboard() {
  try {
    const response = await fetch("/api/analytics/options");
    const payload = await response.json();
    if (!response.ok || !payload.ok) throw new Error(payload.error || "Erro ao carregar filtros.");
    const mapping = [
      ["fTecnologiaAnalise", payload.data.tecnologias || []],
      ["fEpsAnalise", payload.data.eps || []],
      ["fCarteiraAnalise", payload.data.carteiras || []],
    ];
    mapping.forEach(([id, values]) => {
      const select = document.getElementById(id);
      select.innerHTML = '<option value="">Todas</option>' + values.map(value => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
    });
    window.rbMultiSelect?.init(document);
  } catch (error) {
    showAlert(error.message || "Não foi possível carregar os filtros executivos.", "warning");
  }
}

function popularSelectAno() {
  const select = document.getElementById("fAno");
  const anoAtual = new Date().getFullYear();
  for (let ano = anoAtual + 1; ano >= anoAtual - 4; ano--) {
    select.appendChild(new Option(ano, ano));
  }
  window.rbMultiSelect?.init(document);
}

function bindEventosAnalise() {
  document.getElementById("btnAplicarPeriodo").addEventListener("click", () => carregarAnalise());
  document.getElementById("btnLimparPeriodo").addEventListener("click", () => {
    window.rbMultiSelect?.clear("fAno");
    window.rbMultiSelect?.clear("fMes");
    document.getElementById("fDataInicioAnalise").value = "";
    document.getElementById("fDataFimAnalise").value = "";
    ["fTecnologiaAnalise", "fEpsAnalise", "fCarteiraAnalise"].forEach(id => window.rbMultiSelect?.clear(id));
    carregarAnalise();
  });

  document.querySelectorAll(".chart-export-icon").forEach((icon) => {
    icon.addEventListener("click", () => {
      const grafico = icon.dataset.grafico;
      const params = coletarFiltrosPeriodo();
      params.append("grafico", grafico);
      window.open(`/api/analytics/export?${params.toString()}`, "_blank");
    });
  });

  document.getElementById("btnExportarFechamento")
    ?.addEventListener("click", baixarExcelFechamentoGeral);
  document.getElementById("fJanelaClientesCriticos")
    ?.addEventListener("change", () => renderClientesCriticos(analiseState.analitico.clientes_criticos || []));
}

async function baixarExcelFechamentoGeral() {
  const button = document.getElementById("btnExportarFechamento");
  const params = coletarFiltrosPeriodo();
  const originalHtml = button.innerHTML;
  button.disabled = true;
  button.innerHTML = '<span class="spinner-border spinner-border-sm" aria-hidden="true"></span> Gerando...';

  try {
    const resp = await fetch(`/api/analytics/fechamento-geral/export?${params.toString()}`);
    if (!resp.ok) {
      let message = `Erro HTTP ${resp.status}`;
      try {
        const payload = await resp.json();
        message = payload.error || message;
      } catch (_err) {
        // A resposta pode não ser JSON se o servidor for interrompido.
      }
      throw new Error(message);
    }

    const blob = await resp.blob();
    const disposition = resp.headers.get("Content-Disposition") || "";
    const filenameMatch = disposition.match(/filename\*?=(?:UTF-8'')?["']?([^"';]+)["']?/i);
    const filename = filenameMatch
      ? decodeURIComponent(filenameMatch[1])
      : "REDEB2B_fechamento_geral.xlsx";
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  } catch (err) {
    showAlert(err.message || "Erro ao baixar o Excel do Fechamento Geral.");
  } finally {
    button.disabled = false;
    button.innerHTML = originalHtml;
  }
}

function coletarFiltrosPeriodo() {
  const anos = window.rbMultiSelect?.values("fAno") || [];
  const meses = window.rbMultiSelect?.values("fMes") || [];
  const dataInicio = document.getElementById("fDataInicioAnalise").value;
  const dataFim = document.getElementById("fDataFimAnalise").value;
  const tecnologias = window.rbMultiSelect?.values("fTecnologiaAnalise") || [];
  const eps = window.rbMultiSelect?.values("fEpsAnalise") || [];
  const carteiras = window.rbMultiSelect?.values("fCarteiraAnalise") || [];

  const params = new URLSearchParams();
  // Um intervalo de datas tem prioridade sobre ano/mês (mesma regra do backend).
  if (dataInicio || dataFim) {
    if (dataInicio) params.append("dataInicio", dataInicio);
    if (dataFim) params.append("dataFim", dataFim);
  } else {
    if (anos.length) params.append("ano", anos.join(","));
    if (meses.length) params.append("mes", meses.join(","));
  }
  if (tecnologias.length) params.append("tecnologia", tecnologias.join(","));
  if (eps.length) params.append("eps", eps.join(","));
  if (carteiras.length) params.append("carteira", carteiras.join(","));
  return params;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str ?? "";
  return div.innerHTML;
}

function showAlert(message, type = "danger") {
  const container = document.getElementById("alertContainer");
  const el = document.createElement("div");
  el.className = `alert alert-${type} alert-dismissible fade show`;
  el.innerHTML = `${escapeHtml(message)}<button type="button" class="btn-close" data-bs-dismiss="alert"></button>`;
  container.appendChild(el);
  setTimeout(() => el.remove(), 5000);
}

async function carregarAnalise() {
  const params = coletarFiltrosPeriodo();
  atualizarResumoPeriodo();

  try {
    const resp = await fetch(`/api/analytics?${params.toString()}`);
    const payload = await resp.json();
    if (!resp.ok || !payload.ok) {
      throw new Error(payload.error || `Erro HTTP ${resp.status}`);
    }
    renderizarAnalise(payload.data);
  } catch (err) {
    showAlert(err.message || "Erro ao carregar as análises.");
  }
}

function atualizarResumoPeriodo() {
  const resumo = document.getElementById("periodoResumo");
  const anos = window.rbMultiSelect?.values("fAno") || [];
  const meses = window.rbMultiSelect?.values("fMes") || [];
  const dataInicio = document.getElementById("fDataInicioAnalise").value;
  const dataFim = document.getElementById("fDataFimAnalise").value;
  const nomesMeses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

  if (dataInicio || dataFim) {
    const formatar = (iso) => {
      if (!iso) return "";
      const [ano_, mes_, dia_] = iso.split("-");
      return `${dia_}/${mes_}/${ano_}`;
    };
    if (dataInicio && dataFim) {
      resumo.textContent = `Período: ${formatar(dataInicio)} até ${formatar(dataFim)}`;
    } else {
      resumo.textContent = `Período: a partir de ${formatar(dataInicio || dataFim)}`;
    }
  } else if (anos.length && meses.length) {
    const mesesTexto = meses.map(item => nomesMeses[parseInt(item, 10) - 1]).join(", ");
    resumo.textContent = `Período: ${mesesTexto} — ${anos.join(", ")}`;
  } else if (anos.length) {
    resumo.textContent = `Período: ${anos.join(", ")} (todos os meses)`;
  } else if (meses.length) {
    resumo.textContent = `Período: ${meses.map(item => nomesMeses[parseInt(item, 10) - 1]).join(", ")}`;
  } else {
    resumo.textContent = "Período: últimos 6 meses (padrão)";
  }
}

function renderizarAnalise(data) {
  const kpis = data.kpis_executivos || {};
  document.getElementById("kpiAtivacoes").textContent = kpis.ativacoes ?? 0;
  document.getElementById("kpiRedeConcluidas").textContent = kpis.rede_concluidas ?? 0;
  document.getElementById("kpiQualidadeConcluidas").textContent = kpis.qualidade_concluidas ?? 0;
  document.getElementById("kpiRfsMes").textContent = kpis.rfs ?? 0;
  document.getElementById("kpiTmrAnalitico").textContent = Number(kpis.tmr_analitico || 0).toLocaleString("pt-BR", {maximumFractionDigits: 2});

  renderizarFechamentoGeral(data.fechamento_geral);
  renderLineChartAnalise("chartConcluidosTimeline", data.concluidos_timeline, "#59a869");
  renderBarChartAnalise("chartDraftsPorMes", data.drafts_por_mes, "#3d95c4");
  renderBarChartAnalise("chartReparoMotivos", normalizarSerie(data.reparo_motivos), "#805ad5");
  renderLineChartAnalise("chartRfsMes", normalizarSerie(data.rfs_mes), "#e69032");
  renderLineChartAnalise("chartAnaliticoBds", normalizarSerie(data.analitico?.bds_por_mes), "#246bfd");
  renderBacklogResumo(data.backlog_resumo || []);
  renderAnalitico(data.analitico || {});
  renderAvisosFontes(data.avisos_fontes || []);
}

function normalizarSerie(source) {
  return { labels: source?.labels || [], data: source?.data || source?.values || [] };
}

function renderAvisosFontes(warnings) {
  const container = document.getElementById("dashboardSourceWarnings");
  container.innerHTML = warnings.length
    ? `<div class="alert alert-warning py-2 mb-0"><strong>Algumas fontes não foram carregadas:</strong> ${warnings.map(item => `${escapeHtml(item.fonte)} — ${escapeHtml(item.mensagem)}`).join(" | ")}</div>`
    : "";
}

function renderBacklogResumo(items) {
  const body = document.getElementById("dashboardBacklogBody");
  body.innerHTML = items.length ? items.map(item => `<tr><td>${escapeHtml(item.nome)}</td><td>${Number(item.quantidade) || 0}</td><td><span class="rb-pill ${item.nivel === "ok" ? "ok" : "critical"}">${item.nivel === "ok" ? "Adequado" : "Atenção"}</span></td></tr>`).join("") : '<tr><td colspan="3" class="text-muted">Sem registros.</td></tr>';
}

function miniTable(headers, rows) {
  if (!rows.length) return '<p class="dashboard-empty">Sem ocorrências no período.</p>';
  return `<div class="rb-table-wrap"><table class="rb-table"><thead><tr>${headers.map(item => `<th>${escapeHtml(item)}</th>`).join("")}</tr></thead><tbody>${rows.map(row => `<tr>${row.map(value => `<td title="${escapeHtml(value)}">${escapeHtml(value)}</td>`).join("")}</tr>`).join("")}</tbody></table></div>`;
}

function renderAnalitico(data) {
  analiseState.analitico = data;
  document.getElementById("analiticoArmarios").innerHTML = miniTable(["Armário", "Quantidade"], (data.armarios_criticos || []).map(item => [item.nome, item.quantidade]));
  document.getElementById("analiticoErbs").innerHTML = miniTable(["ERB", "Quantidade"], (data.erbs_criticas || []).map(item => [item.nome, item.quantidade]));
  document.getElementById("analiticoR30").innerHTML = miniTable(["LP15", "Cliente", "Repetições"], (data.r30_tratativas || []).map(item => [item.lp15, item.cliente || "—", item.quantidade]));
  renderClientesCriticos(data.clientes_criticos || []);
}

function renderClientesCriticos(groups) {
  const dias = Number(document.getElementById("fJanelaClientesCriticos")?.value || 30);
  const group = groups.find(item => Number(item.dias) === dias);
  const rows = (group?.itens || []).map(item => [item.cliente, item.quantidade]);
  document.getElementById("analiticoClientesCriticos").innerHTML = miniTable(["Cliente", "Quantidade"], rows);
}

function formatarDataIso(iso) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || "")) return iso || "—";
  const [ano, mes, dia] = iso.split("-");
  return `${dia}/${mes}/${ano}`;
}

function renderizarFechamentoGeral(fechamento) {
  const periodo = document.getElementById("fechamentoGeralPeriodo");
  const thead = document.getElementById("fechamentoGeralHead");
  const tbody = document.getElementById("fechamentoGeralBody");

  if (!fechamento || !Array.isArray(fechamento.status) || !Array.isArray(fechamento.linhas)) {
    periodo.textContent = "Período indisponível";
    tbody.innerHTML = '<tr><td colspan="2" class="text-center text-muted py-3">Sem dados para exibir.</td></tr>';
    return;
  }

  const inicio = formatarDataIso(fechamento.data_inicio);
  const fim = formatarDataIso(fechamento.data_fim);
  if (fechamento.data_inicio === fechamento.data_fim) {
    periodo.textContent = `${fechamento.usa_dia_anterior ? "Dia útil anterior" : "Data selecionada"}: ${inicio}`;
  } else {
    periodo.textContent = `Período selecionado: ${inicio} a ${fim}`;
  }

  thead.innerHTML = `
    <tr>
      <th scope="col" class="fechamento-atividade">Atividade</th>
      ${fechamento.status.map((status) => `<th scope="col">${escapeHtml(status)}</th>`).join("")}
      <th scope="col" class="fechamento-total">Total</th>
    </tr>`;

  const maiorValor = Math.max(
    0,
    ...fechamento.linhas.flatMap((linha) =>
      (linha.valores || []).map((valor) => Number(valor) || 0)
    )
  );

  const linhasHtml = fechamento.linhas.map((linha) => {
    const valores = (linha.valores || []).map((valor) => Number(valor) || 0);
    return `
      <tr>
        <th scope="row" class="fechamento-atividade">${escapeHtml(linha.atividade || "Não informado")}</th>
        ${valores.map((valor) => `
          <td class="${maiorValor > 0 && valor === maiorValor ? "fechamento-maior" : ""}">${valor}</td>
        `).join("")}
        <td class="fechamento-total">${Number(linha.total) || 0}</td>
      </tr>`;
  }).join("");

  const totais = (fechamento.totais_status || []).map((valor) => Number(valor) || 0);
  tbody.innerHTML = `${linhasHtml}
    <tr class="fechamento-total-row">
      <th scope="row" class="fechamento-atividade">TOTAL</th>
      ${totais.map((valor) => `<td>${valor}</td>`).join("")}
      <td class="fechamento-total">${Number(fechamento.total_geral) || 0}</td>
    </tr>`;
}

function destruirGrafico(canvasId) {
  if (analiseState.charts[canvasId]) {
    analiseState.charts[canvasId].destroy();
    delete analiseState.charts[canvasId];
  }
}

function renderBarChartAnalise(canvasId, dataset, color) {
  destruirGrafico(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  analiseState.charts[canvasId] = new Chart(ctx, {
    type: "bar",
    data: {
      labels: dataset.labels,
      datasets: [{ data: dataset.data, backgroundColor: color, borderRadius: 3 }],
    },
    options: {
      plugins: { legend: { display: false } },
      layout: { padding: { top: 24, left: 8, right: 8 } },
      scales: {
        y: { beginAtZero: true, grace: "10%", ticks: { precision: 0, font: { size: 8 } } },
        x: { ticks: { font: { size: 8 }, maxRotation: 60, minRotation: 45 } },
      },
      responsive: true,
      maintainAspectRatio: false,
    },
  });
}

function renderLineChartAnalise(canvasId, dataset, color) {
  destruirGrafico(canvasId);
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  analiseState.charts[canvasId] = new Chart(ctx, {
    type: "line",
    data: {
      labels: dataset.labels,
      datasets: [{
        data: dataset.data,
        borderColor: color,
        backgroundColor: `${color}26`,
        tension: 0.3,
        fill: true,
        pointRadius: 3,
        pointHoverRadius: 5,
      }],
    },
    options: {
      plugins: { legend: { display: false } },
      layout: { padding: { top: 24, left: 16, right: 16 } },
      scales: {
        y: { beginAtZero: true, grace: "10%", ticks: { precision: 0, font: { size: 8 } } },
        x: { ticks: { font: { size: 8 } } },
      },
      responsive: true,
      maintainAspectRatio: false,
    },
  });
}

const CORES_TECNOLOGIA = ["#3d95c4", "#59a869", "#f0913e", "#6c5ce7", "#d9483a", "#17a2b8"];

function renderGroupedBarChart(canvasId, dataset) {
  destruirGrafico(canvasId);
  const ctx = document.getElementById(canvasId).getContext("2d");
  const datasets = (dataset.datasets || []).map((ds, idx) => ({
    label: ds.label,
    data: ds.data,
    backgroundColor: CORES_TECNOLOGIA[idx % CORES_TECNOLOGIA.length],
    borderRadius: 3,
  }));
  analiseState.charts[canvasId] = new Chart(ctx, {
    type: "bar",
    data: { labels: dataset.labels, datasets },
    options: {
      plugins: {
        legend: { display: true, position: "bottom", labels: { font: { size: 8 }, boxWidth: 10 } },
        valueLabels: { fontSize: 9 },
      },
      layout: { padding: { top: 24, left: 8, right: 8 } },
      scales: {
        y: { beginAtZero: true, grace: "10%", ticks: { precision: 0, font: { size: 8 } } },
        x: { ticks: { font: { size: 8 } } },
      },
      responsive: true,
      maintainAspectRatio: false,
    },
  });
}
