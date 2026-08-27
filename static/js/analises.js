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
};

document.addEventListener("DOMContentLoaded", () => {
  popularSelectAno();
  bindEventosAnalise();
  carregarAnalise();
});

function popularSelectAno() {
  const select = document.getElementById("fAno");
  const anoAtual = new Date().getFullYear();
  for (let ano = anoAtual + 1; ano >= anoAtual - 4; ano--) {
    select.appendChild(new Option(ano, ano));
  }
}

function bindEventosAnalise() {
  document.getElementById("btnAplicarPeriodo").addEventListener("click", () => carregarAnalise());
  document.getElementById("btnLimparPeriodo").addEventListener("click", () => {
    document.getElementById("fAno").value = "";
    document.getElementById("fMes").value = "";
    document.getElementById("fDataInicioAnalise").value = "";
    document.getElementById("fDataFimAnalise").value = "";
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
}

function coletarFiltrosPeriodo() {
  const ano = document.getElementById("fAno").value;
  const mes = document.getElementById("fMes").value;
  const dataInicio = document.getElementById("fDataInicioAnalise").value;
  const dataFim = document.getElementById("fDataFimAnalise").value;

  const params = new URLSearchParams();
  // Um intervalo de datas tem prioridade sobre ano/mês (mesma regra do backend).
  if (dataInicio || dataFim) {
    if (dataInicio) params.append("dataInicio", dataInicio);
    if (dataFim) params.append("dataFim", dataFim);
  } else {
    if (ano) params.append("ano", ano);
    if (mes) params.append("mes", mes);
  }
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
  const ano = document.getElementById("fAno").value;
  const mes = document.getElementById("fMes").value;
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
  } else if (ano && mes) {
    resumo.textContent = `Período: ${nomesMeses[parseInt(mes, 10) - 1]} de ${ano}`;
  } else if (ano) {
    resumo.textContent = `Período: ano de ${ano} (todos os meses)`;
  } else {
    resumo.textContent = "Período: últimos 6 meses (padrão)";
  }
}

function renderizarAnalise(data) {
  document.getElementById("kpiTotalPeriodo").textContent = data.total_registros_periodo ?? 0;
  const clienteTop = data.cliente_top || {};
  document.getElementById("kpiClienteTop").textContent = clienteTop.total
    ? `${clienteTop.cliente} (${clienteTop.total})`
    : "—";

  renderLineChartAnalise("chartConcluidosTimeline", data.concluidos_timeline, "#59a869");
  renderGroupedBarChart("chartMetragemMesTecnologia", data.metragem_por_mes_tecnologia);
  renderGroupedBarChart("chartTipoCaboPorMes", data.por_mes_tipocabo);
  renderBarChartAnalise("chartExecutadoPorAnalise", data.por_executado_por, "#f0913e");
  renderBarChartAnalise("chartStatusAnalise", data.por_status, "#6c5ce7");
  renderBarChartAnalise("chartDraftsPorMes", data.drafts_por_mes, "#3d95c4");
}

function destruirGrafico(canvasId) {
  if (analiseState.charts[canvasId]) {
    analiseState.charts[canvasId].destroy();
    delete analiseState.charts[canvasId];
  }
}

function renderBarChartAnalise(canvasId, dataset, color) {
  destruirGrafico(canvasId);
  const ctx = document.getElementById(canvasId).getContext("2d");
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
  const ctx = document.getElementById(canvasId).getContext("2d");
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
