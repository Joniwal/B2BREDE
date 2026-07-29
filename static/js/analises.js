
/* ==========================================================================
   REDEB2B — analises.js
   Lógica da página de Análises: filtro de período (ano/mês ou últimos 6
   meses por padrão), busca dos dados agregados e renderização dos gráficos.
   ========================================================================== */

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
    carregarAnalise();
  });
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
  const ano = document.getElementById("fAno").value;
  const mes = document.getElementById("fMes").value;

  const params = new URLSearchParams();
  if (ano) params.append("ano", ano);
  if (mes) params.append("mes", mes);

  atualizarResumoPeriodo(ano, mes);

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

function atualizarResumoPeriodo(ano, mes) {
  const resumo = document.getElementById("periodoResumo");
  const nomesMeses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  if (ano && mes) {
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
  renderMetragemMesTecnologia("chartMetragemMesTecnologia", data.metragem_por_mes_tecnologia);
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
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0, font: { size: 8 } } },
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
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0, font: { size: 8 } } },
        x: { ticks: { font: { size: 8 } } },
      },
      responsive: true,
      maintainAspectRatio: false,
    },
  });
}

const CORES_TECNOLOGIA = ["#3d95c4", "#59a869", "#f0913e", "#6c5ce7", "#d9483a", "#17a2b8"];

function renderMetragemMesTecnologia(canvasId, dataset) {
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
      },
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0, font: { size: 8 } } },
        x: { ticks: { font: { size: 8 } } },
      },
      responsive: true,
      maintainAspectRatio: false,
    },
  });
}
=======
/* ==========================================================================
   REDEB2B — analises.js
   Lógica da página de Análises: filtro de período (ano/mês ou últimos 6
   meses por padrão), busca dos dados agregados e renderização dos gráficos.
   ========================================================================== */

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
    carregarAnalise();
  });
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
  const ano = document.getElementById("fAno").value;
  const mes = document.getElementById("fMes").value;

  const params = new URLSearchParams();
  if (ano) params.append("ano", ano);
  if (mes) params.append("mes", mes);

  atualizarResumoPeriodo(ano, mes);

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

function atualizarResumoPeriodo(ano, mes) {
  const resumo = document.getElementById("periodoResumo");
  const nomesMeses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  if (ano && mes) {
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
  renderMetragemMesTecnologia("chartMetragemMesTecnologia", data.metragem_por_mes_tecnologia);
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
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0, font: { size: 8 } } },
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
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0, font: { size: 8 } } },
        x: { ticks: { font: { size: 8 } } },
      },
      responsive: true,
      maintainAspectRatio: false,
    },
  });
}

const CORES_TECNOLOGIA = ["#3d95c4", "#59a869", "#f0913e", "#6c5ce7", "#d9483a", "#17a2b8"];

function renderMetragemMesTecnologia(canvasId, dataset) {
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
      },
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0, font: { size: 8 } } },
        x: { ticks: { font: { size: 8 } } },
      },
      responsive: true,
      maintainAspectRatio: false,
    },
  });
}
>>>>>>> 402272d (Card Total de Metragem + nova página de Análises com filtro de período)
