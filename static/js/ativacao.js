/* Página ATIVAÇÃO: painel filtrado e CRUD no arquivo ATIVACAO.xlsx. */
(() => {
  "use strict";

  if (typeof Chart !== "undefined") {
    const numberFormat = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 1 });
    Chart.register({
      id: "activationValueLabels",
      defaults: { color: "#334155", backgroundColor: "rgba(235,239,244,.96)", fontSize: 10, offset: 4, showZero: false },
      afterDatasetsDraw(chart, _args, options) {
        if (!chart.chartArea) return;
        const ctx = chart.ctx;
        ctx.save();
        ctx.font = `600 ${options.fontSize}px 'Segoe UI', Arial, sans-serif`;
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
            const value = Number(parsed?.[valueAxis]);
            if (!Number.isFinite(value) || (value === 0 && !options.showZero)) return;
            const label = numberFormat.format(value);
            const height = options.fontSize + 6;
            const width = ctx.measureText(label).width + 8;
            const horizontal = meta.type === "bar" && valueAxis === "x";
            let x = horizontal ? element.x + options.offset : element.x - width / 2;
            let y = horizontal ? element.y - height / 2 : element.y - options.offset - height;
            x = Math.max(2, Math.min(x, chart.width - width - 2));
            y = Math.max(2, Math.min(y, chart.height - height - 2));
            ctx.fillStyle = options.backgroundColor;
            ctx.fillRect(x, y, width, height);
            ctx.fillStyle = options.color;
            ctx.fillText(label, x + width / 2, y + height / 2);
          });
        });
        ctx.restore();
      },
    });
  }

  const state = {
    page: 1,
    pageSize: 20,
    filters: {},
    options: {},
    records: new Map(),
    charts: {},
    newModal: null,
    detailModal: null,
  };

  const byId = (id) => document.getElementById(id);
  const todayIso = () => {
    const now = new Date();
    return new Date(now.getTime() - now.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  };
  const escapeHtml = (value) => String(value ?? "")
    .replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;").replaceAll("'", "&#039;");
  const formatDate = (iso) => {
    if (!iso) return "—";
    const [year, month, day] = String(iso).slice(0, 10).split("-");
    return year && month && day ? `${day}/${month}/${year}` : iso;
  };
  const formatShortDate = (iso) => iso ? `${String(iso).slice(8, 10)}/${String(iso).slice(5, 7)}` : "";
  const shortText = (value, limit = 20) => {
    const text = String(value ?? "");
    return text.length > limit ? `${text.slice(0, limit)}(...)` : text;
  };

  async function api(url, options = {}) {
    const response = await fetch(url, {
      headers: { "Content-Type": "application/json", ...(options.headers || {}) },
      ...options,
    });
    const body = await response.json().catch(() => ({}));
    if (!response.ok || body.ok === false) throw new Error(body.error || "Não foi possível concluir a operação.");
    return body.data;
  }

  function alertUser(message, type = "success") {
    const alert = document.createElement("div");
    alert.className = `alert alert-${type} alert-dismissible fade show shadow-sm`;
    alert.innerHTML = `${escapeHtml(message)}<button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Fechar"></button>`;
    byId("ativacaoAlertContainer").appendChild(alert);
    setTimeout(() => alert.remove(), 5000);
  }

  function fillSelect(id, values, firstLabel = "Selecione...") {
    const select = byId(id);
    const previous = select.value;
    select.innerHTML = `<option value="">${escapeHtml(firstLabel)}</option>` +
      (values || []).map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
    if ((values || []).includes(previous)) select.value = previous;
  }

  function fillDatalist(id, values) {
    byId(id).innerHTML = (values || []).map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
  }

  async function loadOptions() {
    state.options = await api("/api/ativacao/options");
    fillSelect("fAtivacaoServico", state.options.servicos, "Todos");
    fillSelect("fAtivacaoTecnologia", state.options.tecnologias, "Todas");
    fillSelect("fAtivacaoStatus", state.options.status, "Todos");
    fillSelect("fAtivacaoFaturado", state.options.sim_nao, "Todos");
    fillSelect("fAtivacaoRfs", state.options.sim_nao, "Todos");
    fillDatalist("listaClientesAtivacao", state.options.clientes);
    fillDatalist("listaServicosAtivacao", state.options.servicos);
    ["new", "edit"].forEach((prefix) => {
      fillSelect(`${prefix}Cidade`, state.options.cidades);
      fillSelect(`${prefix}Tecnologia`, state.options.tecnologias);
      fillSelect(`${prefix}Empresa`, state.options.empresas);
      fillSelect(`${prefix}Status`, state.options.status);
      fillSelect(`${prefix}Tecnico`, state.options.tecnicos);
      fillSelect(`${prefix}NoMes`, state.options.sim_nao);
      fillSelect(`${prefix}Faturado`, state.options.sim_nao);
      fillSelect(`${prefix}ComRfs`, state.options.sim_nao);
    });
  }

  function collectFilters() {
    const values = {
      q: byId("fAtivacaoBusca").value.trim(),
      dataExecucao: byId("fDataExecucao").value,
      dataAgendamento: byId("fDataAgendamento").value,
      servico: byId("fAtivacaoServico").value,
      tecnologia: byId("fAtivacaoTecnologia").value,
      status: byId("fAtivacaoStatus").value,
      faturado: byId("fAtivacaoFaturado").value,
      comRfs: byId("fAtivacaoRfs").value,
    };
    return Object.fromEntries(Object.entries(values).filter(([, value]) => value));
  }

  function queryString(extra = {}) {
    return new URLSearchParams({ ...state.filters, ...extra }).toString();
  }

  async function loadFileStatus() {
    try {
      const status = await api("/api/ativacao/status");
      byId("ativacaoArquivoStatus").innerHTML = status.exists
        ? `<i class="bi bi-file-earmark-excel text-success"></i> ${escapeHtml(status.path)}`
        : `<i class="bi bi-exclamation-triangle text-warning"></i> O arquivo será criado em ${escapeHtml(status.path)}`;
    } catch (error) {
      byId("ativacaoArquivoStatus").textContent = error.message;
    }
  }

  function chartOptions(horizontal = false, line = false) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "nearest", intersect: false },
      indexAxis: horizontal ? "y" : "x",
      layout: { padding: { top: 22, right: horizontal ? 28 : 5 } },
      plugins: {
        legend: { display: false },
        tooltip: { displayColors: false },
        activationValueLabels: { showZero: false },
      },
      scales: {
        x: { beginAtZero: horizontal, grid: { color: "rgba(148,163,184,.22)" }, ticks: { color: "#64748b", font: { size: 9 }, maxRotation: line ? 45 : 50, minRotation: line ? 0 : 15, precision: 0 } },
        y: { beginAtZero: !horizontal, grid: { color: "rgba(148,163,184,.22)" }, ticks: { color: "#64748b", font: { size: 9 }, precision: 0 } },
      },
    };
  }

  function renderChart(key, canvasId, series, config = {}) {
    if (state.charts[key]) state.charts[key].destroy();
    const line = config.type === "line";
    state.charts[key] = new Chart(byId(canvasId), {
      type: config.type || "bar",
      data: {
        labels: (series.labels || []).map(config.labelFormatter || ((value) => value)),
        datasets: [{
          data: series.values || [],
          borderColor: config.color || "#2563eb",
          backgroundColor: config.background || (line ? "rgba(37,99,235,.14)" : "rgba(37,99,235,.78)"),
          borderWidth: line ? 2 : 1,
          pointRadius: line ? 3 : 0,
          pointHoverRadius: line ? 5 : 0,
          tension: line ? .25 : 0,
          fill: line,
          borderRadius: line ? 0 : 4,
        }],
      },
      options: chartOptions(Boolean(config.horizontal), line),
    });
  }

  async function loadDashboard() {
    const data = await api(`/api/ativacao/dashboard?${queryString()}`);
    byId("kpiAtivacaoTotal").textContent = data.kpis.total;
    byId("kpiAtivacaoOk").textContent = data.kpis.ok;
    byId("kpiAtivacaoNok").textContent = data.kpis.nok;
    byId("kpiAtivacaoFaturado").textContent = data.kpis.faturado;
    byId("kpiAtivacaoRfs").textContent = data.kpis.com_rfs;
    renderChart("status", "chartAtivacaoStatus", data.por_status, { background: "rgba(37,99,235,.8)" });
    renderChart("tecnologia", "chartAtivacaoTecnologia", data.por_tecnologia, { background: "rgba(15,118,110,.8)", color: "#0f766e" });
    renderChart("servico", "chartAtivacaoServico", data.por_servico, { horizontal: true, background: "rgba(249,115,22,.82)", color: "#f97316" });
    renderChart("faturado", "chartAtivacaoFaturado", data.por_faturado, { background: "rgba(22,163,74,.8)", color: "#16a34a" });
    renderChart("rfs", "chartAtivacaoRfs", data.por_rfs, { background: "rgba(124,58,237,.8)", color: "#7c3aed" });
    renderChart("dataExecucao", "chartAtivacaoDataExecucao", data.por_data_execucao, { type: "line", labelFormatter: formatShortDate });
  }

  const badge = (value, positive = "SIM") => `<span class="badge text-bg-${value === positive ? "success" : "secondary"}">${escapeHtml(value || "—")}</span>`;

  function renderTable(data) {
    state.records = new Map((data.items || []).map((item) => [String(item.id), item]));
    const body = byId("ativacaoTableBody");
    if (!data.items?.length) {
      body.innerHTML = '<tr><td colspan="14" class="text-center text-muted py-4">Nenhum registro encontrado para os filtros selecionados.</td></tr>';
    } else {
      body.innerHTML = data.items.map((item) => `
        <tr data-id="${escapeHtml(item.id)}" title="Duplo clique para visualizar">
          <td>${escapeHtml(item.id)}</td><td title="${escapeHtml(item.cliente)}">${escapeHtml(shortText(item.cliente))}</td><td>${escapeHtml(item.cidade)}</td>
          <td>${escapeHtml(item.servico)}</td><td>${escapeHtml(item.tecnologia)}</td><td>${escapeHtml(item.empresa)}</td>
          <td>${badge(item.status, "OK")}</td><td>${formatDate(item.data_agendamento)}</td><td>${formatDate(item.data_execucao)}</td>
          <td>${badge(item.no_mes)}</td><td>${escapeHtml(item.tecnico)}</td><td>${badge(item.faturado)}</td><td>${badge(item.com_rfs)}</td>
          <td class="text-end"><button class="btn btn-outline-primary btn-sm btn-view-ativacao" data-id="${escapeHtml(item.id)}" type="button" title="Visualizar"><i class="bi bi-eye"></i></button></td>
        </tr>`).join("");
    }
    const from = data.total ? (data.page - 1) * data.page_size + 1 : 0;
    const to = Math.min(data.page * data.page_size, data.total);
    byId("ativacaoPaginationInfo").textContent = `${from}–${to} de ${data.total} registro(s)`;
    renderPagination(data.page, data.total_pages);
  }

  function renderPagination(page, totalPages) {
    const values = [...new Set([1, totalPages, page - 1, page, page + 1])]
      .filter((value) => value >= 1 && value <= totalPages).sort((a, b) => a - b);
    const items = [];
    let previous = 0;
    values.forEach((value) => {
      if (previous && value - previous > 1) items.push('<li class="page-item disabled"><span class="page-link">…</span></li>');
      items.push(`<li class="page-item ${value === page ? "active" : ""}"><button class="page-link" data-page="${value}" type="button">${value}</button></li>`);
      previous = value;
    });
    byId("ativacaoPagination").innerHTML = items.join("");
  }

  async function loadRecords() {
    const data = await api(`/api/ativacao/records?${queryString({ page: state.page, page_size: state.pageSize, sort: "data_execucao:desc" })}`);
    renderTable(data);
  }

  async function refreshAll(announce = false) {
    try {
      await Promise.all([loadDashboard(), loadRecords(), loadFileStatus()]);
      if (announce) alertUser("Dados de ativação atualizados.");
    } catch (error) {
      alertUser(error.message, "danger");
    }
  }

  function payload(prefix) {
    return {
      id: byId(`${prefix}Id`).value.trim(),
      cliente: byId(`${prefix}Cliente`).value.trim(),
      cidade: byId(`${prefix}Cidade`).value.trim(),
      servico: byId(`${prefix}Servico`).value.trim(),
      tecnologia: byId(`${prefix}Tecnologia`).value,
      empresa: byId(`${prefix}Empresa`).value,
      status: byId(`${prefix}Status`).value,
      data_agendamento: byId(`${prefix}DataAgendamento`).value,
      data_execucao: byId(`${prefix}DataExecucao`).value,
      no_mes: byId(`${prefix}NoMes`).value,
      tecnico: byId(`${prefix}Tecnico`).value.trim(),
      faturado: byId(`${prefix}Faturado`).value,
      com_rfs: byId(`${prefix}ComRfs`).value,
    };
  }

  function setFormValues(prefix, item) {
    const mapping = {
      Id: "id", Cliente: "cliente", Cidade: "cidade", Servico: "servico", Tecnologia: "tecnologia",
      Empresa: "empresa", Status: "status", DataAgendamento: "data_agendamento",
      DataExecucao: "data_execucao", NoMes: "no_mes", Tecnico: "tecnico",
      Faturado: "faturado", ComRfs: "com_rfs",
    };
    Object.entries(mapping).forEach(([suffix, field]) => { byId(`${prefix}${suffix}`).value = item[field] || ""; });
  }

  function openNewModal() {
    const form = byId("novaAtivacaoForm");
    form.reset();
    form.classList.remove("was-validated");
    byId("novaAtivacaoError").classList.add("d-none");
    byId("newDataAgendamento").value = todayIso();
    byId("newStatus").value = "OK";
    byId("newNoMes").value = "SIM";
    byId("newFaturado").value = "NÃO";
    byId("newComRfs").value = "NÃO";
    state.newModal.show();
  }

  function setDetailEditable(enabled) {
    byId("editarAtivacaoForm").querySelectorAll("input:not([type=hidden]), select").forEach((control) => { control.disabled = !enabled; });
    byId("btnHabilitarEdicao").classList.toggle("d-none", enabled);
    byId("btnSalvarEdicaoAtivacao").classList.toggle("d-none", !enabled);
    byId("btnExcluirAtivacao").classList.toggle("d-none", !enabled);
    byId("detalheAtivacaoModalTitle").textContent = enabled ? "Editar atividade" : "Visualizar atividade";
  }

  function openDetailModal(item) {
    if (!item) return;
    byId("editOriginalId").value = item.id;
    byId("detalheAtivacaoId").textContent = `ID ${item.id}`;
    byId("editarAtivacaoForm").classList.remove("was-validated");
    byId("editarAtivacaoError").classList.add("d-none");
    const technicianSelect = byId("editTecnico");
    if (item.tecnico && !Array.from(technicianSelect.options).some((option) => option.value === item.tecnico)) {
      technicianSelect.add(new Option(item.tecnico, item.tecnico));
    }
    setFormValues("edit", item);
    setDetailEditable(false);
    state.detailModal.show();
  }

  async function saveNew(event) {
    event.preventDefault();
    const form = byId("novaAtivacaoForm");
    form.classList.add("was-validated");
    if (!form.checkValidity()) return;
    const button = byId("btnSalvarNovaAtivacao");
    const errorBox = byId("novaAtivacaoError");
    button.disabled = true;
    errorBox.classList.add("d-none");
    try {
      await api("/api/ativacao/records", { method: "POST", body: JSON.stringify(payload("new")) });
      state.newModal.hide();
      state.page = 1;
      await loadOptions();
      await refreshAll();
      alertUser("Atividade inserida com sucesso.");
    } catch (error) {
      errorBox.textContent = error.message;
      errorBox.classList.remove("d-none");
    } finally {
      button.disabled = false;
    }
  }

  async function saveEdit(event) {
    event.preventDefault();
    const form = byId("editarAtivacaoForm");
    form.classList.add("was-validated");
    if (!form.checkValidity()) return;
    const id = byId("editOriginalId").value;
    const button = byId("btnSalvarEdicaoAtivacao");
    const errorBox = byId("editarAtivacaoError");
    button.disabled = true;
    errorBox.classList.add("d-none");
    try {
      await api(`/api/ativacao/records/${id}`, { method: "PATCH", body: JSON.stringify(payload("edit")) });
      state.detailModal.hide();
      await loadOptions();
      await refreshAll();
      alertUser("Atividade atualizada com sucesso.");
    } catch (error) {
      errorBox.textContent = error.message;
      errorBox.classList.remove("d-none");
    } finally {
      button.disabled = false;
    }
  }

  async function deleteRecord() {
    const id = byId("editOriginalId").value;
    if (!id || !window.confirm(`Excluir definitivamente a atividade ID ${id}?`)) return;
    try {
      await api(`/api/ativacao/records/${id}`, { method: "DELETE" });
      state.detailModal.hide();
      state.page = 1;
      await refreshAll();
      alertUser("Atividade excluída com sucesso.");
    } catch (error) {
      const errorBox = byId("editarAtivacaoError");
      errorBox.textContent = error.message;
      errorBox.classList.remove("d-none");
    }
  }

  function applyFilters() {
    state.filters = collectFilters();
    state.page = 1;
    refreshAll();
  }

  function resetFilterControls() {
    ["fAtivacaoBusca", "fDataExecucao", "fDataAgendamento", "fAtivacaoServico", "fAtivacaoTecnologia", "fAtivacaoStatus", "fAtivacaoFaturado", "fAtivacaoRfs"]
      .forEach((id) => { byId(id).value = ""; });
  }

  function showToday() {
    resetFilterControls();
    byId("fDataExecucao").value = todayIso();
    applyFilters();
  }

  function showAll() {
    resetFilterControls();
    applyFilters();
  }

  function bindEvents() {
    byId("btnNovaAtivacao").addEventListener("click", openNewModal);
    byId("btnAtualizarAtivacao").addEventListener("click", () => refreshAll(true));
    byId("btnAplicarFiltrosAtivacao").addEventListener("click", applyFilters);
    byId("btnHojeAtivacao").addEventListener("click", showToday);
    byId("btnMostrarTodosAtivacao").addEventListener("click", showAll);
    byId("novaAtivacaoForm").addEventListener("submit", saveNew);
    byId("editarAtivacaoForm").addEventListener("submit", saveEdit);
    byId("btnHabilitarEdicao").addEventListener("click", () => setDetailEditable(true));
    byId("btnExcluirAtivacao").addEventListener("click", deleteRecord);
    byId("fDataExecucao").addEventListener("change", () => {
      if (byId("fDataExecucao").value) byId("fDataAgendamento").value = "";
      applyFilters();
    });
    byId("fDataAgendamento").addEventListener("change", () => {
      if (byId("fDataAgendamento").value) byId("fDataExecucao").value = "";
      applyFilters();
    });
    byId("ativacaoPageSize").addEventListener("change", (event) => {
      state.pageSize = Number(event.target.value);
      state.page = 1;
      loadRecords().catch((error) => alertUser(error.message, "danger"));
    });
    byId("ativacaoPagination").addEventListener("click", (event) => {
      const button = event.target.closest("[data-page]");
      if (!button) return;
      state.page = Number(button.dataset.page);
      loadRecords().catch((error) => alertUser(error.message, "danger"));
    });
    byId("ativacaoTableBody").addEventListener("click", (event) => {
      const button = event.target.closest(".btn-view-ativacao");
      if (button) openDetailModal(state.records.get(button.dataset.id));
    });
    byId("ativacaoTableBody").addEventListener("dblclick", (event) => {
      const row = event.target.closest("tr[data-id]");
      if (row) openDetailModal(state.records.get(row.dataset.id));
    });
    byId("btnExportarAtivacao").addEventListener("click", () => {
      window.location.href = `/api/ativacao/export?${queryString()}`;
    });
    let searchTimer;
    byId("fAtivacaoBusca").addEventListener("input", () => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(applyFilters, 400);
    });
  }

  document.addEventListener("DOMContentLoaded", async () => {
    state.newModal = new bootstrap.Modal(byId("novaAtivacaoModal"));
    state.detailModal = new bootstrap.Modal(byId("detalheAtivacaoModal"));
    bindEvents();
    try {
      await loadOptions();
      byId("fDataExecucao").value = todayIso();
      state.filters = collectFilters();
      await refreshAll();
    } catch (error) {
      alertUser(error.message, "danger");
    }
  });
})();
