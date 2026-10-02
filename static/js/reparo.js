const COLS = ["BD","ID_VANTIVE","STATUS","RECLAMACAO","CLIENTE","ENDERECO","CIDADE","UF","CLUSTER","LP_15",
"DATA_ABERTURA","DATA_ENCERRAMENTO","BAIXA_CODIGO","GRUPO_BAIXA","USUARIO_BAIXA","REINC_30D","REINC_TIPO","KPI","TMR",
"TEMPO_PARADA","EPS","TIPO_DEFEITO","PARADA_RELOGIO","INICIO","FIM","RECLAMACAO_CLIENTE","DESCRICAO_FALHA",
"MOTIVO_REAL_BD","ARD_ERB","TECNICO","FILA_ENCERRAMENTO","TECNOLOGIA","TMR_PARADA","TIPO_ARD_ERB","ATUALIZADO_EM"];
const TABLE_COLS = ["BD","CLIENTE","STATUS","CIDADE","DATA_ABERTURA","MOTIVO_REAL_BD","TMR","EPS","TECNOLOGIA"];
const REQUIRED = ["ID_VANTIVE"];
const EPS_OPTIONS = ["VIVO","RS TELECOM","STAFF"];
const TIPO_DEFEITO_OPTIONS = ["CAMPO","SISTEMICO"];
const SIM_NAO_OPTIONS = ["SIM","NÃO"];
const TIPO_ARD_ERB_OPTIONS = ["ERB","ARD"];

let DATA = [];
let DASHBOARD = {};
let FORM_OPTIONS = {};
let view = "dashboard";
let filters = {status:"",uf:"",eps:"",tecnologia:"",busca:"",de:"",ate:""};
let dashboardFilters = {tecnologia:"",de:"",ate:""};
let sortCol = "DATA_ABERTURA", sortDir = "desc";
let page = 1, pageSize = 25;

async function api(url, options={}){
  const response = await fetch(url, {headers:{"Content-Type":"application/json", ...(options.headers||{})}, ...options});
  const body = await response.json().catch(()=>({}));
  if(!response.ok || body.ok===false) throw new Error(body.error || "Não foi possível concluir a operação.");
  return body.data;
}
async function load(){
  const [result,dashboard] = await Promise.all([
    api("/api/reparo/records?page=1&page_size=5000&sort=DATA_ABERTURA:desc"),
    api("/api/reparo/dashboard?"+serverQuery(dashboardFilters)),
  ]);
  DATA = (result.items || []).map(item=>({...item, _id:String(item.row_number)}));
  DASHBOARD = dashboard || {};
}
async function loadOptions(){
  FORM_OPTIONS = await api("/api/reparo/options");
}
async function loadFileStatus(){
  const element=document.getElementById("reparoArquivoStatus");
  try{
    const status=await api("/api/reparo/status");
    element.textContent=status.path;
    element.title=status.path;
  }catch(error){
    element.textContent=error.message;
    element.title=error.message;
  }
}
function serverQuery(source=filters){
  const params=new URLSearchParams();
  Object.entries(source).forEach(([key,value])=>{ if(value) params.set(key,value); });
  return params.toString();
}
async function loadDashboardData(){
  DASHBOARD=await api("/api/reparo/dashboard?"+serverQuery(dashboardFilters));
}
function mk(o){ const r={}; COLS.forEach(c=>r[c]=o[c]??""); r._id = crypto.randomUUID(); return r; }
function esc(value){
  return (value??"").toString().replace(/[&<>"']/g,char=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));
}
function datetimeLocal(value){
  const raw=(value||"").toString().trim().replace(" ","T");
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(raw) ? raw.slice(0,16) : "";
}
function formatDateTime(value){
  const raw=datetimeLocal(value);
  if(!raw) return "—";
  const [datePart,timePart]=raw.split("T");
  const [year,month,day]=datePart.split("-");
  return `${day}/${month}/${year} ${timePart}`;
}
function formatDate(value){
  const raw=(value||"").toString().slice(0,10);
  const parts=raw.split("-");
  return parts.length===3?`${parts[2]}/${parts[1]}/${parts[0]}`:raw;
}
function formOptions(optionKey,current){
  const configured=FORM_OPTIONS[optionKey]||[];
  const values=[...new Set([...(current?[current]:[]),...configured].filter(Boolean))];
  return values;
}

function normStatus(s){
  s=(s||"").toString().trim().toLowerCase();
  if(s.startsWith("enc")||s.startsWith("final")) return "Encerrado";
  if(s.startsWith("and")) return "Em Andamento";
  if(s.startsWith("ab")) return "Aberto";
  return s ? s[0].toUpperCase()+s.slice(1) : "";
}
function statusClass(s){ s=(s||"").toLowerCase(); if(s.includes("enc"))return "encerrado"; if(s.includes("and"))return "andamento"; return "aberto"; }
function fmtNum(n){ const v=parseFloat(n); return isNaN(v)?"—":v.toLocaleString("pt-BR",{maximumFractionDigits:2}); }
function toast(msg){ const t=document.getElementById("toast"); t.textContent=msg; t.classList.add("show"); setTimeout(()=>t.classList.remove("show"),2400); }

function filtered(){
  return DATA.filter(r=>{
    if(filters.status && normStatus(r.STATUS)!==filters.status) return false;
    if(filters.uf && r.UF!==filters.uf) return false;
    if(filters.eps && r.EPS!==filters.eps) return false;
    if(filters.tecnologia && r.TECNOLOGIA!==filters.tecnologia) return false;
    if(filters.de && r.DATA_ABERTURA && r.DATA_ABERTURA < filters.de) return false;
    if(filters.ate && r.DATA_ABERTURA && r.DATA_ABERTURA > filters.ate) return false;
    if(filters.busca){
      const q=filters.busca.toLowerCase();
      const hay=[r.CLIENTE,r.RECLAMACAO_CLIENTE,r.DESCRICAO_FALHA,r.ID_VANTIVE,r.BD].join(" ").toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  }).sort((a,b)=>{
    let x=a[sortCol]||"", y=b[sortCol]||"";
    const nx=parseFloat(x), ny=parseFloat(y);
    if(!isNaN(nx)&&!isNaN(ny)){ x=nx; y=ny; }
    if(x<y) return sortDir==="asc"?-1:1;
    if(x>y) return sortDir==="asc"?1:-1;
    return 0;
  });
}
function uniq(field){ return [...new Set(DATA.map(r=>r[field]).filter(Boolean))].sort(); }

function renderNumberedPagination(currentPage,totalPages,totalRows){
  const from=totalRows?(currentPage-1)*pageSize+1:0;
  const to=Math.min(currentPage*pageSize,totalRows);
  const values=[...new Set([1,totalPages,currentPage-1,currentPage,currentPage+1])]
    .filter(value=>value>=1&&value<=totalPages).sort((a,b)=>a-b);
  const controls=[];
  controls.push(`<button class="page-btn" type="button" data-page="${currentPage-1}" aria-label="Página anterior" title="Página anterior" ${currentPage<=1?"disabled":""}>&laquo;</button>`);
  let previous=0;
  values.forEach(value=>{
    if(previous&&value-previous>1) controls.push('<span class="page-ellipsis" aria-hidden="true">…</span>');
    controls.push(`<button class="page-btn ${value===currentPage?"active":""}" type="button" data-page="${value}" ${value===currentPage?'aria-current="page"':''}>${value}</button>`);
    previous=value;
  });
  controls.push(`<button class="page-btn" type="button" data-page="${currentPage+1}" aria-label="Próxima página" title="Próxima página" ${currentPage>=totalPages?"disabled":""}>&raquo;</button>`);
  return `<div class="pager"><span>${from}–${to} de ${totalRows} registro(s)</span><div class="page-controls">${controls.join("")}</div></div>`;
}

function render(){
  document.querySelectorAll(".nav-btn[data-view]").forEach(b=>b.classList.toggle("active", b.dataset.view===view));
  const m = document.getElementById("main");
  if(view==="dashboard") m.innerHTML = renderDashboard();
  else m.innerHTML = renderListagem();
  wire();
  if(view==="dashboard") drawCharts();
}

function renderDashboard(){
  const kpis=DASHBOARD.kpis||{};
  const total=kpis.total||0;
  const abertos=kpis.abertos||0;
  const encerrados=kpis.encerrados||0;
  const tmrMed=kpis.tmr_medio||0;
  const taxaReinc=kpis.taxa_reincidencia||0;
  const repeated=DASHBOARD.reincidencias_lp_30d||{};
  const repeatedRows=repeated.itens||[];
  const repeatedPeriod=repeated.data_inicio&&repeated.data_fim
    ? `${formatDate(repeated.data_inicio)} a ${formatDate(repeated.data_fim)}`
    : "Últimos 30 dias";
  return `
  <div class="topline"><div><h1>Dashboard</h1><p>Visão geral dos registros de BD</p></div>
    <div class="actions"><button class="btn" id="exportAll"><i class="bi bi-file-earmark-excel"></i> Baixar Excel</button><button class="btn" id="refreshData"><i class="bi bi-arrow-clockwise"></i> Atualizar dados</button></div></div>
  <div class="dashboard-filters">
    <div><label for="dashDe">Data de abertura — de</label><input type="date" id="dashDe" value="${dashboardFilters.de}"></div>
    <div><label for="dashAte">Data de abertura — até</label><input type="date" id="dashAte" value="${dashboardFilters.ate}"></div>
    <div><label for="dashTec">Tecnologia</label><select id="dashTec"><option value="">Todas</option>${uniq("TECNOLOGIA").map(value=>`<option value="${esc(value)}" ${dashboardFilters.tecnologia===value?"selected":""}>${esc(value)}</option>`).join("")}</select></div>
    <button class="btn primary" id="applyDashboardFilters"><i class="bi bi-funnel"></i> Aplicar filtros</button>
    <button class="btn" id="clearDashboardFilters"><i class="bi bi-x-circle"></i> Limpar</button>
  </div>
  <div class="kpis">
    <div class="kpi reparo-kpi-total"><i class="bi bi-collection"></i><div><div class="l">Total de BDs</div><div class="v">${total}</div></div></div>
    <div class="kpi reparo-kpi-open"><i class="bi bi-hourglass-split"></i><div><div class="l">Em aberto / andamento</div><div class="v">${abertos}</div></div></div>
    <div class="kpi reparo-kpi-closed"><i class="bi bi-check-circle"></i><div><div class="l">Encerrados</div><div class="v">${encerrados}</div></div></div>
    <div class="kpi reparo-kpi-tmr"><i class="bi bi-stopwatch"></i><div><div class="l">TMR médio</div><div class="v">${fmtNum(tmrMed)}h</div></div></div>
    <div class="kpi reparo-kpi-repeat"><i class="bi bi-arrow-repeat"></i><div><div class="l">Taxa reincidência 30d</div><div class="v">${fmtNum(taxaReinc)}%</div></div></div>
  </div>
  <div class="charts">
    <div class="panel"><h2>BDs abertos por período</h2><div class="chart-box"><canvas id="chSerie"></canvas></div></div>
    <div class="panel"><h2>Status</h2><div class="chart-box"><canvas id="chStatus"></canvas></div></div>
  </div>
  <div class="charts">
    <div class="panel"><h2>Top motivo real do BD</h2><div class="chart-box"><canvas id="chMotivo"></canvas></div></div>
    <div class="panel"><h2>TMR médio por tecnologia</h2><div class="chart-box"><canvas id="chTec"></canvas></div></div>
  </div>
  <div class="panel repeated-panel">
    <div class="analysis-heading"><div><h2>Reparos repetidos por LP_15 — últimos 30 dias</h2><p>Somente números de cliente com dois ou mais reparos.</p></div><span>${repeatedPeriod}</span></div>
    <div class="tablewrap"><table class="repeated-table">
      <thead><tr><th>LP_15</th><th>Cliente</th><th>Qtd. reparos</th><th>Motivo real do BD</th></tr></thead>
      <tbody>${repeatedRows.length?repeatedRows.map(item=>`<tr>
        <td class="id">${esc(item.lp_15)}</td>
        <td>${esc(item.cliente||"—")}</td>
        <td><span class="repeat-count">${item.quantidade}</span></td>
        <td class="repeat-reasons">${(item.motivos||[]).map(reason=>`${esc(reason.motivo)}${reason.quantidade>1?` (${reason.quantidade})`:""}`).join("<br>")||"—"}</td>
      </tr>`).join(""):`<tr><td colspan="4"><div class="empty"><b>Nenhuma reincidência encontrada</b>Não há LP_15 com dois ou mais reparos no período.</div></td></tr>`}</tbody>
    </table></div>
  </div>`;
}

function drawCharts(){
  Chart.defaults.color = "#64748B";
  Chart.defaults.borderColor = "rgba(148, 163, 184, .28)";
  Chart.defaults.font.family = "IBM Plex Sans";
  Chart.defaults.scale.grid.color = "rgba(148, 163, 184, .24)";
  Chart.defaults.scale.grid.lineWidth = 1;
  Chart.defaults.scale.border.color = "#CBD5E1";
  Chart.defaults.scale.ticks.color = "#64748B";
  const serie=DASHBOARD.serie||{labels:[],values:[]};
  const months=serie.labels||[];
  new Chart(document.getElementById("chSerie"),{type:"line",data:{labels:months,
    datasets:[{data:serie.values||[],borderColor:"#45B8AC",backgroundColor:"rgba(69,184,174,.15)",fill:true,tension:.3}]},
    options:{maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{display:false}},y:{beginAtZero:true}}}});

  const statusData=DASHBOARD.status||{labels:[],values:[]};
  new Chart(document.getElementById("chStatus"),{type:"doughnut",
    data:{labels:statusData.labels||[],datasets:[{data:statusData.values||[],
    backgroundColor:["#E8A33D","#45B8AC","#4C9F6B","#DB5A4C"]}]},
    options:{maintainAspectRatio:false,plugins:{legend:{position:"bottom",labels:{boxWidth:10,padding:12}}}}});

  const motivos=DASHBOARD.motivos||{labels:[],values:[]};
  new Chart(document.getElementById("chMotivo"),{type:"bar",
    data:{labels:motivos.labels||[],datasets:[{data:motivos.values||[],backgroundColor:"#45B8AC"}]},
    options:{maintainAspectRatio:false,indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{beginAtZero:true},y:{grid:{display:false}}}}});

  const technologies=DASHBOARD.tecnologias||{labels:[],values:[]};
  const tecLabels=technologies.labels||[];
  new Chart(document.getElementById("chTec"),{type:"bar",
    data:{labels:tecLabels,datasets:[{data:technologies.values||[],backgroundColor:"#E8A33D"}]},
    options:{maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{x:{grid:{display:false}},y:{beginAtZero:true}}}});
}

function renderListagem(){
  const rows = filtered();
  const pages = Math.max(1, Math.ceil(rows.length/pageSize));
  page = Math.min(page, pages);
  const pageRows = rows.slice((page-1)*pageSize, page*pageSize);
  const opt = (arr)=>arr.map(v=>`<option value="${v}">${v}</option>`).join("");
  return `
  <div class="topline"><div><h1>Listagem</h1><p>${rows.length} registro(s) encontrados</p></div>
    <div class="actions">
      <button class="btn" id="exportFiltered"><i class="bi bi-file-earmark-excel"></i> Baixar Excel</button>
      <button class="btn" id="refreshData"><i class="bi bi-arrow-clockwise"></i> Atualizar dados</button>
      <button class="btn primary" id="addNew">+ Novo registro</button>
    </div></div>
  <div class="filters">
    <input type="text" id="fBusca" placeholder="Buscar cliente, reclamação, ID..." value="${filters.busca}">
    <select id="fStatus"><option value="">Status</option>${opt(["Aberto","Em Andamento","Encerrado"])}</select>
    <select id="fUf"><option value="">UF</option>${opt(uniq("UF"))}</select>
    <select id="fEps"><option value="">EPS</option>${opt(uniq("EPS"))}</select>
    <select id="fTec"><option value="">Tecnologia</option>${opt(uniq("TECNOLOGIA"))}</select>
    <input type="date" id="fDe" title="Data abertura — de" value="${filters.de}">
    <input type="date" id="fAte" title="Data abertura — até" value="${filters.ate}">
  </div>
  <div class="tablewrap"><table class="reparo-list-table">
    <thead><tr>${TABLE_COLS.map(c=>`<th data-col="${c}">${c.replace(/_/g," ")}${sortCol===c?(sortDir==="asc"?" ▲":" ▼"):""}</th>`).join("")}<th>Ações</th></tr></thead>
    <tbody>${pageRows.length? pageRows.map(r=>`
      <tr>${TABLE_COLS.map(c=>{
        if(c==="ID_VANTIVE"||c==="BD") return `<td class="id">${r[c]||"—"}</td>`;
        if(c==="STATUS") return `<td><span class="status ${statusClass(r[c])}">${normStatus(r[c])||"—"}</span></td>`;
        if(c==="TMR") return `<td>${r[c]?fmtNum(r[c])+"h":"—"}</td>`;
        return `<td>${r[c]||"—"}</td>`;
      }).join("")}
      <td class="row-actions"><button class="rowbtn rowbtn-icon" type="button" data-edit="${r._id}" aria-label="Editar registro" title="Editar registro"><i class="bi bi-pencil-square" aria-hidden="true"></i></button><button class="rowbtn rowbtn-icon danger" type="button" data-del="${r._id}" aria-label="Excluir registro" title="Excluir registro"><i class="bi bi-trash3" aria-hidden="true"></i></button></td></tr>`).join("")
      : `<tr><td colspan="${TABLE_COLS.length+1}"><div class="empty"><b>Nenhum registro</b>Ajuste os filtros ou importe uma planilha.</div></td></tr>`}
    </tbody></table></div>
  ${renderNumberedPagination(page,pages,rows.length)}`;
}

function openModal(html){
  document.getElementById("modalBody").innerHTML = html;
  document.getElementById("overlay").classList.add("open");
  bindModalForm();
}
function closeModal(){ document.getElementById("overlay").classList.remove("open"); }

function recordForm(r){
  const editing = !!r;
  r = r || mk({});
  const field=(label,key,type="text",extra="")=>`<div class="field"><label>${label}${REQUIRED.includes(key)?" *":""}</label>
    <input type="${type}" data-f="${key}" value="${esc(type==="datetime-local"?datetimeLocal(r[key]):r[key])}" ${extra}></div>`;
  const select=(label,key,values,placeholder="Selecione")=>`<div class="field"><label>${label}</label><select data-f="${key}">
    <option value="">${placeholder}</option>${values.map(value=>`<option value="${esc(value)}" ${String(r[key]||"").toUpperCase()===String(value).toUpperCase()?"selected":""}>${esc(value)}</option>`).join("")}</select></div>`;
  const motiveOptions=formOptions("motivo_real_bd",r.MOTIVO_REAL_BD);
  const technicianOptions=formOptions("tecnico",r.TECNICO);
  const queueOptions=formOptions("fila_encerramento",r.FILA_ENCERRAMENTO);
  const technologyOptions=formOptions("tecnologia",r.TECNOLOGIA);
  const userOptions=formOptions("usuario",r.USUARIO_BAIXA);
  return `<h3>${editing?"Editar reparo":"Novo reparo"}</h3><div class="sub">${editing? "Registro "+esc(r.ID_VANTIVE):"Preencha os dados do reparo"}</div>
  <div class="grid2">
    ${field("ID","ID_VANTIVE")}
    ${select("EPS","EPS",EPS_OPTIONS)}
    ${select("Tipo de defeito","TIPO_DEFEITO",TIPO_DEFEITO_OPTIONS)}
    ${select("Parada de relógio","PARADA_RELOGIO",SIM_NAO_OPTIONS)}
    ${field("Início (dia e hora)","INICIO","datetime-local")}
    ${field("Fim (dia e hora)","FIM","datetime-local")}
    <div class="field full"><label>Reclamação do cliente</label><textarea data-f="RECLAMACAO_CLIENTE">${esc(r.RECLAMACAO_CLIENTE)}</textarea></div>
    <div class="field full"><label>Descrição da falha</label><textarea data-f="DESCRICAO_FALHA">${esc(r.DESCRICAO_FALHA)}</textarea></div>
    ${select("Motivo real do BD","MOTIVO_REAL_BD",motiveOptions,"Opções a definir")}
    <div class="field"><label>Tipo</label><div class="radio-group">
      ${TIPO_ARD_ERB_OPTIONS.map(value=>`<label><input type="radio" name="tipoArdErb" data-radio-f="TIPO_ARD_ERB" value="${value}" ${String(r.TIPO_ARD_ERB||"").toUpperCase()===value?"checked":""}> ${value}</label>`).join("")}
    </div></div>
    ${field("ARD/ERB","ARD_ERB","text",'maxlength="10" placeholder="Sigla (máximo 10 caracteres)"')}
    ${select("Técnico","TECNICO",technicianOptions,"Opções a definir")}
    ${select("Fila de encerramento","FILA_ENCERRAMENTO",queueOptions,"Opções a definir")}
    ${select("Tecnologia","TECNOLOGIA",technologyOptions,"Opções a definir")}
    ${field("TMR-(PARADA) em horas","TMR_PARADA","text",`readonly data-abertura="${esc(r.DATA_ABERTURA)}" data-encerramento="${esc(r.DATA_ENCERRAMENTO)}"`)}
    ${select("Usuário","USUARIO_BAIXA",userOptions,"Opções a definir")}
    <div class="field audit-info"><label>Última inclusão/alteração</label><div>${formatDateTime(r.ATUALIZADO_EM)}</div></div>
  </div>
  <div class="modal-actions"><button class="btn" id="cancelForm">Cancelar</button><button class="btn primary" id="saveForm" data-id="${r._id}">Salvar</button></div>`;
}

function recordView(r){
  const fields=[
    ["ID","ID_VANTIVE"],["EPS","EPS"],["Tipo de defeito","TIPO_DEFEITO"],
    ["Parada de relógio","PARADA_RELOGIO"],["Início","INICIO"],["Fim","FIM"],
    ["Reclamação do cliente","RECLAMACAO_CLIENTE"],["Descrição da falha","DESCRICAO_FALHA"],
    ["Motivo real do BD","MOTIVO_REAL_BD"],["Tipo","TIPO_ARD_ERB"],["ARD/ERB","ARD_ERB"],
    ["Técnico","TECNICO"],["Fila de encerramento","FILA_ENCERRAMENTO"],
    ["Tecnologia","TECNOLOGIA"],["TMR-(PARADA)","TMR_PARADA"],["Usuário","USUARIO_BAIXA"],
    ["Última inclusão/alteração","ATUALIZADO_EM"],
  ];
  return `<h3>Reparo ${esc(r.ID_VANTIVE)}</h3><div class="sub">Visualização do registro</div>
  <div class="grid2 view-grid">${fields.map(([label,key])=>`<div class="field"><label>${label}</label><div>${esc(r[key]||"—")}</div></div>`).join("")}</div>
  <div class="modal-actions"><button class="btn primary" id="closeView">Fechar</button></div>`;
}

function wire(){
  document.querySelectorAll(".nav-btn[data-view]").forEach(b=>b.onclick=()=>{ view=b.dataset.view; page=1; render(); });
  document.getElementById("overlay").onclick=(e)=>{ if(e.target.id==="overlay") closeModal(); };

  if(view==="dashboard"){
    document.getElementById("exportAll").onclick=()=>{ window.location.href="/api/reparo/export?"+serverQuery(dashboardFilters); };
    document.getElementById("applyDashboardFilters").onclick=async()=>{
      dashboardFilters={
        de:document.getElementById("dashDe").value,
        ate:document.getElementById("dashAte").value,
        tecnologia:document.getElementById("dashTec").value,
      };
      try{await loadDashboardData();render();}catch(error){alert(error.message);}
    };
    document.getElementById("clearDashboardFilters").onclick=async()=>{
      dashboardFilters={tecnologia:"",de:"",ate:""};
      try{await loadDashboardData();render();}catch(error){alert(error.message);}
    };
  }
  if(view==="listagem"){
    document.getElementById("fBusca").oninput=e=>{ filters.busca=e.target.value; page=1; render(); };
    document.getElementById("fStatus").onchange=e=>{ filters.status=e.target.value; page=1; render(); };
    document.getElementById("fUf").onchange=e=>{ filters.uf=e.target.value; page=1; render(); };
    document.getElementById("fEps").onchange=e=>{ filters.eps=e.target.value; page=1; render(); };
    document.getElementById("fTec").onchange=e=>{ filters.tecnologia=e.target.value; page=1; render(); };
    document.getElementById("fDe").onchange=e=>{ filters.de=e.target.value; page=1; render(); };
    document.getElementById("fAte").onchange=e=>{ filters.ate=e.target.value; page=1; render(); };
    document.getElementById("exportFiltered").onclick=()=>{ window.location.href="/api/reparo/export?"+serverQuery(); };
    document.getElementById("addNew").onclick=()=>openModal(recordForm(null));
    document.querySelectorAll("th[data-col]").forEach(th=>th.onclick=()=>{
      const c=th.dataset.col; if(sortCol===c) sortDir = sortDir==="asc"?"desc":"asc"; else { sortCol=c; sortDir="asc"; } render();
    });
    document.querySelectorAll("[data-edit]").forEach(b=>b.onclick=()=>openModal(recordForm(DATA.find(r=>r._id===b.dataset.edit))));
    document.querySelectorAll("[data-del]").forEach(b=>b.onclick=async()=>{
      const record=DATA.find(r=>r._id===b.dataset.del);
      if(record && confirm("Excluir este registro?")){
        try{
          await api("/api/reparo/records/"+encodeURIComponent(record.ID_VANTIVE)+"?row="+encodeURIComponent(record.row_number),{method:"DELETE"});
          await load(); render(); toast("Registro excluído do Excel");
        }catch(error){ alert(error.message); }
      }
    });
    document.querySelectorAll("[data-page]").forEach(button=>button.onclick=()=>{
      page=Number(button.dataset.page); render();
    });
  }
  const refreshButton=document.getElementById("refreshData");
  if(refreshButton) refreshButton.onclick=refreshData;
}

async function refreshData(){
  try{
    await Promise.all([load(),loadOptions(),loadFileStatus()]);
    render();
    toast("Dados atualizados");
  }catch(error){
    document.getElementById("main").innerHTML='<div class="panel"><h2>Não foi possível carregar a base</h2><p>'+esc(error.message)+'</p></div>';
  }
}

function bindModalForm(){
  const cancel=document.getElementById("cancelForm");
  const close=document.getElementById("closeView");
  if(cancel) cancel.onclick=closeModal;
  if(close) close.onclick=closeModal;
  const saveButton=document.getElementById("saveForm");
  if(!saveButton) return;
  const tmrField=document.querySelector('[data-f="TMR_PARADA"]');
  const inicioField=document.querySelector('[data-f="INICIO"]');
  const fimField=document.querySelector('[data-f="FIM"]');
  const updateTmrPreview=()=>{
    if(!tmrField) return;
    const opened=tmrField.dataset.abertura;
    const closed=tmrField.dataset.encerramento;
    if(!opened||!closed){ tmrField.value=""; return; }
    const start=new Date(opened.length===10?`${opened}T00:00`:opened);
    const end=new Date(closed.length===10?`${closed}T00:00`:closed);
    if(Number.isNaN(start.getTime())||Number.isNaN(end.getTime())) return;
    let paused=0;
    if(inicioField?.value&&fimField?.value){
      const pauseStart=new Date(inicioField.value);
      const pauseEnd=new Date(fimField.value);
      if(!Number.isNaN(pauseStart.getTime())&&!Number.isNaN(pauseEnd.getTime())){
        paused=Math.max(0,pauseEnd-pauseStart);
      }
    }
    tmrField.value=(Math.max(0,end-start-paused)/3600000).toFixed(2);
  };
  if(inicioField) inicioField.addEventListener("change",updateTmrPreview);
  if(fimField) fimField.addEventListener("change",updateTmrPreview);
  saveButton.onclick=async()=>{
    const id=saveButton.dataset.id;
    const vals={}; document.querySelectorAll("[data-f]").forEach(el=>vals[el.dataset.f]=el.value);
    const selectedType=document.querySelector('[data-radio-f="TIPO_ARD_ERB"]:checked');
    vals.TIPO_ARD_ERB=selectedType?selectedType.value:"";
    const missing=REQUIRED.filter(key=>!vals[key]);
    if(missing.length){ alert("O campo ID é obrigatório."); return; }
    if((vals.ARD_ERB||"").length>10){ alert("ARD/ERB deve ter no máximo 10 caracteres."); return; }
    if(Boolean(vals.INICIO)!==Boolean(vals.FIM)){ alert("Preencha Início e Fim para calcular a parada do relógio."); return; }
    if(vals.INICIO&&vals.FIM&&new Date(vals.FIM)<new Date(vals.INICIO)){ alert("Fim não pode ser anterior ao Início."); return; }
    const record=DATA.find(item=>item._id===id);
    const url=record
      ? "/api/reparo/records/"+encodeURIComponent(record.ID_VANTIVE)+"?row="+encodeURIComponent(record.row_number)
      : "/api/reparo/records";
    try{
      await api(url,{method:record?"PATCH":"POST",body:JSON.stringify(vals)});
      await load(); closeModal(); render(); toast("Registro salvo no Excel");
    }catch(error){ alert(error.message); }
  };
}

document.addEventListener("DOMContentLoaded",async()=>{
  try{
    await Promise.all([load(),loadOptions(),loadFileStatus()]);
    render();
  }catch(error){
    document.getElementById("main").innerHTML='<div class="panel"><h2>Não foi possível carregar REPARO.xlsx</h2><p>'+error.message+'</p></div>';
  }
});
