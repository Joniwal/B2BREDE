const QUALIDADE_STATUS = ["AGENDADO","CABO NA PORTA","CANCELADO","CONCLUIDO","EM CAMPO","INICIADO NAO CONCLUIDO","NOVO","PCC","PENDENTE AGENDAMENTO","SEM ACAO OSP","SEM VT","VERSIONAMENTO","VISTORIA AGENDADA","VISTORIA CONCLUIDA"];
const QUALIDADE_TECNOLOGIAS = ["ERB","GPON","SWT"];
const QUALIDADE_EXECUTORES = ["CINTIA","MARIA CRISTINA","JONI WILSON","MARCOS NEVES","ANA VITORIA","ISABELLA"];
const QUALIDADE_USUARIOS = ["CINTIA","MARIA CRISTINA","ISABELLA","JONI WILSON","MARCOS NEVES","ANA VITORIA"];
const QUALIDADE_FIELDS = ["IDCLIENTE","CLIENTE","ENDERECO","CIDADE","PRODUTO","TECNOLOGIA","VT","DATADISPARO","DATAAGENDAMENTO","DATACONCLUSAO","OBSERVACAO","STATUS","EXECUTADOPOR","OBSERVACAOCONCLUSAO","NUMDRAFT","ROTA","USUARIO"];

const qualidadeState = { page: 1, pageSize: 20, filters: {}, data: null, dashboard: null, chart: null };
const qById = (id) => document.getElementById(id);

function qEscape(value){return String(value??"").replace(/[&<>"']/g,(char)=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));}
function qNormalize(value){return String(value||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toLowerCase();}
function qIsConcluded(value){return qNormalize(value)==="concluido";}
function qFormatDate(value){const raw=String(value||"").slice(0,10);const parts=raw.split("-");return parts.length===3?`${parts[2]}/${parts[1]}/${parts[0]}`:(raw||"—");}

async function qApi(url, options={}){
  const response=await fetch(url,{headers:{"Content-Type":"application/json",...(options.headers||{})},...options});
  const body=await response.json().catch(()=>({}));
  if(!response.ok||body.ok===false) throw new Error(body.error||"Não foi possível concluir a operação.");
  return body.data;
}

function qQuery(extra={}){
  const params=new URLSearchParams();
  Object.entries({...qualidadeState.filters,...extra}).forEach(([key,value])=>{if(value!==""&&value!==undefined&&value!==null)params.set(key,value);});
  return params.toString();
}

function qFillSelect(id, values, placeholder){
  const select=qById(id); select.innerHTML=""; select.appendChild(new Option(placeholder,""));
  values.forEach(value=>select.appendChild(new Option(value,value)));
}

function qSetSelect(id,value){
  const select=qById(id); const wanted=String(value||"");
  if(wanted&&!Array.from(select.options).some(option=>option.value===wanted)) select.appendChild(new Option(wanted,wanted));
  select.value=wanted;
}

function qCollectFilters(){
  qualidadeState.filters={
    q:qById("qBusca").value.trim(),status:qById("qStatus").value,tecnologia:qById("qTecnologia").value,
    dataInicio:qById("qDataInicio").value,dataFim:qById("qDataFim").value,
  };
}

async function qLoad(){
  const query=qQuery({page:qualidadeState.page,page_size:qualidadeState.pageSize,sort:"DATAAGENDAMENTO:desc"});
  const [data,dashboard]=await Promise.all([
    qApi(`/api/qualidade/records?${query}`),qApi(`/api/qualidade/dashboard?${qQuery()}`),
  ]);
  qualidadeState.data=data; qualidadeState.dashboard=dashboard;
  qRender();
}

function qShowAlert(message,type="success"){
  const host=qById("qualidadeAlertContainer");
  host.innerHTML=`<div class="alert alert-${type} shadow-sm" role="alert">${qEscape(message)}</div>`;
  setTimeout(()=>{host.innerHTML="";},3500);
}

function qShowModal(){
  const modal=qById("qualidadeModal");
  if(window.bootstrap?.Modal){bootstrap.Modal.getOrCreateInstance(modal).show();return;}
  modal.style.display="block";modal.classList.add("show");modal.removeAttribute("aria-hidden");modal.setAttribute("aria-modal","true");modal.setAttribute("role","dialog");
  document.body.classList.add("modal-open");
  if(!document.querySelector(".qualidade-modal-backdrop")){
    const backdrop=document.createElement("div");backdrop.className="modal-backdrop fade show qualidade-modal-backdrop";document.body.appendChild(backdrop);
  }
}

function qHideModal(){
  const modal=qById("qualidadeModal");
  if(window.bootstrap?.Modal){bootstrap.Modal.getOrCreateInstance(modal).hide();return;}
  modal.style.display="none";modal.classList.remove("show");modal.setAttribute("aria-hidden","true");modal.removeAttribute("aria-modal");modal.removeAttribute("role");
  document.body.classList.remove("modal-open");document.querySelector(".qualidade-modal-backdrop")?.remove();
}

function qRenderDashboard(){
  const data=qualidadeState.dashboard||{};
  qById("qKpiTotal").textContent=data.total||0;
  qById("qKpiConcluidas").textContent=data.concluidas||0;
  qById("qKpiNaoConcluidas").textContent=data.nao_concluidas||0;
  if(qualidadeState.chart) qualidadeState.chart.destroy();
  qualidadeState.chart=new Chart(qById("qChartConclusao"),{
    type:"doughnut",data:{labels:data.grafico?.labels||[],datasets:[{data:data.grafico?.values||[],backgroundColor:["#15945c","#e07b22"],borderWidth:0}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:"64%",plugins:{legend:{display:false},valueLabels:{showZero:true,fontSize:10}}},
  });
}

function qPaginationItems(page,pages){
  const values=[...new Set([1,pages,page-1,page,page+1])].filter(value=>value>=1&&value<=pages).sort((a,b)=>a-b);
  const items=[]; let previous=0;
  items.push(`<li class="page-item ${page<=1?"disabled":""}"><button class="page-link" data-page="${page-1}" type="button">«</button></li>`);
  values.forEach(value=>{if(previous&&value-previous>1)items.push('<li class="page-item disabled"><span class="page-link">…</span></li>');items.push(`<li class="page-item ${value===page?"active":""}"><button class="page-link" data-page="${value}" type="button">${value}</button></li>`);previous=value;});
  items.push(`<li class="page-item ${page>=pages?"disabled":""}"><button class="page-link" data-page="${page+1}" type="button">»</button></li>`);
  return items.join("");
}

function qRenderTable(){
  const data=qualidadeState.data||{items:[],total:0,page:1,page_size:qualidadeState.pageSize};
  const body=qById("qTableBody");
  body.innerHTML=data.items.length?data.items.map(item=>`<tr>
    <td>${qEscape(item.IDCLIENTE||"—")}</td><td title="${qEscape(item.CLIENTE)}">${qEscape(item.CLIENTE||"—")}</td><td>${qEscape(item.CIDADE||"—")}</td><td>${qEscape(item.TECNOLOGIA||"—")}</td>
    <td><span class="q-status ${qIsConcluded(item.STATUS)?"concluido":""}">${qEscape(item.STATUS||"—")}</span></td><td>${qFormatDate(item.DATAAGENDAMENTO)}</td><td>${qFormatDate(item.DATACONCLUSAO)}</td><td>${qEscape(item.EXECUTADOPOR||"—")}</td>
    <td class="text-end q-actions"><button class="btn btn-outline-primary btn-sm q-icon-btn" data-edit="${qEscape(item.IDCLIENTE)}" title="Editar" aria-label="Editar"><i class="bi bi-pencil-square"></i></button> <button class="btn btn-outline-danger btn-sm q-icon-btn" data-delete="${qEscape(item.IDCLIENTE)}" title="Excluir" aria-label="Excluir"><i class="bi bi-trash3"></i></button></td></tr>`).join(""):'<tr><td colspan="9" class="text-center text-muted py-4">Nenhuma ação de qualidade encontrada.</td></tr>';
  const pages=Math.max(1,Math.ceil(data.total/data.page_size)); const from=data.total?(data.page-1)*data.page_size+1:0; const to=Math.min(data.page*data.page_size,data.total);
  qById("qPaginationInfo").textContent=`${from}–${to} de ${data.total} registro(s)`;
  qById("qPagination").innerHTML=qPaginationItems(data.page,pages);
  body.querySelectorAll("[data-edit]").forEach(button=>{
    button.addEventListener("click",()=>qOpenEdit(button.dataset.edit));
  });
  body.querySelectorAll("[data-delete]").forEach(button=>{
    button.addEventListener("click",()=>qDelete(button.dataset.delete));
  });
}

function qRender(){qRenderDashboard();qRenderTable();}

function qClearForm(){
  QUALIDADE_FIELDS.forEach(field=>{const element=qById(`qf_${field}`);if(element)element.value="";});
  qById("qOriginalId").value="";qById("qualidadeModalAlert").innerHTML="";qById("btnExcluirQualidade").classList.add("d-none");qById("qf_IDCLIENTE").disabled=false;
}

function qOpenCreate(){
  qClearForm();qById("qualidadeModalTitle").textContent="Nova ação de qualidade";
  qShowModal();setTimeout(()=>qById("qf_IDCLIENTE").focus(),250);
}

async function qOpenEdit(id){
  qClearForm();
  try{
    const item=await qApi(`/api/qualidade/records/${encodeURIComponent(id)}`);
    QUALIDADE_FIELDS.forEach(field=>{if(field==="USUARIO")return;const element=qById(`qf_${field}`);if(!element)return;if(["TECNOLOGIA","STATUS","EXECUTADOPOR"].includes(field))qSetSelect(`qf_${field}`,item[field]);else element.value=item[field]||"";});
    qById("qOriginalId").value=id;qById("qf_IDCLIENTE").disabled=true;qById("qualidadeModalTitle").textContent=`Editar ação ${id}`;qById("btnExcluirQualidade").classList.remove("d-none");
    qShowModal();
  }catch(error){qShowAlert(error.message,"danger");}
}

async function qSave(){
  const form=qById("qualidadeForm");if(!form.checkValidity()){form.reportValidity();return;}
  const originalId=qById("qOriginalId").value;const payload={};
  QUALIDADE_FIELDS.forEach(field=>{const element=qById(`qf_${field}`);if(element)payload[field]=element.value;});
  const now=new Date().toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",year:"numeric",hour:"2-digit",minute:"2-digit"});
  payload.USUARIO=`${qById("qf_USUARIO").value} - ${now}`;
  try{
    await qApi(originalId?`/api/qualidade/records/${encodeURIComponent(originalId)}`:"/api/qualidade/records",{method:originalId?"PATCH":"POST",body:JSON.stringify(payload)});
    qHideModal();await qLoad();qShowAlert(originalId?"Ação atualizada com sucesso.":"Ação criada com sucesso.");
  }catch(error){qById("qualidadeModalAlert").innerHTML=`<div class="alert alert-danger">${qEscape(error.message)}</div>`;}
}

async function qDelete(id){
  if(!id||!confirm(`Excluir a ação de qualidade ${id}?`))return;
  try{await qApi(`/api/qualidade/records/${encodeURIComponent(id)}`,{method:"DELETE"});qHideModal();await qLoad();qShowAlert("Ação excluída com sucesso.");}catch(error){qShowAlert(error.message,"danger");}
}

function qBind(){
  qById("btnAplicarQualidade").onclick=()=>{qCollectFilters();qualidadeState.page=1;qLoad().catch(error=>qShowAlert(error.message,"danger"));};
  qById("btnLimparQualidade").onclick=()=>{["qBusca","qStatus","qTecnologia","qDataInicio","qDataFim"].forEach(id=>qById(id).value="");qualidadeState.filters={};qualidadeState.page=1;qLoad().catch(error=>qShowAlert(error.message,"danger"));};
  qById("btnAtualizarQualidade").onclick=()=>qLoad().then(()=>qShowAlert("Dados atualizados.")).catch(error=>qShowAlert(error.message,"danger"));
  qById("btnExportarQualidade").onclick=()=>{window.location.href=`/api/qualidade/export?${qQuery()}`;};qById("btnNovaQualidade").addEventListener("click",qOpenCreate);qById("btnSalvarQualidade").onclick=qSave;
  qById("btnExcluirQualidade").onclick=()=>qDelete(qById("qOriginalId").value);
  qById("qualidadeModal").querySelectorAll('[data-bs-dismiss="modal"]').forEach(button=>button.addEventListener("click",event=>{event.preventDefault();qHideModal();}));
  qById("qPageSize").onchange=(event)=>{qualidadeState.pageSize=Number(event.target.value);qualidadeState.page=1;qLoad().catch(error=>qShowAlert(error.message,"danger"));};
  qById("qPagination").onclick=(event)=>{const button=event.target.closest("[data-page]");if(!button||button.closest(".disabled")||button.closest(".active"))return;qualidadeState.page=Number(button.dataset.page);qLoad().catch(error=>qShowAlert(error.message,"danger"));};
  qById("qBusca").addEventListener("keydown",event=>{if(event.key==="Enter"){qCollectFilters();qualidadeState.page=1;qLoad().catch(error=>qShowAlert(error.message,"danger"));}});
}

document.addEventListener("DOMContentLoaded",async()=>{
  qFillSelect("qStatus",QUALIDADE_STATUS,"Todos");qFillSelect("qTecnologia",QUALIDADE_TECNOLOGIAS,"Todas");qFillSelect("qf_STATUS",QUALIDADE_STATUS,"Selecione...");qFillSelect("qf_TECNOLOGIA",QUALIDADE_TECNOLOGIAS,"Selecione...");qFillSelect("qf_EXECUTADOPOR",QUALIDADE_EXECUTORES,"Selecione...");qFillSelect("qf_USUARIO",QUALIDADE_USUARIOS,"Selecione...");qBind();
  try{const status=await qApi("/api/excel-status");qById("qualidadeArquivoStatus").textContent=status.caminho||status.mensagem||"Base localizada.";}catch(error){qById("qualidadeArquivoStatus").textContent=error.message;}
  try{await qLoad();const cities=[...new Set((qualidadeState.data?.items||[]).map(item=>item.CIDADE).filter(Boolean))].sort();qById("qCidades").innerHTML=cities.map(city=>`<option value="${qEscape(city)}"></option>`).join("");}catch(error){qById("qTableBody").innerHTML=`<tr><td colspan="9" class="text-center text-danger py-4">${qEscape(error.message)}</td></tr>`;}
});
