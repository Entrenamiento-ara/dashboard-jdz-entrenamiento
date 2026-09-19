const excelWeb='https://jmartins.sharepoint.com/:x:/r/sites/EntrenamientoR3/_layouts/15/Doc.aspx?sourcedoc=%7BB762772B-6048-4002-AF66-F0BC9A0A1055%7D&file=Ingresos%20con%20tiendas%20de%20entrenamiento%20y%20pares%20formadores.xlsm&action=default&mobileredirect=true';
function abrirExcel(){location.href='ms-excel:ofe|u|'+excelWeb}

let arandoRecords=[];
let interviewRecords=[];
let agendaRecords=[];
let calendarCursor=null;
const loadedFlags={agenda:false,arando:false,interview:false};

function norm(v){return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toLowerCase()}
function safe(v){return String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}
function statusClass(v){const x=norm(v);if(x.includes('no aprobado')||x.includes('no aprueba')||x.includes('cancel')||x.includes('no apto'))return'bad';if(x.includes('aprobado')||x.includes('aprueba')||x.includes('realizado')||x.includes('complet'))return'ok';if(x.includes('pend')||x.includes('proceso')||x.includes('evalu'))return'warn';return'info'}
function dateDisplay(iso){if(!iso)return'';const p=iso.split('-');return p.length===3?`${p[2]}/${p[1]}/${p[0]}`:iso}
function isoDate(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function addDays(d,n){const x=new Date(d);x.setDate(x.getDate()+n);return x}
function startOfWeek(d){const x=new Date(d);const day=x.getDay();const diff=day===0?-6:1-day;x.setDate(x.getDate()+diff);x.setHours(12,0,0,0);return x}

/* ===================== PERSISTENCIA LOCAL =====================
   Los datos procesados quedan guardados SOLO en este navegador.
   Refrescar la página no los elimina. Cuando se selecciona un nuevo
   archivo para una fuente, esa fuente reemplaza a la anterior.
   ============================================================= */
const CACHE_DB='PortalJDZ_Data_v2';
const CACHE_STORE='datasets';
function openCache(){return new Promise((resolve,reject)=>{const req=indexedDB.open(CACHE_DB,1);req.onupgradeneeded=()=>{const db=req.result;if(!db.objectStoreNames.contains(CACHE_STORE))db.createObjectStore(CACHE_STORE,{keyPath:'key'})};req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)})}
async function cachePut(key,records,fileName){const db=await openCache();return new Promise((resolve,reject)=>{const tx=db.transaction(CACHE_STORE,'readwrite');tx.objectStore(CACHE_STORE).put({key,records,fileName:fileName||'',savedAt:new Date().toISOString()});tx.oncomplete=()=>{db.close();resolve()};tx.onerror=()=>{db.close();reject(tx.error)}})}
async function cacheGet(key){const db=await openCache();return new Promise((resolve,reject)=>{const tx=db.transaction(CACHE_STORE,'readonly');const req=tx.objectStore(CACHE_STORE).get(key);req.onsuccess=()=>{db.close();resolve(req.result||null)};req.onerror=()=>{db.close();reject(req.error)}})}
function showSourceState(id,ok,text){const el=document.getElementById(id);if(!el)return;el.textContent=text;el.className='source-state '+(ok?'source-ok':'source-bad')}
function markLoaded(key){loadedFlags[key]=true;const el=document.getElementById('kpiSources');if(el)el.textContent=Object.values(loadedFlags).filter(Boolean).length}
function cacheText(n,fileName,unit='registros'){return `${n} ${unit} guardados en este navegador${fileName?' · '+fileName:''}`}

function excelDate(value){
  if(value===null||value===undefined||value==='')return'';
  if(value instanceof Date&&!isNaN(value))return isoDate(value);
  if(typeof value==='number'){const d=XLSX.SSF.parse_date_code(value);if(!d)return'';return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`}
  const s=String(value).trim();if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
  let m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);if(m){let y=Number(m[3]);if(y<100)y+=2000;return `${y}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`}
  const d=new Date(s);return isNaN(d)?'':isoDate(d);
}
function excelTime(value){
  if(value===null||value===undefined||value==='')return'';
  if(value instanceof Date&&!isNaN(value))return `${String(value.getHours()).padStart(2,'0')}:${String(value.getMinutes()).padStart(2,'0')}`;
  if(typeof value==='number'){const total=Math.round((value%1)*24*60);return `${String(Math.floor(total/60)%24).padStart(2,'0')}:${String(total%60).padStart(2,'0')}`}
  const s=String(value).trim();const m=s.match(/(\d{1,2}):(\d{2})/);return m?`${String(m[1]).padStart(2,'0')}:${m[2]}`:s;
}
function readWorkbook(file){return new Promise((resolve,reject)=>{const fr=new FileReader();fr.onload=e=>{try{resolve(XLSX.read(e.target.result,{type:'array',cellDates:false,cellText:true}))}catch(err){reject(err)}};fr.onerror=()=>reject(fr.error||new Error('No se pudo abrir el archivo'));fr.readAsArrayBuffer(file)})}
function headerIndex(headers,name){const n=norm(name);return headers.findIndex(h=>norm(h)===n)}
function firstIndex(headers,names){for(const n of names){const i=headerIndex(headers,n);if(i>=0)return i}return-1}

/* ===================== ARANDO TALENTO ===================== */
async function loadArandoFile(file){
  try{
    const wb=await readWorkbook(file);const result=[];
    for(const sheetName of wb.SheetNames){
      const sn=norm(sheetName);if(!sn.includes('operador')&&!sn.includes('supervisor'))continue;
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,raw:true,defval:''});if(!rows.length)continue;
      const h=rows[0];const iSap=firstIndex(h,['SAP']);const iName=firstIndex(h,['NOMBRE']);const iCargo=firstIndex(h,['CARGO','Cargo']);const iGrupo=firstIndex(h,['GRUPO']);const iConcept=firstIndex(h,['Concept Final','CONCEPTO']);const iNew=firstIndex(h,['Nuevo cargo']);const iStore=firstIndex(h,['NOMBRE TIENDA','TIENDA']);
      if(iSap<0)continue;
      for(let r=1;r<rows.length;r++){
        const row=rows[r];const sap=String(row[iSap]??'').trim();if(!sap)continue;
        result.push({sap,nombre:iName>=0?String(row[iName]??'').trim():'',cargo:iCargo>=0?String(row[iCargo]??'').trim():'',grupo:iGrupo>=0?String(row[iGrupo]??'').trim():'',concepto:iConcept>=0?String(row[iConcept]??'').trim():'',nuevo:iNew>=0?String(row[iNew]??'').trim():'',tienda:iStore>=0?String(row[iStore]??'').trim():'',tipo:sheetName.trim(),anio:'2026'});
      }
    }
    if(!result.length)throw new Error('No se encontraron registros de Operadores o Supervisores');
    arandoRecords=result;await cachePut('arando',result,file.name);applyArandoData(result,file.name);
  }catch(e){console.error('Arando:',e);showSourceState('arandoState',false,`No se pudo leer: ${e.message||'error desconocido'}`)}
}
function applyArandoData(records,fileName){arandoRecords=records||[];showSourceState('arandoState',true,cacheText(arandoRecords.length,fileName));const k=document.getElementById('kpiArando');if(k)k.textContent=arandoRecords.length;markLoaded('arando');const out=document.getElementById('sapResult');if(out)out.innerHTML='<div class="empty-state">Datos listos. Escribe un SAP para consultar el historial.</div>'}
function buscarSAP(){
  const sap=document.getElementById('sap').value.trim();const out=document.getElementById('sapResult');
  if(!arandoRecords.length){out.innerHTML='<div class="empty-state">Primero carga el archivo de Arando Talento.</div>';return}
  if(!sap){out.innerHTML='<div class="empty-state">Escribe un SAP para consultar.</div>';return}
  const rows=arandoRecords.filter(r=>r.sap===sap);
  if(!rows.length){out.innerHTML=`<div class="empty-state">No se encontraron registros para el SAP <b>${safe(sap)}</b>.</div>`;return}
  out.innerHTML='<table><thead><tr><th>SAP</th><th>Nombre</th><th>Tipo</th><th>Grupo</th><th>Concepto</th><th>Nuevo cargo</th></tr></thead><tbody>'+rows.map(r=>`<tr><td>${safe(r.sap)}</td><td>${safe(r.nombre)}</td><td>${safe(r.tipo)}</td><td>${safe(r.grupo)}</td><td><span class="status ${statusClass(r.concepto)}">${safe(r.concepto||'Sin concepto')}</span></td><td>${safe(r.nuevo||'-')}</td></tr>`).join('')+'</tbody></table>';
}

/* ===================== ENTREVISTAS ===================== */
function configureInterviewUI(){
  const panel=[...document.querySelectorAll('.panel')].find(p=>p.querySelector('h3')?.textContent.includes('Consulta de Entrevistas'));if(!panel)return;
  const filters=panel.querySelector('.filters.many');if(filters){filters.style.gridTemplateColumns='1.35fr .95fr auto';filters.innerHTML='<input id="iText" placeholder="SAP / Nombre"><input id="iDate" type="date"><button class="btn primary" onclick="filtrarEntrevistas()">Buscar</button>'}
  const table=panel.querySelector('#it');if(table){table.innerHTML='<thead><tr><th>SAP</th><th>Nombre</th><th>Cargo actual</th><th>Cargo al que aspira</th><th>Fecha</th><th>Hora</th><th>Resultado</th></tr></thead><tbody><tr><td colspan="7" class="empty-cell">Carga el archivo de entrevistas para activar la consulta.</td></tr></tbody>'}
  const desc=panel.querySelector('.ph p');if(desc)desc.textContent='Busca por SAP o nombre y consulta el cargo actual, cargo al que aspira y resultado.';
}
async function loadInterviewFile(file){
  try{
    const wb=await readWorkbook(file);const ws=wb.Sheets['Master']||wb.Sheets[wb.SheetNames.find(n=>norm(n)==='master')];if(!ws)throw new Error('No existe la hoja Master');
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});const h=rows[0]||[];
    const ix={sap:firstIndex(h,['SAP']),name:firstIndex(h,['Nombre']),current:firstIndex(h,['Posición Actual','Posicion Actual','Cargo Actual']),future:firstIndex(h,['Posición Futura','Posicion Futura','Cargo Futuro','Cargo al que aspira']),dateInterview:firstIndex(h,['Fecha Entrevista']),time:firstIndex(h,['Hora']),concept:firstIndex(h,['Concepto']),state:firstIndex(h,['Estado']),date:firstIndex(h,['Fecha','Fecha ']),start:firstIndex(h,['Hora de Inicio']),end:firstIndex(h,['Hora Fin','Hora Fin '])};
    if(ix.sap<0&&ix.name<0)throw new Error('No se encontraron columnas SAP/Nombre en Master');
    const result=[];
    for(let r=1;r<rows.length;r++){
      const row=rows[r];const sap=ix.sap>=0?String(row[ix.sap]??'').trim():'';const nombre=ix.name>=0?String(row[ix.name]??'').trim():'';if(!sap&&!nombre)continue;
      const date=excelDate(ix.date>=0&&row[ix.date]?row[ix.date]:(ix.dateInterview>=0?row[ix.dateInterview]:''));
      const concepto=ix.concept>=0?String(row[ix.concept]??'').trim():'';const estado=ix.state>=0?String(row[ix.state]??'').trim():'';
      if(!date&&!concepto&&!estado)continue;
      result.push({sap,nombre,cargoActual:ix.current>=0?String(row[ix.current]??'').trim():'',cargoAspira:ix.future>=0?String(row[ix.future]??'').trim():'',fecha:date,hora:excelTime(ix.start>=0&&row[ix.start]?row[ix.start]:(ix.time>=0?row[ix.time]:'')),horaFin:excelTime(ix.end>=0?row[ix.end]:''),concepto,estado});
    }
    if(!result.length)throw new Error('No se encontraron entrevistas en la hoja Master');
    interviewRecords=result;await cachePut('interview',result,file.name);applyInterviewData(result,file.name);
  }catch(e){console.error('Entrevistas:',e);showSourceState('interviewState',false,`No se pudo leer: ${e.message||'error desconocido'}`)}
}
function applyInterviewData(records,fileName){interviewRecords=records||[];showSourceState('interviewState',true,cacheText(interviewRecords.length,fileName));const k=document.getElementById('kpiInterview');if(k)k.textContent=interviewRecords.filter(r=>r.fecha).length;markLoaded('interview');renderInterviewRows(interviewRecords.slice(0,8))}
function renderInterviewRows(rows){
  const body=document.querySelector('#it tbody');if(!body)return;
  if(!rows.length){body.innerHTML='<tr><td colspan="7" class="empty-cell">Sin resultados</td></tr>';return}
  body.innerHTML=rows.map(r=>`<tr><td>${safe(r.sap)}</td><td>${safe(r.nombre)}</td><td>${safe(r.cargoActual||'-')}</td><td>${safe(r.cargoAspira||'-')}</td><td>${safe(dateDisplay(r.fecha)||'-')}</td><td>${safe(r.hora||'-')}</td><td><span class="status ${statusClass(r.concepto||r.estado)}">${safe(r.concepto||r.estado||'Pendiente')}</span></td></tr>`).join('');
}
function filtrarEntrevistas(){
  if(!interviewRecords.length){renderInterviewRows([]);return}
  const q=norm(document.getElementById('iText')?.value),date=document.getElementById('iDate')?.value||'';
  const rows=interviewRecords.filter(r=>(!q||norm(r.sap+' '+r.nombre).includes(q))&&(!date||r.fecha===date));renderInterviewRows(rows.slice(0,50));
}

/* ===================== AGENDA ===================== */
const monthMap={enero:0,febrero:1,marzo:2,abril:3,mayo:4,junio:5,julio:6,agosto:7,septiembre:8,setiembre:8,octubre:9,noviembre:10,diciembre:11};
function classifyEvent(t){const x=norm(t);if(x.includes('arando'))return'evt-purple';if(x.includes('aula'))return'evt-blue';if(x.includes('circuit'))return'evt-pink';if(x.includes('bienvenida')||x.includes('onboarding'))return'evt-green';if(x.includes('timesoft'))return'evt-yellow';if(x.includes('reunion'))return'evt-orange';return'evt-gray'}
function dayNumber(v){const s=String(v??'').trim();if(!/^\d{1,2}$/.test(s))return null;const n=Number(s);return n>=1&&n<=31?n:null}
async function loadAgendaFile(file){
  try{
    const wb=await readWorkbook(file);const result=[];
    for(const sheetName of wb.SheetNames){
      const mi=monthMap[norm(sheetName)];if(mi===undefined)continue;
      const ws=wb.Sheets[sheetName];if(!ws)continue;
      // raw:false es intencional: este archivo es un calendario visual y así
      // leemos los números de día exactamente como se muestran en Excel.
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''});
      for(let r=0;r<rows.length-1;r++){
        const dayRow=rows[r]||[];const activityRow=rows[r+1]||[];
        const days=[];let valid=0;
        for(let c=0;c<7;c++){const d=dayNumber(dayRow[c]);days.push(d);if(d)valid++}
        if(!valid)continue;
        let found=0;
        for(let c=0;c<7;c++){
          const day=days[c];if(!day)continue;
          const activity=String(activityRow[c]??'').replace(/\s+/g,' ').trim();
          if(!activity||/^[DLMXJVS]$/i.test(activity))continue;
          const date=new Date(2026,mi,day,12,0,0);if(date.getMonth()!==mi)continue;
          result.push({fecha:isoDate(date),actividad,mes:sheetName,tipo:classifyEvent(activity)});found++;
        }
        // Si encontramos actividades, la fila siguiente ya fue usada como detalle.
        if(found)r++;
      }
    }
    if(!result.length)throw new Error('El archivo se abrió, pero no se encontraron actividades. Verifica que sea Agenda de Entrenamiento 2026.xlsx');
    result.sort((a,b)=>a.fecha.localeCompare(b.fecha)||a.actividad.localeCompare(b.actividad));
    agendaRecords=result;await cachePut('agenda',result,file.name);applyAgendaData(result,file.name);
  }catch(e){console.error('Agenda:',e);showSourceState('agendaState',false,`No se pudo leer: ${e.message||'error desconocido'}`)}
}
function applyAgendaData(records,fileName){
  agendaRecords=records||[];showSourceState('agendaState',true,cacheText(agendaRecords.length,fileName,'eventos'));const k=document.getElementById('kpiAgenda');if(k)k.textContent=agendaRecords.length;markLoaded('agenda');
  if(agendaRecords.length){calendarCursor=startOfWeek(new Date(agendaRecords[agendaRecords.length-1].fecha+'T12:00:00'));renderAgendaWeek()}
}
function moveWeek(n){if(!calendarCursor)calendarCursor=startOfWeek(new Date());calendarCursor=addDays(calendarCursor,7*n);renderAgendaWeek()}
function goToday(){calendarCursor=startOfWeek(new Date());renderAgendaWeek()}
function goLatest(){if(!agendaRecords.length)return;calendarCursor=startOfWeek(new Date(agendaRecords[agendaRecords.length-1].fecha+'T12:00:00'));renderAgendaWeek()}
function renderAgendaWeek(){
  const wrap=document.getElementById('calendarView');if(!wrap)return;
  if(!calendarCursor){wrap.innerHTML='<div class="empty-state agenda-empty">Carga el archivo de Agenda para mostrar la programación real.</div>';return}
  const labels=['Lun','Mar','Mié','Jue','Vie','Sáb','Dom'];let html='<div class="agenda-week">';
  for(let i=0;i<7;i++){
    const d=addDays(calendarCursor,i),iso=isoDate(d),events=agendaRecords.filter(e=>e.fecha===iso);
    html+=`<div class="agenda-day"><div class="agenda-day-head"><b>${labels[i]} ${d.getDate()}</b><span>${d.toLocaleDateString('es-CO',{month:'short'})}</span></div><div class="agenda-day-body">${events.length?events.map(e=>`<div class="agenda-event ${e.tipo}"><b>${safe(e.actividad).split(' / ').join('<br>')}</b><span>Actividad del calendario de entrenamiento</span></div>`).join(''):'<div class="no-event">Sin programación</div>'}</div></div>`;
  }
  html+='</div>';wrap.innerHTML=html;
  const title=document.getElementById('calendarPeriod');if(title){const end=addDays(calendarCursor,6);title.textContent=`${calendarCursor.getDate()} ${calendarCursor.toLocaleDateString('es-CO',{month:'short'})} – ${end.getDate()} ${end.toLocaleDateString('es-CO',{month:'short'})} ${end.getFullYear()}`}
}

/* ===================== INICIO / RESTAURACIÓN ===================== */
function wireFile(id,handler){const input=document.getElementById(id);if(input)input.addEventListener('change',e=>{const f=e.target.files?.[0];if(f)handler(f)})}
async function restoreCachedData(){
  try{
    const [a,i,g]=await Promise.all([cacheGet('arando'),cacheGet('interview'),cacheGet('agenda')]);
    if(a?.records?.length)applyArandoData(a.records,a.fileName);
    if(i?.records?.length)applyInterviewData(i.records,i.fileName);
    if(g?.records?.length)applyAgendaData(g.records,g.fileName);
  }catch(e){console.warn('No se pudo restaurar la caché local:',e)}
}

configureInterviewUI();
wireFile('agendaFile',loadAgendaFile);wireFile('arandoFile',loadArandoFile);wireFile('interviewFile',loadInterviewFile);
document.getElementById('globalSearch')?.addEventListener('input',e=>{const q=norm(e.target.value);document.querySelectorAll('.searchable').forEach(x=>x.classList.toggle('hidden',q&&!norm((x.dataset.search||'')+' '+x.innerText).includes(q)))})
document.getElementById('sap')?.addEventListener('keydown',e=>{if(e.key==='Enter')buscarSAP()});
restoreCachedData().then(()=>{if(!agendaRecords.length)renderAgendaWeek()});