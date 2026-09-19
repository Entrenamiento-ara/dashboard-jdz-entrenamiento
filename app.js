const excelWeb='https://jmartins.sharepoint.com/:x:/r/sites/EntrenamientoR3/_layouts/15/Doc.aspx?sourcedoc=%7BB762772B-6048-4002-AF66-F0BC9A0A1055%7D&file=Ingresos%20con%20tiendas%20de%20entrenamiento%20y%20pares%20formadores.xlsm&action=default&mobileredirect=true';
function abrirExcel(){location.href='ms-excel:ofe|u|'+excelWeb}

let arandoRecords=[];
let interviewRecords=[];
let agendaRecords=[];
let calendarCursor=null;
let loadedSources=0;
const loadedFlags={agenda:false,arando:false,interview:false};

function norm(v){return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toLowerCase()}
function safe(v){return String(v??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}
function statusClass(v){const x=norm(v);if((x.includes('aprueba')&&!x.includes('no aprueba'))||x.includes('realizado')||x.includes('complet'))return'ok';if(x.includes('no aprueba')||x.includes('cancel')||x.includes('no apto'))return'bad';if(x.includes('pend')||x.includes('proceso')||x.includes('evalu'))return'warn';return'info'}
function dateDisplay(iso){if(!iso)return'';const [y,m,d]=iso.split('-');return `${d}/${m}/${y}`}
function isoDate(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function addDays(d,n){const x=new Date(d);x.setDate(x.getDate()+n);return x}
function startOfWeek(d){const x=new Date(d);const day=x.getDay();const diff=day===0?-6:1-day;x.setDate(x.getDate()+diff);x.setHours(12,0,0,0);return x}

function excelDate(value){
  if(value===null||value===undefined||value==='')return'';
  if(value instanceof Date&&!isNaN(value))return isoDate(value);
  if(typeof value==='number'){
    const d=XLSX.SSF.parse_date_code(value);
    if(!d)return'';
    return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }
  const s=String(value).trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;
  let m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if(m){let y=Number(m[3]);if(y<100)y+=2000;return `${y}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`}
  const d=new Date(s);return isNaN(d)?'':isoDate(d);
}
function excelTime(value){
  if(value===null||value===undefined||value==='')return'';
  if(value instanceof Date&&!isNaN(value))return `${String(value.getHours()).padStart(2,'0')}:${String(value.getMinutes()).padStart(2,'0')}`;
  if(typeof value==='number'){
    const total=Math.round((value%1)*24*60);
    return `${String(Math.floor(total/60)%24).padStart(2,'0')}:${String(total%60).padStart(2,'0')}`;
  }
  const s=String(value).trim();const m=s.match(/(\d{1,2}):(\d{2})/);return m?`${String(m[1]).padStart(2,'0')}:${m[2]}`:s;
}
function readWorkbook(file){
  return new Promise((resolve,reject)=>{
    const fr=new FileReader();
    fr.onload=e=>{try{resolve(XLSX.read(e.target.result,{type:'array',cellDates:false,cellText:true}))}catch(err){reject(err)}};
    fr.onerror=()=>reject(fr.error||new Error('No se pudo abrir el archivo'));
    fr.readAsArrayBuffer(file);
  });
}
function showSourceState(id,ok,text){const el=document.getElementById(id);if(!el)return;el.textContent=text;el.className='source-state '+(ok?'source-ok':'source-bad')}
function markLoaded(key){if(loadedFlags[key])return;loadedFlags[key]=true;loadedSources++;const el=document.getElementById('kpiSources');if(el)el.textContent=loadedSources}
function headerIndex(headers,name){const n=norm(name);return headers.findIndex(h=>norm(h)===n)}
function firstIndex(headers,names){for(const n of names){const i=headerIndex(headers,n);if(i>=0)return i}return-1}

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
    arandoRecords=result;showSourceState('arandoState',true,`${result.length} registros cargados`);const k=document.getElementById('kpiArando');if(k)k.textContent=result.length;markLoaded('arando');document.getElementById('sapResult').innerHTML='<div class="empty-state">Archivo cargado. Escribe un SAP para consultar el historial.</div>';
  }catch(e){console.error('Arando:',e);showSourceState('arandoState',false,`No se pudo leer: ${e.message||'error desconocido'}`)}
}
function buscarSAP(){
  const sap=document.getElementById('sap').value.trim();const out=document.getElementById('sapResult');
  if(!arandoRecords.length){out.innerHTML='<div class="empty-state">Primero carga el archivo de Arando Talento.</div>';return}
  if(!sap){out.innerHTML='<div class="empty-state">Escribe un SAP para consultar.</div>';return}
  const rows=arandoRecords.filter(r=>r.sap===sap);
  if(!rows.length){out.innerHTML=`<div class="empty-state">No se encontraron registros para el SAP <b>${safe(sap)}</b>.</div>`;return}
  out.innerHTML='<table><thead><tr><th>SAP</th><th>Nombre</th><th>Tipo</th><th>Grupo</th><th>Concepto</th><th>Nuevo cargo</th></tr></thead><tbody>'+rows.map(r=>`<tr><td>${safe(r.sap)}</td><td>${safe(r.nombre)}</td><td>${safe(r.tipo)}</td><td>${safe(r.grupo)}</td><td><span class="status ${statusClass(r.concepto)}">${safe(r.concepto||'Sin concepto')}</span></td><td>${safe(r.nuevo||'-')}</td></tr>`).join('')+'</tbody></table>';
}

function simpleProgram(v){const x=norm(v);if(x.includes('arando talento'))return'Arando Talento';if(x.includes('promocion directa'))return'Promoción directa';if(x.includes('entrevista de promocion'))return'Entrevista de Promoción';return String(v||'Otro').trim()||'Otro'}
async function loadInterviewFile(file){
  try{
    const wb=await readWorkbook(file);const ws=wb.Sheets['Master']||wb.Sheets[wb.SheetNames.find(n=>norm(n)==='master')];if(!ws)throw new Error('No existe la hoja Master');
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:''});const h=rows[0]||[];
    const ix={sap:firstIndex(h,['SAP']),name:firstIndex(h,['Nombre']),program:firstIndex(h,['Programa']),dateInterview:firstIndex(h,['Fecha Entrevista']),time:firstIndex(h,['Hora']),rrhh:firstIndex(h,['Responsable Entrevista RRHH']),ops:firstIndex(h,['Responsable Entrevista Operaciones']),concept:firstIndex(h,['Concepto']),state:firstIndex(h,['Estado']),date:firstIndex(h,['Fecha']),start:firstIndex(h,['Hora de Inicio']),end:firstIndex(h,['Hora Fin']),mode:firstIndex(h,['Modalidad']),place:firstIndex(h,['Lugar entrevista presencial'])};
    if(ix.sap<0&&ix.name<0)throw new Error('No se encontraron columnas SAP/Nombre en Master');
    const result=[];
    for(let r=1;r<rows.length;r++){
      const row=rows[r];const sap=ix.sap>=0?String(row[ix.sap]??'').trim():'';const nombre=ix.name>=0?String(row[ix.name]??'').trim():'';if(!sap&&!nombre)continue;
      const date=excelDate(ix.date>=0&&row[ix.date]?row[ix.date]:(ix.dateInterview>=0?row[ix.dateInterview]:''));
      if(!date&&norm(ix.concept>=0?row[ix.concept]:'')==='')continue;
      result.push({sap,nombre,programa:simpleProgram(ix.program>=0?row[ix.program]:''),fecha:date,hora:excelTime(ix.start>=0&&row[ix.start]?row[ix.start]:(ix.time>=0?row[ix.time]:'')),horaFin:excelTime(ix.end>=0?row[ix.end]:''),modalidad:ix.mode>=0?String(row[ix.mode]??'').trim():'',rrhh:ix.rrhh>=0?String(row[ix.rrhh]??'').trim():'',ops:ix.ops>=0?String(row[ix.ops]??'').trim():'',concepto:ix.concept>=0?String(row[ix.concept]??'').trim():'',estado:ix.state>=0?String(row[ix.state]??'').trim():'',lugar:ix.place>=0?String(row[ix.place]??'').trim():''});
    }
    interviewRecords=result;showSourceState('interviewState',true,`${result.length} registros cargados`);const k=document.getElementById('kpiInterview');if(k)k.textContent=result.filter(r=>r.fecha).length;markLoaded('interview');
    const sel=document.getElementById('iType');const types=[...new Set(result.map(r=>r.programa).filter(Boolean))].sort();sel.innerHTML='<option>Todos</option>'+types.map(t=>`<option>${safe(t)}</option>`).join('');renderInterviewRows(result.slice(0,8));
  }catch(e){console.error('Entrevistas:',e);showSourceState('interviewState',false,`No se pudo leer: ${e.message||'error desconocido'}`)}
}
function renderInterviewRows(rows){const body=document.querySelector('#it tbody');if(!body)return;if(!rows.length){body.innerHTML='<tr><td colspan="7" class="empty-cell">Sin resultados</td></tr>';return}body.innerHTML=rows.map(r=>`<tr><td>${safe(r.sap)}</td><td>${safe(r.nombre)}</td><td>${safe(r.programa)}</td><td>${safe(dateDisplay(r.fecha))}</td><td>${safe(r.hora||'-')}</td><td>${safe(r.modalidad||'-')}</td><td><span class="status ${statusClass(r.concepto||r.estado)}">${safe(r.concepto||r.estado||'Pendiente')}</span></td></tr>`).join('')}
function filtrarEntrevistas(){if(!interviewRecords.length){renderInterviewRows([]);return}const q=norm(document.getElementById('iText').value),type=document.getElementById('iType').value,date=document.getElementById('iDate').value;const rows=interviewRecords.filter(r=>(!q||norm(r.sap+' '+r.nombre).includes(q))&&(type==='Todos'||r.programa===type)&&(!date||r.fecha===date));renderInterviewRows(rows.slice(0,50))}

const monthMap={enero:0,febrero:1,marzo:2,abril:3,mayo:4,junio:5,julio:6,agosto:7,septiembre:8,setiembre:8,octubre:9,noviembre:10,diciembre:11};
function classifyEvent(t){const x=norm(t);if(x.includes('arando'))return'evt-purple';if(x.includes('aula'))return'evt-blue';if(x.includes('circuit'))return'evt-pink';if(x.includes('bienvenida')||x.includes('onboarding'))return'evt-green';if(x.includes('timesoft'))return'evt-yellow';if(x.includes('reunion'))return'evt-orange';return'evt-gray'}
function agendaDayFromCell(cell){
  if(!cell)return null;
  const v=cell.v;
  if(typeof v==='number'){
    if(Number.isInteger(v)&&v>=1&&v<=31)return v;
    const d=XLSX.SSF.parse_date_code(v);
    if(d&&d.d>=1&&d.d<=31&&d.y===1900)return d.d;
  }
  if(v instanceof Date&&!isNaN(v))return v.getDate();
  const text=String(cell.w??v??'').trim();
  if(/^\d{1,2}$/.test(text)){const n=Number(text);if(n>=1&&n<=31)return n}
  const m=text.match(/^(\d{1,2})[\/\-]/);if(m){const n=Number(m[1]);if(n>=1&&n<=31)return n}
  return null;
}
function agendaCellText(cell){if(!cell)return'';return String(cell.w??cell.v??'').replace(/\s+/g,' ').trim()}
function agendaYearMonth(ws,sheetName){
  let year=2026,month=monthMap[norm(sheetName)];
  const c=ws['A1'];if(c){
    if(typeof c.v==='number'){
      const d=XLSX.SSF.parse_date_code(c.v);if(d&&d.y>=2000&&d.y<=2100){year=d.y;month=d.m-1}
    }else if(c.v instanceof Date&&!isNaN(c.v)){year=c.v.getFullYear();month=c.v.getMonth()}
    else{const d=new Date(String(c.v));if(!isNaN(d)&&d.getFullYear()>=2000){year=d.getFullYear();month=d.getMonth()}}
  }
  return{year,month};
}
async function loadAgendaFile(file){
  try{
    const wb=await readWorkbook(file);const result=[];
    for(const sheetName of wb.SheetNames){
      const ws=wb.Sheets[sheetName];if(!ws||!ws['!ref'])continue;
      const ym=agendaYearMonth(ws,sheetName);if(ym.month===undefined||ym.month===null)continue;
      const range=XLSX.utils.decode_range(ws['!ref']);
      for(let r=range.s.r;r<range.e.r;r++){
        const days=[];let validDays=0;
        for(let c=0;c<7;c++){
          const day=agendaDayFromCell(ws[XLSX.utils.encode_cell({r,c})]);days.push(day);if(day)validDays++;
        }
        if(!validDays)continue;
        let activities=0;
        for(let c=0;c<7;c++){
          const activity=agendaCellText(ws[XLSX.utils.encode_cell({r:r+1,c})]);
          if(!activity)continue;
          const day=days[c];if(!day)continue;
          if(/^[DLMXJVS]$/i.test(activity))continue;
          const d=new Date(ym.year,ym.month,day,12,0,0);if(d.getMonth()!==ym.month)continue;
          result.push({fecha:isoDate(d),actividad,mes:sheetName,tipo:classifyEvent(activity)});activities++;
        }
        if(activities)r++;
      }
    }
    agendaRecords=result.sort((a,b)=>a.fecha.localeCompare(b.fecha)||a.actividad.localeCompare(b.actividad));
    if(!agendaRecords.length)throw new Error('El archivo se abrió, pero no se encontraron actividades en el formato calendario esperado');
    showSourceState('agendaState',true,`${agendaRecords.length} eventos cargados`);const k=document.getElementById('kpiAgenda');if(k)k.textContent=agendaRecords.length;markLoaded('agenda');
    calendarCursor=startOfWeek(new Date(agendaRecords[agendaRecords.length-1].fecha+'T12:00:00'));renderAgendaWeek();
  }catch(e){console.error('Agenda:',e);showSourceState('agendaState',false,`No se pudo leer: ${e.message||'error desconocido'}`)}
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
    html+=`<div class="agenda-day"><div class="agenda-day-head"><b>${labels[i]} ${d.getDate()}</b><span>${d.toLocaleDateString('es-CO',{month:'short'})}</span></div><div class="agenda-day-body">${events.length?events.map(e=>`<div class="agenda-event ${e.tipo}"><b>${safe(e.actividad).split(' / ').join('<br>')}</b><span>Sin hora registrada en el archivo</span></div>`).join(''):'<div class="no-event">Sin programación</div>'}</div></div>`;
  }
  html+='</div>';wrap.innerHTML=html;
  const title=document.getElementById('calendarPeriod');if(title){const end=addDays(calendarCursor,6);title.textContent=`${calendarCursor.getDate()} ${calendarCursor.toLocaleDateString('es-CO',{month:'short'})} – ${end.getDate()} ${end.toLocaleDateString('es-CO',{month:'short'})} ${end.getFullYear()}`}
}

function wireFile(id,handler){const input=document.getElementById(id);if(input)input.addEventListener('change',e=>{const f=e.target.files?.[0];if(f)handler(f)})}
wireFile('agendaFile',loadAgendaFile);wireFile('arandoFile',loadArandoFile);wireFile('interviewFile',loadInterviewFile);

document.getElementById('globalSearch')?.addEventListener('input',e=>{const q=norm(e.target.value);document.querySelectorAll('.searchable').forEach(x=>x.classList.toggle('hidden',q&&!norm((x.dataset.search||'')+' '+x.innerText).includes(q)))})
document.getElementById('sap')?.addEventListener('keydown',e=>{if(e.key==='Enter')buscarSAP()});
renderAgendaWeek();