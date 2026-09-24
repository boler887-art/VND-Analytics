(() => {
'use strict';

const DEFAULT_ALLOWED_EXTENSIONS = ['pdf','docx','doc','xlsx','xls','txt','jpg','jpeg','png','tif','tiff','webp'];
const DEFAULT_CFG = {
  systemName:'NORMATRIX VND Analytics',version:'1.2.3',locale:'ru-RU',
  dataSources:{vndRegistry:'vnd_master.xlsx',structure:'structure.xlsx'},
  privacy:{zeroPersistence:true,useLocalStorage:false,useSessionStorage:false,persistUploadedFiles:false,persistAnalysisResults:false},
  limits:{maxVndPackages:300,maxPhysicalFiles:5000,maxSingleFileMB:100,maxSessionMB:2048,maxAutoOcrImages:40,maxManualOcrPages:1200},
  analysis:{exactDuplicateMinChars:60,nearDuplicateThreshold:.72,conflictSimilarityThreshold:.55,maxCandidatePairs:100000,autoRunAfterIngest:false,autoOcrImages:true,includeHistoricalRegistryRows:true,legacyDocLocalExtraction:true,manualOcrQueue:true,conflictEngine:'2.0'},
  legalVerification:{enabled:true,mode:'metadata-only',endpoint:'',allowSendingInternalText:false,requireExplicitConsentForText:true}
};

const OPTIONAL_DEPENDENCIES = [
  {global:'XLSX',label:'SheetJS / XLSX',urls:['https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js','https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js']},
  {global:'mammoth',label:'Mammoth',urls:['https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js','https://unpkg.com/mammoth@1.8.0/mammoth.browser.min.js']},
  {global:'pdfjsLib',label:'PDF.js',urls:['https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js','https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js']},
  {global:'Tesseract',label:'Tesseract OCR',urls:['https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js','https://unpkg.com/tesseract.js@5.1.1/dist/tesseract.min.js']}
];
const PDF_WORKERS=['https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js','https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js'];

const S = {
  cfg:structuredClone(DEFAULT_CFG),rules:null,mapping:null,aliases:null,glossary:null,legalSources:null,version:null,
  registry:[],structure:[],structureMap:new Map(),packages:[],issues:[],legalRefs:[],duplicates:[],conflicts:[],quality:[],ingestErrors:[],recommendations:[],
  currentPage:'dashboard',issueFilter:'',conflictFilter:'',processing:false,staticReady:false,objectUrls:new Set(),ocrCount:0,importedReviewDecisions:new Map(),importedReviewData:new Map(),analysisPairLimitHit:false,analysisCandidatePairs:0
};
const $ = s => document.querySelector(s);
const $$ = s => [...document.querySelectorAll(s)];
const esc = v => String(v ?? '').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));
const norm = v => String(v ?? '').toLowerCase().replace(/ё/g,'е').replace(/[«»„“”"'`]/g,'').replace(/[^a-zа-яәіңғүұқөһ0-9]+/giu,' ').replace(/\s+/g,' ').trim();
const words = v => norm(v).split(' ').filter(x=>x.length>2 && !STOP.has(x));
const uniq = a => [...new Set(a.filter(Boolean))];
const ext = name => (name.split('.').pop()||'').toLowerCase();
const bytesMB = n => n/1024/1024;
const nowIso = () => new Date().toISOString();
const STOP = new Set('для при или как что это его ее их был была были быть который которая которые также согласно соответствии настоящего настоящей настоящих настоящим путем после перед между либо только такой такие этого этой этих той той же все всех каждой каждый пункт пункта статье статья раздел приложение товарищества организации документ документа документов'.split(' '));

function toast(msg, kind='info'){
  const d=document.createElement('div'); d.className=`callout ${kind==='error'?'danger':kind==='ok'?'ok':kind==='warn'?'warn':''}`;
  Object.assign(d.style,{position:'fixed',right:'18px',bottom:'18px',zIndex:9999,maxWidth:'440px',boxShadow:'0 14px 40px rgba(0,0,0,.18)'});
  d.innerHTML=esc(msg); document.body.appendChild(d); setTimeout(()=>d.remove(),4200);
}
function fmtDate(v){
  if(v==null||v==='') return '';
  if(typeof v==='number' && v>20000 && v<80000){ const d=new Date(Math.round((v-25569)*86400*1000)); return d.toLocaleDateString('ru-RU'); }
  const d=new Date(v); return isNaN(d)?String(v):d.toLocaleDateString('ru-RU');
}
function normalizeHeader(v){ return norm(v).replace(/\s+/g,' '); }
async function fetchNoCache(url, as='json'){
  const sep=url.includes('?')?'&':'?'; const res=await fetch(`${url}${sep}_=${Date.now()}`,{cache:'no-store',credentials:'same-origin'});
  if(!res.ok) throw new Error(`${url}: HTTP ${res.status}`);
  return as==='arrayBuffer'?res.arrayBuffer():as==='text'?res.text():res.json();
}
function findHeaderRow(aoa, aliases){
  const targets=Object.values(aliases).flat().map(normalizeHeader);
  for(let i=0;i<Math.min(20,aoa.length);i++){
    const row=(aoa[i]||[]).map(normalizeHeader);
    const hits=targets.filter(t=>row.includes(t)).length;
    if(hits>=Math.min(3,Object.keys(aliases).length)) return i;
  }
  return -1;
}
function headerIndexMap(row, aliases){
  const out={}; const nr=row.map(normalizeHeader);
  for(const [key,arr] of Object.entries(aliases)){
    let idx=-1; for(const a of arr){ idx=nr.indexOf(normalizeHeader(a)); if(idx>=0) break; } out[key]=idx;
  } return out;
}
function val(row, idx){ return idx>=0 ? row[idx] : ''; }
function inferType(title){
  const t=norm(title); const types=[['правила','Правила'],['положение','Положение'],['инструкция','Инструкция'],['регламент','Регламент'],['политик','Политика'],['методик','Методика'],['руководств','Руководство'],['кодекс','Кодекс'],['реестр','Реестр'],['стандарт','Стандарт'],['программ','Программа'],['план','План'],['порядок','Порядок']];
  for(const [k,n] of types) if(t.includes(k)) return n; return 'Иное';
}
function canonicalPerson(v){
  const n=norm(v); if(!n) return '';
  const persons=S.aliases?.persons||{};
  for(const [canon,als] of Object.entries(persons)) if(norm(canon)===n || als.some(a=>norm(a)===n)) return canon;
  return String(v).trim();
}

function loadScriptUrl(url,timeoutMs=5000){
  return new Promise((resolve,reject)=>{
    const sc=document.createElement('script');let done=false;
    const finish=(ok,err)=>{if(done)return;done=true;clearTimeout(timer);sc.onload=sc.onerror=null;if(!ok)sc.remove();ok?resolve(url):reject(err||new Error(`Не удалось загрузить ${url}`));};
    const timer=setTimeout(()=>finish(false,new Error(`Таймаут загрузки ${url}`)),timeoutMs);
    sc.async=true;sc.src=url;sc.referrerPolicy='no-referrer';sc.onload=()=>finish(true);sc.onerror=()=>finish(false,new Error(`Ошибка загрузки ${url}`));document.head.appendChild(sc);
  });
}
async function loadDependency(dep){
  if(window[dep.global])return {label:dep.label,ok:true,url:'already-loaded'};
  let last=null;for(const url of dep.urls){try{await loadScriptUrl(url);if(window[dep.global])return {label:dep.label,ok:true,url};last=new Error(`${dep.label}: глобальный объект не создан`);}catch(e){last=e;}}
  return {label:dep.label,ok:false,error:last?.message||'недоступна'};
}
async function loadOptionalDependencies(){
  const results=await Promise.all(OPTIONAL_DEPENDENCIES.map(loadDependency));
  if(window.pdfjsLib){pdfjsLib.GlobalWorkerOptions.workerSrc=PDF_WORKERS[0];window.__NORMATRIX_PDF_WORKER_FALLBACK=PDF_WORKERS[1];}
  for(const r of results)if(!r.ok)console.warn(`${r.label} unavailable: ${r.error}`);
  return results;
}
function updateRunAnalysisButton(){const b=$('#runAnalysisBtn');if(b)b.disabled=!S.staticReady||!S.packages.some(p=>!p.importedAnalysis);}

async function loadStaticData(){
  try{
    const results=await Promise.allSettled([
      fetchNoCache('config.json'),fetchNoCache('rules.json'),fetchNoCache('mapping.json'),fetchNoCache('aliases.json'),fetchNoCache('glossary.json'),fetchNoCache('legal_sources.json'),fetchNoCache('system_version.json')
    ]);
    const [cfg,rules,mapping,aliases,glossary,legalSources,version]=results.map(r=>r.status==='fulfilled'?r.value:null);
    if(cfg) S.cfg={...structuredClone(DEFAULT_CFG),...cfg,limits:{...DEFAULT_CFG.limits,...(cfg.limits||{})},analysis:{...DEFAULT_CFG.analysis,...(cfg.analysis||{})},dataSources:{...DEFAULT_CFG.dataSources,...(cfg.dataSources||{})}};
    S.rules=rules||S.rules||{legalMonitoringCategories:[]};
    S.mapping=mapping||S.mapping; S.aliases=aliases||S.aliases||{persons:{}}; S.glossary=glossary||S.glossary||{}; S.legalSources=legalSources||S.legalSources||{}; S.version=version||S.version||{version:'1.2.3'};
    if(window.pdfjsLib && !pdfjsLib.GlobalWorkerOptions.workerSrc) pdfjsLib.GlobalWorkerOptions.workerSrc=PDF_WORKERS[0];
    if(window.XLSX && S.mapping){
      try{ await Promise.all([loadRegistry(),loadStructure()]); buildStructureMap(); linkRegistryStructure();
        if(S.packages.length)refreshMatches(); setChip('#chipRegistry',`Реестр: ${S.registry.length} ВНД`,true); setChip('#chipStructure',`Структура: ${S.structure.length} записей`,true);
      }catch(e){console.error('Static XLSX load failed',e);setChip('#chipRegistry','Ошибка реестра',false);setChip('#chipStructure','Ошибка структуры',false);S.ingestErrors.push({level:'error',type:'Корневые данные',message:e.message||String(e)});}
    }else{
      setChip('#chipRegistry','XLSX-библиотека недоступна',false);setChip('#chipStructure','XLSX-библиотека недоступна',false);
      S.ingestErrors.push({level:'error',type:'Библиотека XLSX',message:'Не загружена SheetJS/XLSX. Загрузка ВНД в сессию доступна, но реестр/structure.xlsx и XLSX-анализ недоступны до загрузки библиотеки.'});
    }
    S.staticReady=true;if(S.packages.some(p=>p.importedAnalysis))rebuildAnalysisFromParsedFiles();updateRunAnalysisButton();renderAll();
  }catch(e){ console.error(e); setChip('#chipRegistry','Ошибка источников',false); setChip('#chipStructure','Ошибка источников',false); S.ingestErrors.push({level:'error',type:'Инициализация',message:e.message||String(e)}); S.staticReady=true;updateRunAnalysisButton();renderAll(); toast('Часть справочников не загрузилась. Загрузка файлов остаётся доступной; подробности — «Качество данных».','warn'); }
}
async function loadRegistry(){
  const ab=await fetchNoCache(S.cfg.dataSources.vndRegistry,'arrayBuffer'); const wb=XLSX.read(ab,{type:'array'});
  const sn=wb.SheetNames.includes(S.mapping.registrySheet)?S.mapping.registrySheet:wb.SheetNames[0];
  const aoa=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:''}); const hi=findHeaderRow(aoa,S.mapping.registryHeaders);
  if(hi<0) throw new Error('Шапка реестра не найдена'); const idx=headerIndexMap(aoa[hi],S.mapping.registryHeaders);
  const out=[]; let temp=900000;
  for(let r=hi+1;r<aoa.length;r++){
    const row=aoa[r]; const title=String(val(row,idx.title)||'').trim(); if(!title) continue;
    const seq=val(row,idx.sequence); if(typeof seq==='string' && /совет.*директор/i.test(seq) && !title) continue;
    let vndId=String(val(row,idx.vndId)||'').trim(); let idTemporary=false;
    if(!/^VND-\d{6}$/i.test(vndId)){ vndId=`VND-${temp++}`; idTemporary=true; }
    const responsible=canonicalPerson(val(row,idx.responsible));
    out.push({
      row:r+1,vndId:vndId.toUpperCase(),idTemporary,sequence:seq,number:String(val(row,idx.number)||'').trim(),date:val(row,idx.date),year:String(val(row,idx.year)||'').trim(),
      title,responsible,change:String(val(row,idx.change)||'').trim(),status:String(val(row,idx.status)||'').trim(),packageName:String(val(row,idx.packageName)||'').trim(),type:inferType(title),department:'Не определено',spId:''
    });
  }
  S.registry=out;
}
async function loadStructure(){
  const ab=await fetchNoCache(S.cfg.dataSources.structure,'arrayBuffer'); const wb=XLSX.read(ab,{type:'array'});
  const sn=wb.SheetNames.includes(S.mapping.structureSheet)?S.mapping.structureSheet:wb.SheetNames[0];
  const aoa=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:''}); const hi=findHeaderRow(aoa,S.mapping.structureHeaders);
  if(hi<0) throw new Error('Шапка structure.xlsx не найдена'); const idx=headerIndexMap(aoa[hi],S.mapping.structureHeaders);
  S.structure=aoa.slice(hi+1).filter(r=>val(r,idx.spId)||val(r,idx.department)||val(r,idx.responsible)).map(r=>({
    spId:String(val(r,idx.spId)||'').trim(),department:String(val(r,idx.department)||'').trim(),shortName:String(val(r,idx.shortName)||'').trim(),responsible:canonicalPerson(val(r,idx.responsible)),aliases:String(val(r,idx.aliases)||'').split(';').map(x=>x.trim()).filter(Boolean),status:String(val(r,idx.status)||'').trim()
  }));
}
function buildStructureMap(){
  S.structureMap=new Map(); for(const s of S.structure){ const keys=[s.responsible,...s.aliases]; for(const k of keys) if(k) S.structureMap.set(norm(canonicalPerson(k)),s); }
}
function linkRegistryStructure(){
  for(const r of S.registry){ const s=S.structureMap.get(norm(r.responsible)); if(s){r.department=s.department||'Не определено';r.spId=s.spId||'';} }
}
function setChip(sel,text,ok){ const el=$(sel); if(!el)return; el.textContent=text; el.classList.remove('ok','err'); el.classList.add(ok?'ok':'err'); }

function initNav(){
  $('#mainNav').addEventListener('click',e=>{const b=e.target.closest('button[data-page]');if(b) goPage(b.dataset.page);});
  document.addEventListener('click',e=>{const b=e.target.closest('[data-go]');if(b)goPage(b.dataset.go);});
}
function goPage(p){ S.currentPage=p; $$('.page').forEach(x=>x.classList.toggle('active',x.id===`page-${p}`)); $$('#mainNav button').forEach(x=>x.classList.toggle('active',x.dataset.page===p)); const b=$(`#mainNav button[data-page="${p}"]`); $('#topTitle').textContent=b?b.textContent.replace(/\d+$/,'').trim():'NORMATRIX'; window.scrollTo({top:0,behavior:'smooth'}); }

function renderAll(){
  $('#navRegistryCount').textContent=S.registry.length; $('#navIssuesCount').textContent=S.issues.length;
  renderDashboard(); renderRegistryFilters(); renderRegistry(); renderPackages(); renderMonitoring(); renderDuplicates(); renderConflicts(); renderLegalRefs(); renderDepartments(); renderQuality(); renderReports(); updateOcrButton();
}
function loadedRegistryIds(){ return new Set(S.packages.filter(p=>p.match?.vndId).map(p=>p.match.vndId)); }
function packageFor(vndId){ return S.packages.find(p=>p.match?.vndId===vndId || p.vndId===vndId); }
function issuesFor(vndId){ return S.issues.filter(i=>i.vndId===vndId || i.otherVndId===vndId); }
function renderDashboard(){
  const loaded=S.packages.length, sourceLoaded=S.packages.filter(p=>!p.importedAnalysis).length, reportLoaded=S.packages.filter(p=>p.importedAnalysis).length, parsed=S.packages.filter(p=>p.parsed).length, units=S.packages.reduce((a,p)=>a+(p.units?.length||0),0), refs=S.legalRefs.length;
  const confirmed=S.issues.filter(i=>i.reviewStatus==='confirmed').length; const missingIds=S.registry.filter(r=>r.idTemporary).length;
  const cards=[
    ['ВНД в реестре',S.registry.length,'registry',''],['Всего в сессии',loaded,'upload','ok'],['Исходных ВНД',sourceLoaded,'upload',''],['Отчётов ВНД',reportLoaded,'upload','info'],['Готово к анализу',parsed,'upload','info'],['Выделено норм',units,'monitoring',''],['Найдено замечаний',S.issues.length,'monitoring',S.issues.length?'warn':'ok'],['Подтверждено',confirmed,'monitoring',confirmed?'danger':''],
    ['Нормативных ссылок',refs,'legalrefs','info'],['Дубли',S.duplicates.length,'duplicates',S.duplicates.length?'warn':''],['Противоречия',S.conflicts.length,'conflicts',S.conflicts.length?'danger':''],['Рекомендаций СП',S.recommendations.length,'departments',''],['ID требуют фиксации',missingIds,'quality',missingIds?'warn':'ok'],['Ошибок данных',S.quality.filter(q=>q.level==='error').length,'quality','danger']
  ];
  $('#dashboardKpis').innerHTML=cards.map(c=>`<div class="kpi-card ${c[3]}" data-go="${c[2]}"><div class="label">${esc(c[0])}</div><div class="value">${c[1]}</div><div class="hint">нажмите для детализации</div></div>`).join('');
  renderCategoryCards('#dashboardCategories'); renderCoverage();
  const top=S.issues.slice(0,8); $('#dashboardIssues').innerHTML=top.length?issueTable(top,true):`<div class="empty"><strong>Замечаний пока нет</strong>Загрузите и проанализируйте документы.</div>`;
  $('#monitoringMeta').textContent=S.issues.length?`${S.issues.length} кандидатов на замечание`:'Нет активного анализа';
}
function renderCoverage(){
  const ids=loadedRegistryIds(); const matched=ids.size; const total=S.registry.length; const pct=total?Math.round(matched/total*100):0;
  $('#coveragePanel').innerHTML=S.packages.length?`<div class="stat-line"><div class="stat-pill"><b>${matched}</b>связано с реестром</div><div class="stat-pill"><b>${S.packages.length-matched}</b>без точного соответствия</div><div class="stat-pill"><b>${pct}%</b>охват реестра</div></div><div class="progress" style="margin-top:14px"><span style="width:${pct}%"></span></div><div class="footer-note">Выводы относятся только к загруженным ВНД. Незагруженные документы не участвуют в междокументном сравнении.</div>`:`<div class="empty"><strong>Документы ещё не загружены</strong>Реестр уже доступен. Загрузите ВНД для анализа содержания.</div>`;
}
function renderCategoryCards(sel){
  const cats=S.rules?.legalMonitoringCategories||[]; const target=$(sel); if(!target)return;
  target.innerHTML=cats.map(c=>{const all=S.issues.filter(i=>i.categoryId===c.id),confirmed=all.filter(i=>i.reviewStatus==='confirmed').length,pending=all.filter(i=>i.reviewStatus==='ai'||i.reviewStatus==='legal').length;return `<div class="category-card" data-cat="${c.id}"><div class="num">${all.length}</div><div class="name">${esc(c.name)}</div><div class="monitor-note">подтверждено: <b>${confirmed}</b>${pending?` · на рассмотрении: <b>${pending}</b>`:''}</div></div>`}).join('');
  target.querySelectorAll('[data-cat]').forEach(x=>x.onclick=()=>{S.issueFilter=x.dataset.cat;goPage('monitoring');renderMonitoring();});
}

function renderRegistryFilters(){
  const opts=(id,vals)=>{const el=$(id),cur=el.value; el.innerHTML='<option value="">Все</option>'+uniq(vals).sort((a,b)=>String(a).localeCompare(String(b),'ru')).map(v=>`<option>${esc(v)}</option>`).join('');el.value=cur;};
  opts('#fDepartment',S.registry.map(r=>r.department)); opts('#fResponsible',S.registry.map(r=>r.responsible)); opts('#fStatus',S.registry.map(r=>statusDerived(r))); opts('#fType',S.registry.map(r=>r.type)); opts('#fYear',S.registry.map(r=>r.year));
  ['#fSearch','#fDepartment','#fResponsible','#fStatus','#fType','#fYear'].forEach(id=>{const el=$(id); if(!el.dataset.bound){el.addEventListener('input',renderRegistry);el.dataset.bound='1';}});
}
function statusDerived(r){ if(/не актуал|утрат/i.test(r.status+' '+r.change))return 'Неактуальный / изменён'; return r.status||'Не указан'; }
function filteredRegistry(){ const q=norm($('#fSearch')?.value),dep=$('#fDepartment')?.value||'',resp=$('#fResponsible')?.value||'',st=$('#fStatus')?.value||'',tp=$('#fType')?.value||'',yr=$('#fYear')?.value||'';
  return S.registry.filter(r=>(!q||norm([r.vndId,r.number,r.title,r.responsible,r.department].join(' ')).includes(q))&&(!dep||r.department===dep)&&(!resp||r.responsible===resp)&&(!st||statusDerived(r)===st)&&(!tp||r.type===tp)&&(!yr||r.year===yr)); }
function renderRegistry(){ const rows=filteredRegistry();
  $('#registryTable').innerHTML=`<div class="table-wrap"><table class="data-table"><thead><tr><th>ID</th><th>ВНД</th><th>№ / дата</th><th>Вид</th><th>Структурное подразделение</th><th>Ответственный</th><th>Статус</th><th>Анализ</th></tr></thead><tbody>${rows.map(r=>{const p=packageFor(r.vndId),n=issuesFor(r.vndId).length;return `<tr class="clickable" data-vnd="${r.vndId}"><td class="mono">${esc(r.vndId)}${r.idTemporary?'<br><span class="badge warn">временный</span>':''}</td><td><div class="truncate" title="${esc(r.title)}">${esc(r.title)}</div></td><td>${esc(r.number)}<br><small>${esc(fmtDate(r.date))}</small></td><td>${esc(r.type)}</td><td>${esc(r.department)}</td><td>${esc(r.responsible)}</td><td><span class="badge ${/не актуал/i.test(statusDerived(r))?'danger':'dark'}">${esc(statusDerived(r))}</span></td><td>${p?`<span class="badge ${p.parsed?'ok':'warn'}">${p.importedAnalysis?'отчёт загружен':(p.parsed?'проверен':'загружен')}</span> ${n?`<span class="badge warn">${n}</span>`:''}`:'<span class="badge">не загружен</span>'}</td></tr>`}).join('')}</tbody></table></div><div class="footer-note">Показано ${rows.length} из ${S.registry.length} записей.</div>`;
  $('#registryTable').querySelectorAll('[data-vnd]').forEach(x=>x.onclick=()=>openVnd(x.dataset.vnd,'overview'));
}

function setupUpload(){
  const dz=$('#dropZone'),inp=$('#fileInput');
  const runSource=async files=>{
    const list=[...(files||[])]; if(!list.length)return;
    $('#ingestProgress')?.classList.remove('hidden'); if($('#ingestText'))$('#ingestText').textContent=`Получено файлов: ${list.length}. Проверка и добавление в сессию…`;
    try{await ingestFiles(list);}catch(e){console.error('Source upload failed',e);S.ingestErrors.push({level:'error',type:'Загрузка ВНД',message:e.message||String(e)});buildQuality();renderAll();toast(`Ошибка загрузки: ${e.message||e}`,'error');}
    finally{if(inp)inp.value='';}
  };
  const runReports=async files=>{const list=[...(files||[])];if(!list.length)return;try{await ingestAnalysisReports(list);}catch(e){console.error('Report upload failed',e);S.ingestErrors.push({level:'error',type:'Загрузка отчёта ВНД',message:e.message||String(e)});buildQuality();renderAll();toast(`Ошибка загрузки отчёта: ${e.message||e}`,'error');}finally{const ri=$('#analysisReportInput');if(ri)ri.value='';}};
  if(dz&&inp){['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('drag')}));['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('drag')}));dz.addEventListener('drop',e=>runSource(e.dataTransfer.files));inp.addEventListener('change',e=>runSource(e.target.files));}
  const runBtn=$('#runAnalysisBtn'),ocrBtn=$('#runPendingOcrBtn');if(runBtn)runBtn.onclick=processAllPackages;if(ocrBtn)ocrBtn.onclick=runPendingOcr;
  const rdz=$('#reportDropZone'),rinp=$('#analysisReportInput');
  if(rdz&&rinp){['dragenter','dragover'].forEach(ev=>rdz.addEventListener(ev,e=>{e.preventDefault();rdz.classList.add('drag')}));['dragleave','drop'].forEach(ev=>rdz.addEventListener(ev,e=>{e.preventDefault();rdz.classList.remove('drag')}));rdz.addEventListener('drop',e=>runReports(e.dataTransfer.files));rinp.addEventListener('change',e=>runReports(e.target.files));}
}

function stableIssueKey(i){return [i.categoryId||'',i.vndId||'',i.page||'',i.point||'',norm(i.issue||''),i.otherVndId||''].join('|');}
async function sha256Text(text){const data=new TextEncoder().encode(String(text||''));const dig=await crypto.subtle.digest('SHA-256',data);return [...new Uint8Array(dig)].map(b=>b.toString(16).padStart(2,'0')).join('');}
function compactUnit(u){return {vndId:u.vndId||'',sourceFile:u.sourceFile||'',page:u.page??null,point:u.point||'',text:u.text||''};}
async function buildNormatrixReport(p){
  if(!p?.parsed)throw new Error('ВНД ещё не проанализирован');
  const vndId=p.match?.vndId||p.vndId||''; if(!vndId)throw new Error('Для ВНД не определён VND_ID');
  const reg=S.registry.find(r=>r.vndId===vndId)||p.match||{};
  if(reg?.idTemporary)throw new Error('Для формирования переносимого отчёта .normatrix сначала присвойте этому ВНД постоянный VND_ID в vnd_master.xlsx.');
  const units=(p.units||[]).map(compactUnit);
  const signature=await sha256Text(vndId+'|'+units.map(u=>[u.sourceFile,u.page,u.point,u.text].join('|')).join('\n'));
  const reviewDecisions={},reviewData={};for(const i of S.issues.filter(x=>x.vndId===vndId||x.otherVndId===vndId)){const k=stableIssueKey(i);if(i.reviewStatus&&i.reviewStatus!=='ai')reviewDecisions[k]=i.reviewStatus;if((i.reviewStatus&&i.reviewStatus!=='ai')||i.measureNote||i.updateNote){reviewData[k]={reviewStatus:i.reviewStatus||'ai',measureNote:i.measureNote||'',updateNote:i.updateNote||''};}}
  const sourceSummary=(p.files||[]).map(f=>({name:f.name,e:f.e,role:f.role,lang:f.lang,parseStatus:f.parseStatus,semanticQuality:f.semanticQuality??null,ocrPendingPages:(f.ocrPendingPages||[]).length,semanticExcluded:!!f.semanticExcluded}));
  const pendingOcr=sourceSummary.reduce((a,f)=>a+(f.ocrPendingPages||0),0)+(p.files||[]).filter(f=>f.ocrPendingImage).length;
  const parseErrors=sourceSummary.filter(f=>['error','doc-unreadable'].includes(f.parseStatus)).length;
  const completeness=(pendingOcr||parseErrors)?'partial':'complete';
  return {format:'NORMATRIX_VND_ANALYSIS',schemaVersion:'1.0',generatedAt:nowIso(),portalVersion:S.version?.portalVersion||S.version?.version||'1.2.3',analysisRulesVersion:S.version?.analysisRulesVersion||'1.2.3',methodology:S.version?.methodology||'NORMATRIX VND Analytics',vndId,signature,completeness,qualityFlags:{pendingOcr,parseErrors,semanticExcluded:sourceSummary.filter(f=>f.semanticExcluded).length},registrySnapshot:{title:reg.title||'',number:reg.number||'',date:reg.date||'',type:reg.type||'',department:reg.department||'',responsible:reg.responsible||'',status:reg.status||'',change:reg.change||''},sourceSummary,units,reviewDecisions,reviewData};
}
async function reportToNormatrixData(report){const z=new JSZip();z.file('analysis.json',JSON.stringify(report));z.file('README.txt','NORMATRIX machine analysis package. Intended for import into NORMATRIX VND Analytics; it does not contain the original VND files.');return z.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});}
async function downloadNormatrixForPackage(p){try{const r=await buildNormatrixReport(p);const b=await reportToNormatrixData(r);downloadBlob(new Blob([b],{type:'application/octet-stream'}),`${r.vndId}.normatrix`);}catch(e){toast(e.message,'warn');}}
async function exportNormatrixBundle(){
  const list=S.packages.filter(p=>p.parsed); if(!list.length){toast('Нет проанализированных ВНД для выгрузки','warn');return;}
  const outer=new JSZip(),manifest=[];let n=0;
  for(const p of list){try{const r=await buildNormatrixReport(p);const b=await reportToNormatrixData(r);outer.file(`${r.vndId}.normatrix`,b);manifest.push({vndId:r.vndId,signature:r.signature,title:r.registrySnapshot.title});n++;}catch(e){S.ingestErrors.push({level:'warn',type:'Экспорт .normatrix',message:`${p.containerName}: ${e.message}`});}}
  outer.file('bundle-manifest.json',JSON.stringify({format:'NORMATRIX_ANALYSIS_BUNDLE',schemaVersion:'1.0',generatedAt:nowIso(),count:n,reports:manifest},null,2));
  const blob=await outer.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}});downloadBlob(blob,`NORMATRIX_REPORTS_${new Date().toISOString().slice(0,10)}.zip`);toast(`Сформировано машинных отчётов: ${n}`,'ok');
}
async function parseNormatrixFile(file){
  const ab=await file.arrayBuffer();let report=null;
  try{const zip=await JSZip.loadAsync(ab);const entry=zip.file('analysis.json')||Object.values(zip.files).find(x=>!x.dir&&/analysis\.json$/i.test(x.name));if(!entry)throw new Error('analysis.json не найден');report=JSON.parse(await entry.async('text'));}
  catch(e){try{report=JSON.parse(new TextDecoder().decode(ab));}catch(_){throw new Error('файл не является корректным .normatrix');}}
  if(report?.format!=='NORMATRIX_VND_ANALYSIS'||!report.vndId||!Array.isArray(report.units))throw new Error('неподдерживаемая структура отчёта');
  if(String(report.schemaVersion||'').split('.')[0]!=='1')throw new Error(`несовместимая версия схемы отчёта: ${report.schemaVersion||'не указана'}`);
  return report;
}
async function extractNormatrixReportsFromUpload(file){
  if(file.name.toLowerCase().endsWith('.normatrix'))return [await parseNormatrixFile(file)];
  if(ext(file.name)==='zip'){
    const z=await loadZipRobust(file);const entries=Object.values(z.files).filter(x=>!x.dir&&x.name.toLowerCase().endsWith('.normatrix'));const out=[];
    for(const e of entries){const b=await e.async('blob');out.push(await parseNormatrixFile(new File([b],e.name,{type:'application/octet-stream'})));}return out;
  }
  throw new Error('поддерживаются .normatrix и ZIP-пакеты отчётов');
}
function sanitizeImportedText(text){let t=String(text||'');let cut=t.length;const rx=[/([A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі])\1{12,}/u,/Microsoft Photo Editor/iu,/VГИ\[СЭ/iu];for(const r of rx){const m=r.exec(t);if(m&&m.index>20)cut=Math.min(cut,m.index);}const changed=cut<t.length;if(changed)t=t.slice(0,cut).trim();const q=semanticTextQuality(t);const usable=changed?(t.length>=80&&q>=.12):(t.length<120||q>=.12);return{text:t,changed,usable,quality:q};}
function reportPackageFromData(r,fileName){
  const rec=S.registry.find(x=>x.vndId===r.vndId)||null;const key=`RPT-${r.vndId}-${Math.random().toString(36).slice(2,8)}`;
  let excluded=0,sanitized=0;const clean=[];for(const u of (r.units||[])){const z=sanitizeImportedText(u.text||'');if(!z.usable){excluded++;continue;}if(z.changed)sanitized++;clean.push({...u,text:z.text});}
  if(excluded)r.importExcludedUnits=excluded;if(sanitized)r.importSanitizedUnits=sanitized;
  const units=clean.map((u,idx)=>({id:`U-${key}-${idx}`,vndId:r.vndId,packageKey:key,sourceFile:u.sourceFile||`${r.vndId}.normatrix`,page:u.page??null,point:u.point||'',text:u.text||'',normText:norm(u.text||''),tokens:uniq(words(u.text||'')),file:null}));
  return {key,containerName:fileName||`${r.vndId}.normatrix`,isZip:false,vndId:r.vndId,files:[],match:rec,matchConfidence:rec?1:0,parsed:true,units,analysisStatus:'отчёт ВНД загружен',errors:[],primaryFiles:[],primaryConfidence:0,primarySource:'analysis-report',importedAnalysis:true,reportSignature:r.signature||'',reportMeta:r};
}

function dedupeSessionPackages(){
  const groups=new Map();for(const p of S.packages){const id=p.match?.vndId||p.vndId||'';if(!id)continue;if(!groups.has(id))groups.set(id,[]);groups.get(id).push(p);}
  const remove=new Set();for(const [id,arr] of groups){if(arr.length<2)continue;const source=arr.filter(p=>!p.importedAnalysis),reports=arr.filter(p=>p.importedAnalysis);const keep=source[0]||arr[0];
    if(source.length>1)S.ingestErrors.push({level:'error',type:'Конфликт пакетов ВНД',message:`${id}: загружено ${source.length} исходных пакета. Для аналитики временно оставлен первый; необходимо проверить, какой пакет является правильным.`});
    if(!source.length&&reports.length>1){const sigs=uniq(reports.map(p=>p.reportSignature).filter(Boolean));if(sigs.length>1)S.ingestErrors.push({level:'error',type:'Конфликт отчётов ВНД',message:`${id}: загружены машинные отчёты с разным содержанием/подписью. Для аналитики временно оставлен первый; требуется выбрать корректный отчёт.`});}
    for(const p of arr){if(p===keep)continue;remove.add(p);for(const f of p.files||[]){if(f.objectUrl){URL.revokeObjectURL(f.objectUrl);S.objectUrls.delete(f.objectUrl);}}}
    S.ingestErrors.push({level:'warn',type:'Дедупликация ВНД',message:`${id}: обнаружено ${arr.length} представления; повторные данные не суммируются. Для текущей аналитики используется одно представление (${keep.importedAnalysis?'машинный отчёт':'исходный ВНД'}).`});}
  if(remove.size)S.packages=S.packages.filter(p=>!remove.has(p));
}
async function ingestAnalysisReports(files){
  if(!files.length)return;let added=0,skipped=0;const max=Number(S.cfg?.limits?.maxVndPackages||300);
  for(const file of files){try{const reports=await extractNormatrixReportsFromUpload(file);for(const r of reports){if(S.packages.length>=max){S.ingestErrors.push({level:'warn',type:'Лимит отчётов ВНД',message:`Достигнут лимит ${max} уникальных ВНД за сессию.`});break;}const existing=S.packages.find(p=>(p.match?.vndId||p.vndId)===r.vndId);if(existing){skipped++;const different=existing.importedAnalysis&&existing.reportSignature&&r.signature&&existing.reportSignature!==r.signature;S.ingestErrors.push({level:different?'error':'warn',type:different?'Конфликт отчётов ВНД':'Дубликат отчёта ВНД',message:different?`${r.vndId}: уже загружен другой машинный отчёт с отличающейся подписью. Второй отчёт не учтён; требуется проверить правильный вариант.`:`${r.vndId}: уже присутствует в текущей сессии; повторный отчёт не учтён.`});continue;}const p=reportPackageFromData(r,file.name);S.packages.push(p);for(const [k,v] of Object.entries(r.reviewDecisions||{}))S.importedReviewDecisions.set(k,v);for(const [k,v] of Object.entries(r.reviewData||{}))S.importedReviewData.set(k,v);added++;}}
    catch(e){S.ingestErrors.push({level:'error',type:'Отчёт ВНД',message:`${file.name}: ${e.message}`});}}
  refreshMatches();dedupeSessionPackages();if(S.staticReady)rebuildAnalysisFromParsedFiles();buildQuality();renderAll();updateRunAnalysisButton();toast(`Отчёты ВНД: добавлено ${added}${skipped?`, дубликатов исключено ${skipped}`:''}${!S.staticReady?', сводный анализ будет построен после загрузки справочников':''}`,'ok');
}
function applyImportedReviewDecisions(){for(const i of S.issues){const k=stableIssueKey(i),v=S.importedReviewDecisions.get(k),d=S.importedReviewData.get(k);if(v)i.reviewStatus=v;if(d){if(d.reviewStatus)i.reviewStatus=d.reviewStatus;i.measureNote=d.measureNote||'';i.updateNote=d.updateNote||'';}}}

function roleFromName(name){ const n=name.toUpperCase(); if(/DRAFT|ПРОЕКТ/.test(n))return'DRAFT'; if(/ORDER|ПРИКАЗ|РЕШЕНИЕ|ПРОТОКОЛ/.test(n))return'ORDER'; if(/AMD\d*|ИЗМЕН|ДОПОЛНЕН/.test(n))return'AMD'; if(/APP\d*|ПРИЛОЖ/.test(n))return'APP'; if(/MAIN|ОСНОВ/.test(n))return'MAIN'; return'OTHER'; }
function langFromName(name){ const n=name.toUpperCase(); if(/(?:^|[_-])KZ(?:[_-]|\.)/.test(n))return'KZ';if(/(?:^|[_-])RU(?:[_-]|\.)/.test(n))return'RU';return''; }
function roleLabel(role){return ({MAIN:'основной ВНД',ORDER:'акт утверждения',APP:'приложение',AMD:'изменение',DRAFT:'проект',OTHER:'не классифицирован'})[role]||role;}
function inferRoleFromContent(f){
  if(f.role!=='OTHER')return f.role;
  const sample=norm((f.text||'').slice(0,2600)); if(!sample)return f.role;
  if(/приказ/u.test(sample.slice(0,700)) || /решение/u.test(sample.slice(0,500))) return 'ORDER';
  if(/приложение/u.test(sample.slice(0,500))) return 'APP';
  if(/о внесении измен|внести измен|изменения и дополнения/u.test(sample.slice(0,900))) return 'AMD';
  if(/проект/u.test(sample.slice(0,250))) return 'DRAFT';
  return f.role;
}
function primaryScore(f,p,contentAware=false){
  let s=0; const fn=norm(f.name), rec=p.match, role=f.role;
  if(role==='MAIN')s+=100; if(role==='DRAFT')s-=8; if(role==='ORDER')s-=7; if(role==='APP')s-=5; if(role==='AMD')s-=4;
  if(f.e==='pdf')s+=3; else if(f.e==='docx')s+=2.6; else if(['jpg','jpeg','png','tif','tiff','webp'].includes(f.e))s+=1.4; else if(f.e==='txt')s+=.8;
  if(/подпис|signed|утвержден/u.test(fn))s+=1.4;
  if(/положение|правил|регламент|инструкц|политик|методик|порядок|стандарт/u.test(fn))s+=2.2;
  if(rec){
    const tw=words(rec.title), fw=new Set(words(fn)); if(tw.length){const ov=tw.filter(x=>fw.has(x)).length/tw.length;s+=Math.min(5,ov*6);}
    if(rec.type && fn.includes(norm(rec.type)))s+=1.5;
  }
  if(contentAware && f.text){
    const sample=norm(f.text.slice(0,5000));
    if(rec){const tw=words(rec.title), sw=new Set(words(sample)); if(tw.length){const ov=tw.filter(x=>sw.has(x)).length/tw.length;s+=Math.min(10,ov*11);} if(rec.type&&sample.includes(norm(rec.type)))s+=2;}
    if(/приказ/u.test(sample.slice(0,600)) && role==='OTHER')s-=2.5;
    if(/приложение/u.test(sample.slice(0,450)) && role==='OTHER')s-=1.8;
    s+=Math.min(2,Math.log10(Math.max(10,(f.text||'').length))/2);
  }
  return s;
}
function detectPrimaryFiles(p,contentAware=false){
  if(!p.files.length){p.primaryFiles=[];p.primaryConfidence=0;p.primarySource='none';return;}
  if(p.primarySource==='manual' && p.primaryFiles?.length){for(const f of p.files)f.isPrimary=p.primaryFiles.includes(f.name);return;}
  const explicit=p.files.filter(f=>f.role==='MAIN');
  if(explicit.length){p.primaryFiles=explicit.map(f=>f.name);p.primaryConfidence=1;p.primarySource='explicit';for(const f of p.files)f.isPrimary=p.primaryFiles.includes(f.name);return;}
  const scored=p.files.map(f=>({f,score:primaryScore(f,p,contentAware)})).sort((a,b)=>b.score-a.score);
  const top=scored[0], second=scored[1]; if(!top){return;}
  let selected=[top.f.name], source=contentAware?'auto-content':'auto-name';
  // Если пакет фактически состоит из последовательности изображений без служебных ролей, считаем их единым основным сканом.
  const imgs=p.files.filter(f=>['jpg','jpeg','png','tif','tiff','webp'].includes(f.e) && !['ORDER','APP','AMD','DRAFT'].includes(f.role));
  const nonService=p.files.filter(f=>!['ORDER','APP','AMD','DRAFT'].includes(f.role));
  if(imgs.length>=2 && nonService.length===imgs.length && imgs.length/p.files.length>=0.6){selected=imgs.map(f=>f.name);source='auto-image-series';}
  const margin=top.score-(second?.score??-99); let conf=top.score>=9&&margin>=2?0.94:top.score>=6&&margin>=1.2?0.84:top.score>=4?0.72:0.55;
  if(source==='auto-image-series')conf=Math.max(conf,0.78);
  p.primaryFiles=selected;p.primaryConfidence=conf;p.primarySource=source;for(const f of p.files)f.isPrimary=selected.includes(f.name);
}
function setPrimaryManual(p,fileName){if(!p)return;const f=p.files.find(x=>x.name===fileName);if(!f)return;p.primaryFiles=[fileName];p.primaryConfidence=1;p.primarySource='manual';for(const x of p.files)x.isPrimary=x.name===fileName;buildQuality();renderAll();openPackage(p.key);toast(`Основным источником выбран ${fileName}`,'ok');}
function mimeByExt(e){ return ({pdf:'application/pdf',docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',jpg:'image/jpeg',jpeg:'image/jpeg',png:'image/png',tif:'image/tiff',tiff:'image/tiff',webp:'image/webp',txt:'text/plain'})[e]||'application/octet-stream'; }
function zipNameScore(s){
  s=String(s||''); if(!s)return -999;
  const bad=(s.match(/\uFFFD/g)||[]).length+(s.match(/[\u0000-\u001F]/g)||[]).length*3;
  const cyr=(s.match(/[А-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]/g)||[]).length;
  const moj=(s.match(/[Џ®«¦ҐЁЈЎ©¬]/g)||[]).length;
  const ascii=(s.match(/[A-Za-z0-9._\-/ ]/g)||[]).length;
  return cyr*2+ascii*.08-moj*3-bad*8;
}
function decodeZipFileName(bytes){
  const arr=bytes instanceof Uint8Array?bytes:new Uint8Array(bytes||[]);
  const variants=[];
  for(const enc of ['utf-8','ibm866','windows-1251']){
    try{variants.push(new TextDecoder(enc,{fatal:enc==='utf-8'}).decode(arr));}catch(_){ }
  }
  if(!variants.length) return Array.from(arr,b=>String.fromCharCode(b)).join('');
  variants.sort((a,b)=>zipNameScore(b)-zipNameScore(a));
  return variants[0];
}
async function loadZipRobust(source){
  if(!window.JSZip)throw new Error('ZIP-модуль JSZip не загружен. Проверьте наличие vendor/jszip.min.js.');
  const data=source instanceof ArrayBuffer?source:source?.arrayBuffer?await source.arrayBuffer():source;
  return JSZip.loadAsync(data,{decodeFileName:decodeZipFileName,createFolders:true});
}
function cleanZipPath(name){
  return String(name||'').replace(/\\/g,'/').replace(/^\.\//,'').replace(/^\/+/, '');
}
async function ingestFiles(files){
  if(!files?.length)return;
  const cfg=S.cfg||DEFAULT_CFG, limits={...DEFAULT_CFG.limits,...(cfg.limits||{})}, analysis={...DEFAULT_CFG.analysis,...(cfg.analysis||{})};
  const max=Number(limits.maxVndPackages||300);
  if(S.packages.length+files.length>max){toast(`Максимум ${max} объектов загрузки за сессию`,'warn');return;}
  // Базовая загрузка не зависит от package_schema.json: допустимые форматы встроены в приложение.
  const allowed=new Set(['zip',...DEFAULT_ALLOWED_EXTENSIONS]);
  const accepted=files.filter(f=>allowed.has(ext(f.name))), bad=files.filter(f=>!allowed.has(ext(f.name)));
  if(bad.length){S.ingestErrors.push({level:'warn',type:'Формат файла',message:`Пропущены неподдерживаемые: ${bad.map(x=>x.name).join(', ')}`});toast(`Пропущены неподдерживаемые файлы: ${bad.map(x=>x.name).join(', ')}`,'warn');}
  if(!accepted.length){buildQuality();renderAll();return;}
  const existingBytes=S.packages.flatMap(p=>p.files||[]).reduce((a,f)=>a+(f.size||0),0), incomingBytes=accepted.reduce((a,f)=>a+(f.size||0),0);
  if(bytesMB(existingBytes+incomingBytes)>Number(limits.maxSessionMB||2048)){toast(`Общий объём сессии превысит ${limits.maxSessionMB} MB`,'warn');return;}
  if(S.packages.reduce((a,p)=>a+(p.files?.length||0),0)+accepted.length>Number(limits.maxPhysicalFiles||5000)){toast(`Превышен лимит ${limits.maxPhysicalFiles} физических файлов за сессию`,'warn');return;}
  let addedBefore=S.packages.length;
  for(const file of accepted){
    try{
      if(bytesMB(file.size)>Number(limits.maxSingleFileMB||100)){S.ingestErrors.push({level:'error',type:'Размер файла',message:`${file.name}: превышен лимит ${limits.maxSingleFileMB} MB`});continue;}
      if(ext(file.name)==='zip') await ingestZip(file); else await addPackage(file.name,[makePhysical(file,file.name)]);
    }catch(e){console.error('File ingest failed',file.name,e);S.ingestErrors.push({level:'error',type:'Загрузка файла',message:`${file.name}: ${e.message||e}`});toast(`${file.name}: ${e.message||e}`,'error');}
  }
  refreshMatches();dedupeSessionPackages();buildQuality();renderAll();
  const added=Math.max(0,S.packages.length-addedBefore);
  if($('#ingestText'))$('#ingestText').textContent=added?`Добавлено в сессию: ${added}. Можно запускать анализ.`:'Новые ВНД не добавлены. Откройте «Качество данных» для причины.';
  updateRunAnalysisButton();
  if(added)toast(`Добавлено ВНД в сессию: ${added}`,'ok');
  if(analysis.autoRunAfterIngest&&added) processAllPackages();
}

function makePhysical(blob,name){ const e=ext(name); const file=blob instanceof File?blob:new File([blob],name,{type:mimeByExt(e)}); const url=URL.createObjectURL(file);S.objectUrls.add(url);return {name,e,blob:file,size:file.size,role:roleFromName(name),lang:langFromName(name),objectUrl:url,text:'',pages:[],html:'',parseStatus:'pending',parseError:''}; }
async function ingestZip(file){
  try{
    const zip=await loadZipRobust(file); const entries=[]; let declaredBytes=0;
    const supported=new Set(DEFAULT_ALLOWED_EXTENSIONS);
    const limits={...DEFAULT_CFG.limits,...((S.cfg||{}).limits||{})};
    const currentPhysical=S.packages.reduce((a,p)=>a+p.files.length,0);
    const currentBytes=S.packages.flatMap(p=>p.files).reduce((a,f)=>a+(f.size||0),0);
    const skipped=[];
    for(const [rawName,z] of Object.entries(zip.files)){
      if(z.dir)continue;
      const name=cleanZipPath(rawName);
      if(!name || /(?:^|\/)__MACOSX\//i.test(name) || /(?:^|\/)\.DS_Store$/i.test(name))continue;
      const e=ext(name); if(!supported.has(e)){skipped.push(name);continue;}
      if(entries.length+currentPhysical>=limits.maxPhysicalFiles) throw new Error(`превышен лимит ${limits.maxPhysicalFiles} физических файлов`);
      const declared=Number(z?._data?.uncompressedSize||0); declaredBytes+=declared;
      if(declared && bytesMB(currentBytes+declaredBytes)>limits.maxSessionMB) throw new Error(`распакованный объём превысит ${limits.maxSessionMB} MB`);
      const blob=await z.async('blob');
      if(bytesMB(blob.size)>limits.maxSingleFileMB) throw new Error(`${name}: после распаковки превышен лимит ${limits.maxSingleFileMB} MB`);
      entries.push(makePhysical(blob,name));
    }
    if(!entries.length)throw new Error(`в архиве не найдено поддерживаемых документов. Поддерживаются: ${[...supported].join(', ')}`);
    await addPackage(file.name,entries,true);
    if(skipped.length)S.ingestErrors.push({level:'warn',type:'ZIP: пропущенные файлы',message:`${file.name}: пропущено неподдерживаемых файлов ${skipped.length}`});
  }catch(e){
    console.error('ZIP ingest failed',file?.name,e);
    S.ingestErrors.push({level:'error',type:'ZIP',message:`${file.name}: ${e.message||e}`});
    toast(`Не удалось загрузить ZIP ${file.name}: ${e.message||e}`,'error');
  }
}
async function addPackage(containerName,physicalFiles,isZip=false){
  const m=containerName.match(/VND-\d{6}/i); const p={key:`PKG-${Date.now()}-${Math.random().toString(36).slice(2,8)}`,containerName,isZip,vndId:m?m[0].toUpperCase():'',files:physicalFiles,match:null,matchConfidence:0,parsed:false,units:[],analysisStatus:'loaded',errors:[],primaryFiles:[],primaryConfidence:0,primarySource:'none'}; S.packages.push(p);
}
function refreshMatches(){
  for(const p of S.packages){
    p.match=null;p.matchConfidence=0;
    if(p.vndId){ const r=S.registry.find(x=>x.vndId===p.vndId); if(r){p.match=r;p.matchConfidence=1;detectPrimaryFiles(p,false);continue;} }
    const fn=norm(p.containerName); let best=null,score=0; for(const r of S.registry){ let s=0; if(r.number&&fn.includes(norm(r.number)))s+=.55; const wt=words(r.title),wf=new Set(words(fn)); if(wt.length){const ov=wt.filter(x=>wf.has(x)).length/wt.length;s+=Math.min(.45,ov*.65);} if(s>score){score=s;best=r;} } if(score>=.6){p.match=best;p.matchConfidence=Math.min(score,0.99);} detectPrimaryFiles(p,false);
  }
}
function renderPackages(){
  const phys=S.packages.reduce((a,p)=>a+p.files.length,0); $('#uploadSummary').textContent=`${S.packages.length} ВНД / ${phys} физических файлов`;
  if(!S.packages.length){$('#packageTable').innerHTML='<div class="empty"><strong>Пакеты не загружены</strong>Используйте VND-ID в имени ZIP для однозначного сопоставления.</div>';return;}
  $('#packageTable').innerHTML=`<div class="table-wrap"><table class="data-table"><thead><tr><th>Пакет</th><th>Связь с реестром</th><th>Состав / основной источник</th><th>Статус</th><th>Проблемы</th></tr></thead><tbody>${S.packages.map(p=>`<tr class="clickable" data-pkg="${p.key}"><td><b>${esc(p.containerName)}</b><br><span class="mono">${esc(p.vndId||'ID не указан')}</span></td><td>${p.match?`<button class="link-btn" data-vnd="${esc(p.match.vndId)}">${esc(p.match.title)}</button><br><small>${esc(p.match.vndId)}</small><br><span class="badge ${p.matchConfidence>.94?'ok':'warn'}">${Math.round(p.matchConfidence*100)}%</span>`:'<span class="badge danger">не сопоставлен</span>'}</td><td>${p.importedAnalysis?`<span class="badge info">машинный отчёт .normatrix</span><br><small>${p.units?.length||0} структурированных норм; исходный файл не загружен</small>`:`${p.files.length} файлов<br><small>${esc(uniq(p.files.map(f=>roleLabel(f.role))).join(', '))}</small><br>${p.primaryFiles?.length?`<span class="badge ${p.primaryConfidence>=.8?'ok':'warn'}">основной: ${esc(p.primaryFiles.length===1?p.primaryFiles[0]:`${p.primaryFiles.length} страниц/файлов`)}</span> <small>${Math.round(p.primaryConfidence*100)}%</small>`:'<span class="badge warn">основной не определён</span>'}` }</td><td><span class="badge ${p.parsed?'ok':p.analysisStatus==='processing'?'info':'warn'}">${esc(p.importedAnalysis?'отчёт ВНД загружен':(p.parsed?'анализ завершён':p.analysisStatus))}</span></td><td>${p.errors.length?`<span class="badge danger">${p.errors.length}</span>`:'—'}</td></tr>`).join('')}</tbody></table></div>`;
  $('#packageTable').querySelectorAll('[data-pkg]').forEach(x=>x.onclick=e=>{if(e.target.closest('[data-vnd]'))return;openPackage(x.dataset.pkg);});$('#packageTable').querySelectorAll('[data-vnd]').forEach(x=>x.onclick=e=>{e.stopPropagation();openVnd(x.dataset.vnd,'overview');});
}

function canReuseParsedFile(f){
  // В пределах текущей ZERO-PERSISTENCE сессии уже извлечённый текст/OCR является
  // рабочим кэшем в памяти. Повторный запуск анализа не должен заново читать
  // тот же PDF и снова ставить успешно распознанные страницы в OCR-очередь.
  return Array.isArray(f.pages) && f.pages.length>0 &&
    ['done','ocr-pending','legacy-doc-partial','doc-unreadable'].includes(f.parseStatus);
}
async function processAllPackages(){
  if(S.processing||!S.packages.length)return;if(!S.staticReady){toast('Справочники и правила ещё загружаются. Файл уже принят в сессию; анализ станет доступен после инициализации.','warn');return;} const sourcePkgs=S.packages.filter(p=>!p.importedAnalysis); if(!sourcePkgs.length){rebuildAnalysisFromParsedFiles();renderAll();goPage('dashboard');toast(`Суммарный анализ построен по ${S.packages.length} отчётам ВНД`,'ok');return;}
  S.processing=true; $('#runAnalysisBtn').disabled=true; $('#ingestProgress').classList.remove('hidden');
  S.issues=[];S.legalRefs=[];S.duplicates=[];S.conflicts=[];S.recommendations=[];S.ocrCount=0;
  let done=0; const total=sourcePkgs.length;
  for(const p of sourcePkgs){ p.analysisStatus='processing'; p.units=[]; p.errors=[]; renderPackages();
    for(const f of p.files){ try{
      if(!canReuseParsedFile(f)) await parsePhysical(f);
      const inferred=inferRoleFromContent(f); if(f.role==='OTHER'&&inferred!=='OTHER')f.role=inferred; p.units.push(...unitsFromFile(f,p));
    }catch(e){f.parseStatus='error';f.parseError=e.message;p.errors.push(`${f.name}: ${e.message}`);} }
    detectPrimaryFiles(p,true); p.parsed=true;p.analysisStatus='done';done++; updateProgress(done,total,`Обработан ${p.containerName}`);
  }
  runLocalAnalysis(); buildRecommendations(); buildQuality(); S.processing=false; $('#runAnalysisBtn').disabled=false; updateOcrButton(); renderAll(); goPage('dashboard'); toast(`Анализ завершён: ${S.packages.length} уникальных ВНД, ${S.issues.length} замечаний-кандидатов`,'ok');
}
function updateProgress(done,total,text){ const pct=Math.round(done/total*100);$('#ingestBar').style.width=`${pct}%`;$('#ingestText').textContent=`${pct}% — ${text}`; }
async function parsePhysical(f){
  f.parseStatus='processing'; const e=f.e;
  if(e==='pdf'){if(!window.pdfjsLib)throw new Error('PDF.js не загружен: PDF добавлен в сессию, но анализ PDF недоступен.');await parsePdf(f);}
  else if(e==='docx'){if(!window.mammoth)throw new Error('Mammoth не загружен: DOCX добавлен в сессию, но анализ DOCX недоступен.');await parseDocx(f);}
  else if(e==='xlsx'||e==='xls'){if(!window.XLSX)throw new Error('SheetJS/XLSX не загружен: Excel добавлен в сессию, но анализ Excel недоступен.');await parseSheet(f);}
  else if(e==='txt') await parseTxt(f);
  else if(['jpg','jpeg','png','tif','tiff','webp'].includes(e)){if(!window.Tesseract)throw new Error('Tesseract OCR не загружен: изображение добавлено в сессию, но OCR недоступен.');await parseImage(f);}
  else if(e==='doc') await parseLegacyDoc(f);
  if(f.parseStatus==='processing') f.parseStatus='done';
}
async function parsePdf(f){
  const ab=await f.blob.arrayBuffer(); const pdf=await pdfjsLib.getDocument({data:ab}).promise; const pages=[]; const pending=[];
  for(let i=1;i<=pdf.numPages;i++){
    const pg=await pdf.getPage(i),tc=await pg.getTextContent();
    const items=(tc.items||[]).filter(x=>x.str&&x.str.trim()).map(x=>({str:x.str,x:Number(x.transform?.[4]||0),y:Number(x.transform?.[5]||0)}));
    items.sort((a,b)=>Math.abs(b.y-a.y)>2?b.y-a.y:a.x-b.x);
    const lines=[]; let cur=[],lastY=null;
    for(const it of items){ if(lastY===null||Math.abs(it.y-lastY)<=2){cur.push(it);}else{cur.sort((a,b)=>a.x-b.x);lines.push(cur.map(z=>z.str).join(' ').replace(/\s+/g,' ').trim());cur=[it];} lastY=it.y; }
    if(cur.length){cur.sort((a,b)=>a.x-b.x);lines.push(cur.map(z=>z.str).join(' ').replace(/\s+/g,' ').trim());}
    let tx=lines.filter(Boolean).join('\n'),ocrConfidence=null;
    if(needsOcrText(tx) && S.cfg.analysis.autoOcrImages){
      if(S.ocrCount<S.cfg.limits.maxAutoOcrImages && window.Tesseract){
        try{ S.ocrCount++; const r=await ocrPdfPage(pg); tx=r.text; ocrConfidence=r.confidence; }
        catch(e){ pending.push(i); }
      }else pending.push(i);
    }
    pages.push({number:i,text:tx,ocrConfidence,ocrPending:pending.includes(i)});
  }
  f.pages=pages; f.ocrPendingPages=pending; f.text=pages.map(p=>p.text).join('\n'); if(pending.length)f.parseStatus='ocr-pending';
}
function textQuality(text){
  const t=String(text||'').replace(/\s+/g,' ').trim();
  const compact=t.replace(/\s/g,'');
  if(!compact) return {length:0,letters:0,letterRatio:0,badRatio:1};
  const letters=(compact.match(/[A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]/g)||[]).length;
  const bad=(compact.match(/[�□■◆◇]|[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g)||[]).length;
  return {length:compact.length,letters,letterRatio:letters/compact.length,badRatio:bad/compact.length};
}
function needsOcrText(text){
  const q=textQuality(text);
  if(q.length===0) return true;
  if(q.length<20) return q.letters<5 || q.letterRatio<.45;
  if(q.badRatio>.08) return true;
  if(q.letterRatio<.28) return true;
  return false;
}

async function recognizeOcr(source){
  if(!window.Tesseract) throw new Error('OCR-библиотека недоступна');
  try{const r=await Tesseract.recognize(source,'rus+kaz',{logger:()=>{}});return {text:r.data.text||'',confidence:r.data.confidence||0};}
  catch(e){const r=await Tesseract.recognize(source,'rus',{logger:()=>{}});return {text:r.data.text||'',confidence:r.data.confidence||0};}
}
async function ocrPdfPage(pg){
  const vp=pg.getViewport({scale:1.75}),canvas=document.createElement('canvas'); canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);
  const ctx=canvas.getContext('2d',{willReadFrequently:true}); await pg.render({canvasContext:ctx,viewport:vp}).promise; return recognizeOcr(canvas);
}
async function parseDocx(f){ const ab=await f.blob.arrayBuffer(); const [raw,html]=await Promise.all([mammoth.extractRawText({arrayBuffer:ab}),mammoth.convertToHtml({arrayBuffer:ab})]);f.text=raw.value||'';f.html=html.value||'';f.pages=[{number:null,text:f.text}]; }
async function parseSheet(f){ const ab=await f.blob.arrayBuffer(); const wb=XLSX.read(ab,{type:'array'}); const pages=[]; for(const sn of wb.SheetNames){ const csv=XLSX.utils.sheet_to_csv(wb.Sheets[sn]);pages.push({number:null,label:sn,text:csv}); }f.pages=pages;f.text=pages.map(p=>`[${p.label}]\n${p.text}`).join('\n'); }
async function parseTxt(f){f.text=await f.blob.text();f.pages=[{number:null,text:f.text}];}
async function parseImage(f){
  if(!S.cfg.analysis.autoOcrImages || S.ocrCount>=S.cfg.limits.maxAutoOcrImages){f.text='';f.pages=[{number:pageNumFromName(f.name),text:'',ocrPending:true}];f.ocrPendingImage=true;f.parseStatus='ocr-pending';return;}
  S.ocrCount++; const r=await recognizeOcr(f.blob); f.text=r.text; f.pages=[{number:pageNumFromName(f.name),text:f.text,ocrConfidence:r.confidence,ocrPending:false}]; f.ocrPendingImage=false;
}
async function parseLegacyDoc(f){
  const ab=await f.blob.arrayBuffer(), bytes=new Uint8Array(ab); let text='', method='';
  const asciiHead=new TextDecoder('windows-1251',{fatal:false}).decode(bytes.slice(0,64));
  if(/\{\\rtf/i.test(asciiHead)){
    const raw=new TextDecoder('windows-1251',{fatal:false}).decode(bytes); text=stripRtf(raw); method='RTF/DOC';
  }else{
    const a=extractLegacyRuns(bytes,'windows-1251'), u=extractUtf16Runs(bytes); text=mergeLegacyText(a,u); method='binary DOC';
  }
  f.text=text.trim(); f.pages=[{number:null,text:f.text}]; f.legacyDocMethod=method;
  if(f.text.length>=80){f.parseStatus='legacy-doc-partial';}
  else {f.parseStatus='doc-unreadable'; f.parseError='Текст старого DOC не удалось надёжно извлечь локально.';}
}
function stripRtf(raw){
  return raw.replace(/\\par[d]?\b/gi,'\n').replace(/\\tab\b/gi,'\t').replace(/\\'[0-9a-f]{2}/gi,m=>{try{return new TextDecoder('windows-1251').decode(Uint8Array.from([parseInt(m.slice(2),16)]));}catch(e){return ' ';}}).replace(/\\u(-?\d+)\??/gi,(_,n)=>String.fromCharCode(Number(n)<0?Number(n)+65536:Number(n))).replace(/\\[a-z]+-?\d* ?/gi,'').replace(/[{}]/g,' ').replace(/\s+\n/g,'\n').replace(/\n\s+/g,'\n').replace(/[ \t]{2,}/g,' ').trim();
}
function extractLegacyRuns(bytes,encoding){
  const ok=b=>b===9||b===10||b===13||(b>=32&&b<=126)||b===160||b===168||b===184||(b>=192&&b<=255); const out=[];let start=-1;
  const flush=i=>{if(start<0)return;const chunk=bytes.slice(start,i);start=-1;if(chunk.length<16)return;let t='';try{t=new TextDecoder(encoding,{fatal:false}).decode(chunk);}catch(e){return;}t=t.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]+/g,' ').replace(/\s+/g,' ').trim();const letters=(t.match(/[A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]/g)||[]).length;if(t.length>=12&&letters>=6&&letters/Math.max(1,t.length)>.22)out.push(t);};
  for(let i=0;i<=bytes.length;i++){const good=i<bytes.length&&ok(bytes[i]);if(good&&start<0)start=i;if(!good&&start>=0)flush(i);} return out;
}
function extractUtf16Runs(bytes){
  const out=[]; let chars=[]; const useful=c=>c===9||c===10||c===13||c===32||(c>=33&&c<=126)||(c>=0x400&&c<=0x52f)||(c>=0x00a0&&c<=0x00ff);
  const flush=()=>{if(chars.length>=8){let t=String.fromCharCode(...chars).replace(/\s+/g,' ').trim();const letters=(t.match(/[A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]/g)||[]).length;if(t.length>=12&&letters>=6&&letters/Math.max(1,t.length)>.22)out.push(t);}chars=[];};
  for(let i=0;i+1<bytes.length;i+=2){const c=bytes[i]|(bytes[i+1]<<8);if(useful(c))chars.push(c);else flush();}flush();return out;
}
function mergeLegacyText(a,u){
  const all=[...a,...u].map(x=>x.trim()).filter(Boolean), seen=new Set(), out=[]; for(const t of all){const k=norm(t);if(k.length<8||seen.has(k))continue;seen.add(k);out.push(t);} return out.join('\n');
}

function pendingOcrTasks(){
  const tasks=[]; for(const p of S.packages)for(const f of p.files){if(f.e==='pdf'&&f.ocrPendingPages?.length)for(const page of f.ocrPendingPages)tasks.push({p,f,page,type:'pdf'});else if(['jpg','jpeg','png','tif','tiff','webp'].includes(f.e)&&f.ocrPendingImage)tasks.push({p,f,page:f.pages?.[0]?.number||null,type:'image'});} return tasks;
}
function updateOcrButton(){
  const b=$('#runPendingOcrBtn'); if(!b)return; const n=pendingOcrTasks().length; b.disabled=S.processing||n===0; b.textContent=n?`OCR: распознать ожидающие (${n})`:'OCR: нет ожидающих';
}
async function runPendingOcr(){
  if(S.processing)return; const tasks=pendingOcrTasks(), max=Number(S.cfg.limits.maxManualOcrPages||1200); if(!tasks.length){toast('Ожидающих OCR страниц нет','ok');return;}
  if(tasks.length>max){toast(`Ожидает ${tasks.length} страниц; за один запуск разрешено до ${max}. Остальные останутся в очереди.`,'warn');}
  const selected=tasks.slice(0,max); if(selected.length>80&&!confirm(`Запустить OCR для ${selected.length} страниц/изображений? Обработка идёт последовательно и может заметно нагружать компьютер.`))return;
  S.processing=true; $('#runAnalysisBtn').disabled=true; $('#runPendingOcrBtn').disabled=true; $('#ingestProgress').classList.remove('hidden'); let done=0;
  const byFile=new Map(); for(const t of selected){if(!byFile.has(t.f))byFile.set(t.f,[]);byFile.get(t.f).push(t);}
  let failed=0;
  for(const [f,fts] of byFile){
    if(f.e==='pdf'){
      let pdf=null;
      try{const ab=await f.blob.arrayBuffer();pdf=await pdfjsLib.getDocument({data:ab}).promise;}catch(e){failed+=fts.length;f.parseStatus='ocr-pending';f.parseError=`OCR: не удалось открыть PDF — ${e.message}`;continue;}
      const pageErrors=[];
      for(const t of fts){
        try{
          const pg=await pdf.getPage(t.page),r=await ocrPdfPage(pg),rec=f.pages.find(x=>x.number===t.page);
          if(rec){rec.text=r.text;rec.ocrConfidence=r.confidence;rec.ocrPending=false;}
          f.ocrPendingPages=(f.ocrPendingPages||[]).filter(x=>x!==t.page);done++;updateProgress(done+failed,selected.length,`${f.name} · стр. ${t.page}`);
        }catch(e){failed++;pageErrors.push(`${t.page}: ${e.message}`);updateProgress(done+failed,selected.length,`${f.name} · стр. ${t.page} — ошибка OCR`);}
        await new Promise(r=>setTimeout(r,0));
      }
      f.text=f.pages.map(x=>x.text||'').join('\n');
      f.parseStatus=(f.ocrPendingPages||[]).length?'ocr-pending':'done';
      f.parseError=pageErrors.length?`OCR не выполнен для страниц: ${pageErrors.slice(0,8).join('; ')}${pageErrors.length>8?'…':''}`:'';
    }else{
      try{
        const r=await recognizeOcr(f.blob);f.text=r.text;f.pages=[{number:pageNumFromName(f.name),text:f.text,ocrConfidence:r.confidence,ocrPending:false}];f.ocrPendingImage=false;f.parseStatus='done';f.parseError='';done++;updateProgress(done+failed,selected.length,f.name);
      }catch(e){failed++;f.parseStatus='ocr-pending';f.parseError=`OCR: ${e.message}`;updateProgress(done+failed,selected.length,`${f.name} — ошибка OCR`);}
      await new Promise(r=>setTimeout(r,0));
    }
  }
  rebuildAnalysisFromParsedFiles(); S.processing=false; $('#runAnalysisBtn').disabled=false; updateOcrButton(); renderAll();
  const left=pendingOcrTasks().length; toast(`OCR завершён: успешно ${done}, ошибок ${failed}, осталось в очереди ${left}`,failed?'warn':'ok');
}
function rebuildAnalysisFromParsedFiles(){
  S.issues=[];S.legalRefs=[];S.duplicates=[];S.conflicts=[];S.recommendations=[];
  for(const p of S.packages){if(p.importedAnalysis)continue;p.units=[];for(const f of p.files)p.units.push(...unitsFromFile(f,p));detectPrimaryFiles(p,true);}
  runLocalAnalysis();buildRecommendations();buildQuality();
}

function pageNumFromName(n){const m=n.match(/P(\d{1,4})/i);return m?Number(m[1]):null;}
function semanticTextQuality(text){
  const t=String(text||'').replace(/\s+/g,' ').trim(); if(t.length<80)return 0;
  const ws=t.match(/[A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі]{2,}/g)||[]; if(ws.length<10)return .15;
  const vowel=/[аеёиоуыэюяәіөұүaeyiou]/i; const natural=ws.filter(w=>vowel.test(w)&&w.length<=28).length/ws.length;
  const veryLong=ws.filter(w=>w.length>28).length/ws.length; const weird=(t.match(/[<>|{}_^~`]/g)||[]).length/Math.max(1,t.length);
  const repeat=/([A-Za-zА-Яа-яЁёӘәҒғҚқҢңӨөҰұҮүҺһІі])\1{7,}/i.test(t)?0.35:0;
  return Math.max(0,Math.min(1,natural-veryLong*.8-weird*4-repeat));
}
function unitsFromFile(f,p){ const units=[];
  if(f.e==='doc' && f.parseStatus==='legacy-doc-partial'){const q=semanticTextQuality(f.text);f.semanticQuality=q;if(q<.42){f.semanticExcluded=true;return units;}} for(const pg of f.pages||[]){ const text=pg.text||''; if(!text.trim())continue; const lines=text.replace(/\r/g,'\n').split(/\n+|(?<=\.)\s{2,}/).map(x=>x.trim()).filter(Boolean); let cur=null; const flush=()=>{if(cur&&cur.text.length>20){cur.normText=norm(cur.text);cur.tokens=uniq(words(cur.text));units.push(cur);}cur=null;};
    for(const line of lines){ const m=line.match(/^((?:\d{1,3}\.){0,5}\d{1,3})[.)]?\s+(.{3,})$/); if(m){flush();cur={id:`U-${p.key}-${f.name}-${pg.number||0}-${units.length}`,vndId:p.match?.vndId||p.vndId||p.key,packageKey:p.key,sourceFile:f.name,page:pg.number,point:m[1],text:m[2],file:f};} else { if(!cur)cur={id:`U-${p.key}-${f.name}-${pg.number||0}-${units.length}`,vndId:p.match?.vndId||p.vndId||p.key,packageKey:p.key,sourceFile:f.name,page:pg.number,point:'',text:'',file:f}; cur.text+=(cur.text?' ':'')+line; if(cur.text.length>1600)flush(); }
    } flush(); }
  if(!units.length && f.text?.trim()) units.push({id:`U-${p.key}-${f.name}-0`,vndId:p.match?.vndId||p.vndId||p.key,packageKey:p.key,sourceFile:f.name,page:null,point:'',text:f.text.slice(0,5000),normText:norm(f.text.slice(0,5000)),tokens:uniq(words(f.text.slice(0,5000))),file:f}); return units; }

function categoryName(id){return S.rules.legalMonitoringCategories.find(c=>c.id===id)?.name||id;}
function reviewBadge(v){const m={confirmed:['ok','Подтверждено'],rejected:['dark','Отклонено'],legal:['warn','Юр. проверка'],ai:['info','Выявлено системой']};const x=m[v]||m.ai;return `<span class="badge ${x[0]}">${x[1]}</span>`;}
function addIssue(obj){ const key=[obj.categoryId,obj.vndId,obj.page,obj.point,norm(obj.issue),obj.otherVndId||''].join('|'); if(S.issues.some(x=>x._key===key))return; S.issues.push({id:`PM-${String(S.issues.length+1).padStart(4,'0')}`,_key:key,severity:'medium',confidence:.65,reviewStatus:'ai',...obj}); }
function runLocalAnalysis(){
  S.analysisPairLimitHit=false;S.analysisCandidatePairs=0;
  const all=S.packages.flatMap(p=>p.units||[]); const pointSets=new Map(); for(const p of S.packages) pointSets.set(p.key,new Set((p.units||[]).map(u=>u.point).filter(Boolean)));
  for(const p of S.packages){ const rec=p.match; if(rec && /не актуал|утрат/i.test(rec.status+' '+rec.change)) addIssue({categoryId:'PM03',vndId:rec.vndId,documentTitle:rec.title,sourceFile:p.files[0]?.name||p.containerName,page:null,point:'',snippet:rec.change||rec.status,issue:'Реестр содержит признак утраты актуальности или изменения документа.',recommendation:'Проверить действующую редакцию и статус ВНД; при необходимости исключить устаревшую редакцию из действующего массива.',confidence:.9,severity:'high'});
    for(const u of p.units){ analyzeUnit(u,p,pointSets.get(p.key)); for(const ref of extractLegalRefs(u)){const cls=classifyReference(ref.raw,u.text);const k=[cls.type,norm(ref.raw)].join('|');let ex=S.legalRefs.find(x=>x.key===k);if(!ex){ex={id:`LR-${String(S.legalRefs.length+1).padStart(4,'0')}`,key:k,raw:ref.raw,type:cls.type,typeLabel:cls.label,event:cls.event||'',status:cls.status,occurrences:[]};S.legalRefs.push(ex);}ex.occurrences.push({vndId:u.vndId,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text.slice(0,280)});}
    }
  }
  compareUnits(all); applyImportedReviewDecisions(); S.duplicates=S.issues.filter(i=>i.categoryId==='PM05'); S.conflicts=S.issues.filter(i=>i.categoryId==='PM02');
}
function analyzeUnit(u,p,pointSet){ const low=norm(u.text);
  for(const pat of S.rules.blanketPatterns||[]) if(low.includes(norm(pat))){addIssue({categoryId:'PM07',vndId:u.vndId,documentTitle:p.match?.title||p.containerName,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text,issue:`Обнаружена общая отсылочная формулировка «${pat}» без автоматически установленного конкретного источника.`,recommendation:'Уточнить необходимость отсылки и, если требуется однозначное применение, указать конкретный документ/норму.',confidence:.78});break;}
  for(const pat of S.rules.discretionPatterns||[]) if(low.includes(norm(pat))){addIssue({categoryId:'PM04',vndId:u.vndId,documentTitle:p.match?.title||p.containerName,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text,issue:`Выявлен потенциальный фактор широкого усмотрения: «${pat}».`,recommendation:'Проверить наличие объективных критериев, оснований, сроков и механизма контроля. Требуется правовая/антикоррупционная оценка специалистом.',confidence:.55,severity:'medium'});break;}
  const registryInactive=!!(p.match&&/не актуал|утрат/i.test((p.match.status||'')+' '+(p.match.change||'')));
  if(!registryInactive)for(const marker of S.rules.legacyMarkers||[]) if(low.includes(norm(marker))){addIssue({categoryId:'PM03',vndId:u.vndId,documentTitle:p.match?.title||p.containerName,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text,issue:`В тексте действующего/не помеченного как утративший силу ВНД обнаружен исторический маркер «${marker}».`,recommendation:'Проверить, соответствует ли историческое наименование/система текущему статусу документа. Если ВНД продолжает действовать — актуализировать формулировку; если документ утратил силу — корректно отразить это в реестре.',confidence:.58});break;}
  const refs=[...u.text.matchAll(/(?:пункт(?:а|у|ом|е)?|п\.)\s*(\d+(?:\.\d+)*)/giu)].map(m=>m[1]); if(/настоящ(?:их|его|ей)\s+(?:правил|положения|инструкции|регламента)/iu.test(u.text)){for(const r of refs)if(pointSet&&!pointSet.has(r)){addIssue({categoryId:'PM01',vndId:u.vndId,documentTitle:p.match?.title||p.containerName,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text,issue:`Внутренняя ссылка на пункт ${r} не найдена среди распознанных пунктов этого документа.`,recommendation:'Проверить нумерацию и целевой пункт ссылки в утверждённом экземпляре.',confidence:.72,severity:'high'});}}
}
function extractLegalRefs(u){
  const out=[]; const t=u.text;
  const patterns=[
    /(?:ст(?:атья|атьи|атье|атью|атьей|\.)\s*\d+(?:[-.]\d+)?(?:\s*(?:част(?:ь|и|ью)|ч\.)\s*\d+)?\s+)?(?:АППК\s*РК|Административн(?:ый|ого)\s+процедурно-процессуальн(?:ый|ого)\s+кодекс(?:а)?\s+Республики Казахстан|Трудов(?:ой|ого)\s+кодекс(?:а)?\s*(?:РК|Республики Казахстан)?|Гражданск(?:ий|ого)\s+кодекс(?:а)?\s*(?:РК|Республики Казахстан)?|Кодекс[^.;\n]{0,90}?Республики Казахстан|Закон(?:а)?\s+Республики Казахстан\s+[«"][^»"]+[»"])(?:[^.;\n]{0,120})?/giu,
    /(?:АППК\s*РК|Административн(?:ый|ого)\s+процедурно-процессуальн(?:ый|ого)\s+кодекс(?:а)?\s+Республики Казахстан|Трудов(?:ой|ого)\s+кодекс(?:а)?\s*(?:РК|Республики Казахстан)?|Гражданск(?:ий|ого)\s+кодекс(?:а)?\s*(?:РК|Республики Казахстан)?|Кодекс[^.;\n]{0,90}?Республики Казахстан|Закон(?:а)?\s+Республики Казахстан\s+[«"][^»"]+[»"])(?:[^.;\n]{0,80})?ст(?:атья|атьи|атье|атью|атьей|\.)\s*\d+(?:[-.]\d+)?(?:\s*(?:част(?:ь|и|ью)|ч\.)\s*\d+)?(?:[^.;\n]{0,60})?/giu,
    /(?:пункт(?:а|у|ом|е)?|п\.)?\s*\d*(?:\.\d+)*\s*(?:приказ(?:а|ом|у)?)[^.;\n]{0,160}?№\s*[\w\-/]+(?:[^.;\n]{0,120})?/giu,
    /(?:пункт(?:а|у|ом|е)?|п\.)?\s*\d*(?:\.\d+)*\s*(?:постановлен(?:ие|ия))[^.;\n]{0,160}?№\s*[\w\-/]+(?:[^.;\n]{0,120})?/giu
  ];
  const seen=new Set(); for(const rx of patterns) for(const m of t.matchAll(rx)){const raw=m[0].trim().replace(/\s+/g,' ');const k=norm(raw);if(raw.length>3&&!seen.has(k)){seen.add(k);out.push({raw});}}
  return out;
}
function classifyReference(raw,context=''){
  const s=norm(raw+' '+context.slice(0,500));
  const isKtz=/(?:ао\s*[«"]?нк\s*[«"]?(?:ктж|қтж)|национальная компания\s*[«"]?(?:казахстан темир жолы|қазақстан темір жолы)|\bктж\b|\bқтж\b)/iu.test(s);
  const isVzhdo=/(?:ао|тоо)?\s*[«"]?(?:вждо|военизированн(?:ая|ой)\s+железнодорожн(?:ая|ой)\s+охран(?:а|ы))|председател[ья]\s+правления\s+(?:ао|тоо)|президент(?:а)?\s+ао\s*[«"]?вждо|приказ(?:ом|а)?\s+общества|приказ(?:ом|а)?\s+товарищества/iu.test(s);
  const isLaw=/(?:аппк\s*рк|кодекс(?:а)?\s+(?:республики казахстан|рк)|закон(?:а)?\s+республики казахстан|постановлен(?:ие|ия)\s+правительства\s+республики казахстан|указ(?:а)?\s+президента\s+республики казахстан|приказ(?:а)?\s+(?:министра|министерства|председателя\s+агентства|генерального\s+прокурора)(?:\s|$))/iu.test(s) && !isKtz && !isVzhdo;
  const repeal=/(?:утратил(?:а|о|и)?\s+силу|считать\s+утративш(?:им|ей|ими)\s+силу|отмен(?:ен|ён|ена|ить))/iu.test(s);
  const amend=/(?:внес(?:ены|ти)\s+изменен|изменен(?:ия|ий)|дополнен(?:ия|ий)|в\s+редакции)/iu.test(s);
  const approval=/(?:об\s+утверждени|утвержден(?:о|а|ы)?\s+приказом|приказ\s+председателя\s+правления|приказ\s+президента)/iu.test(s);
  if(isLaw)return{type:'LAW',label:'НПА Республики Казахстан',status:'не проверено',event:''};
  if(isKtz){if(repeal)return{type:'REPEAL_ACT',label:'Корпоративный акт КТЖ · утрата силы',status:'требует корпоративного подтверждения',event:'repeal'};if(amend)return{type:'AMENDMENT_ACT',label:'Корпоративный акт КТЖ · изменение',status:'требует корпоративного подтверждения',event:'amendment'};if(approval)return{type:'APPROVAL_ACT',label:'Корпоративный акт КТЖ · утверждение',status:'требует корпоративного подтверждения',event:'approval'};return{type:'KTZ_CORPORATE',label:'Корпоративный акт КТЖ',status:'требует корпоративного подтверждения',event:''};}
  if(isVzhdo||/\b№\s*[\w\-/]+/u.test(s)&&/приказ/iu.test(s)){if(repeal)return{type:'REPEAL_ACT',label:'Внутренний акт ВЖДО · утрата силы',status:'требует внутренней проверки',event:'repeal'};if(amend)return{type:'AMENDMENT_ACT',label:'Внутренний акт ВЖДО · изменение',status:'требует внутренней проверки',event:'amendment'};if(approval)return{type:'APPROVAL_ACT',label:'Внутренний акт ВЖДО · утверждение',status:'требует внутренней проверки',event:'approval'};return{type:'INTERNAL_VND',label:'Внутренний акт ВЖДО',status:'требует внутренней проверки',event:''};}
  return{type:'UNKNOWN',label:'Ссылка требует идентификации',status:'требует уточнения',event:''};
}
function tokenJaccard(a,b){ const A=new Set(a),B=new Set(b);if(!A.size||!B.size)return 0;let n=0;for(const x of A)if(B.has(x))n++;return n/(A.size+B.size-n); }

function normListOverlap(a,b){const A=new Set((a||[]).map(norm)),B=new Set((b||[]).map(norm));if(!A.size||!B.size)return 0;let n=0;for(const x of A)if(B.has(x))n++;return n/Math.max(A.size,B.size);}
function extractNormProfile(text){
  const raw=String(text||''),n=norm(raw),tokens=uniq(words(raw));
  const actionPatterns={
    control:['контролирует','контроль','осуществляет контроль','проверяет','проверка'],
    monitoring:['мониторинг','осуществляет мониторинг','отслеживает'],
    coordinate:['координирует','координация'],
    organize:['организует','организация'],
    execute:['исполняет','исполняют','исполнить','исполнение','выполняет','выполняют','выполнить','выполнение'],
    approve:['утверждает','утверждают','утвердить','утверждение','утвержден','утверждён'],
    agree:['согласовывает','согласовывают','согласовать','согласование','согласовывается'],
    decide:['принимает решение','решение принимает','решает'],
    send:['направляет','направляют','направить','направление','предоставляет','предоставляют','предоставить','представляет','представить','передает','передаёт','передать'],
    register:['регистрирует','регистрация','учет','учёт'],
    develop:['разрабатывает','разработка','формирует','подготавливает','готовит'],
    inform:['информирует','уведомляет','сообщает'],
    store:['хранит','хранение','архивирует','архивирование']
  };
  const actions=[];for(const [k,arr] of Object.entries(actionPatterns))if(arr.some(x=>n.includes(norm(x))))actions.push(k);
  const actorCandidates=[];
  const known=[];
  for(const r of S.registry||[]){if(r.department&&r.department!=='Не определено')known.push(r.department);if(r.responsible)known.push(r.responsible);}
  for(const st of S.structure||[]){known.push(st.department,st.shortName,st.responsible,...(st.aliases||[]));}
  for(const k of uniq(known)){const nk=norm(k);if(nk&&nk.length>4&&n.includes(nk))actorCandidates.push(k);}
  const genericActors=[
    ['Организационно-контрольная служба',/организационно[ -]?контрольн(?:ая|ой)\s+служб/iu],
    ['Руководитель структурного подразделения',/руководител[ья]\s+структурн(?:ого|ым)\s+подразделени/iu],
    ['Структурное подразделение',/структурн(?:ое|ого|ым)\s+подразделени/iu],
    ['Ответственный исполнитель',/ответственн(?:ый|ого)\s+исполнител/iu],
    ['Исполнитель',/(?:^|\s)исполнител[ьяюем]?\b/iu],
    ['Директор филиала',/директор(?:а|ом)?\s+филиал/iu],
    ['Филиал',/(?:^|\s)филиал(?:а|е|ом|ы)?\b/iu],
    ['Председатель Правления',/председател[ья]\s+правлени/iu],
    ['Правление',/(?:^|\s)правлени(?:е|я|ем)\b/iu],
    ['Руководитель',/(?:^|\s)руководител[ьяюем]?\b/iu],
    ['Комиссия',/(?:^|\s)комисси(?:я|и|ей)\b/iu]
  ];
  for(const [lab,rx] of genericActors)if(rx.test(raw))actorCandidates.push(lab);
  let actors=uniq(actorCandidates);if(actors.includes('Руководитель структурного подразделения'))actors=actors.filter(x=>x!=='Структурное подразделение'&&x!=='Руководитель');if(actors.some(x=>x!=='Руководитель'&&/руководител/i.test(x)))actors=actors.filter(x=>x!=='Руководитель');
  let modality='neutral';
  if(/(?:^|[\s,.;:()])(запрещается|не допускается|не вправе|запрещено)(?=$|[\s,.;:()])/iu.test(raw))modality='prohibited';
  else if(/(?:^|[\s,.;:()])(обязан|обязана|обязаны|должен|должна|должны|подлежит|необходимо)(?=$|[\s,.;:()])/iu.test(raw))modality='must';
  else if(/(?:^|[\s,.;:()])(вправе|может|могут|допускается|имеет право)(?=$|[\s,.;:()])/iu.test(raw))modality='may';
  const deadlines=[...raw.matchAll(/(?:^|[\s,.;:()])(\d{1,3})\s*(рабочих|календарных)?\s*(дн(?:я|ей)?|час(?:а|ов)?|месяц(?:а|ев)?)(?=$|[\s,.;:()])/giu)].map(m=>`${m[1]} ${(m[2]||'').trim()} ${(m[3]||'').trim()}`.replace(/\s+/g,' ').trim());
  const conditions=[];for(const rx of [/\bпри\s+[^,.;]{3,90}/giu,/\bв\s+случае\s+[^,.;]{3,90}/giu,/\bесли\s+[^,.;]{3,90}/giu,/\bпосле\s+[^,.;]{3,90}/giu,/\bдо\s+[^,.;]{3,90}/giu]){for(const m of raw.matchAll(rx))conditions.push(m[0].trim());}
  const exceptions=[];for(const rx of [/\bза\s+исключением\s+[^.;]{3,120}/giu,/\bкроме\s+[^.;]{3,120}/giu,/\bв\s+исключительных\s+случаях\b/giu]){for(const m of raw.matchAll(rx))exceptions.push(m[0].trim());}
  const scopes=[];if(/центральн(?:ый|ого)\s+аппарат/iu.test(raw))scopes.push('Центральный аппарат');if(/филиал/iu.test(raw))scopes.push('Филиалы');if(/все\s+работник/iu.test(raw))scopes.push('Все работники');if(/структурн(?:ые|ых)\s+подразделени/iu.test(raw))scopes.push('Структурные подразделения');if(/товариществ/iu.test(raw))scopes.push('Товарищество');
  const responsibilityActors=/ответственност[ьи]\s+(?:возлагается|несет|несут|за)/iu.test(raw)||/(?:несет|несёт|несут)\s+ответственность/iu.test(raw)?actors:[];
  const authority=/(?:вправе|имеет право|может требовать|принимает решение|утверждает|разрешает|запрещает|возвращает)/iu.test(raw);
  const obligation=/(?:обязан|обязана|обязаны|должен|должна|должны|обеспечивает)/iu.test(raw);
  const controlRole=actions.includes('control')||actions.includes('monitoring');
  const decisionRole=actions.includes('approve')||actions.includes('decide');
  const approvalRole=actions.includes('approve')||actions.includes('agree');
  const negated=/(?:^|[\s,.;:()])не(?=$|[\s,.;:()])/iu.test(raw);
  const basis=[...raw.matchAll(/(?:ст(?:атья|\.)\s*\d+(?:\.\d+)?|пункт(?:а|\.)?\s*\d+(?:\.\d+)*|приказ[^.;]{0,80}?№\s*[\w\-/]+|закон[^.;]{0,100}?республики казахстан)/giu)].map(m=>m[0].trim()).slice(0,8);
  const objectStop=new Set([...STOP,'осуществляет','обеспечивает','организует','контролирует','координирует','согласовывает','утверждает','принимает','направляет','предоставляет','обязан','должен','вправе','может','служба','отдел','департамент','руководитель','структурное','подразделение']);
  const objects=tokens.filter(x=>!objectStop.has(x)&&!actors.some(a=>norm(a).split(' ').includes(x))).slice(0,18);
  return {actors,actions,objects,modality,deadlines:uniq(deadlines),conditions:uniq(conditions),exceptions:uniq(exceptions),scopes:uniq(scopes),responsibilityActors,authority,obligation,controlRole,decisionRole,approvalRole,negated,basis:uniq(basis)};
}
function modalityLabel(v){return({must:'обязанность',may:'право/возможность',prohibited:'запрет',neutral:'не выражена явно'})[v]||v;}
function conflictTypeLabel(t){return ({DEADLINE_CONFLICT:'Срок',MODALITY_CONFLICT:'Право / обязанность',PROHIBITION_CONFLICT:'Запрет / разрешение',SUBJECT_CONFLICT:'Субъект / исполнитель',RESPONSIBILITY_CONFLICT:'Ответственность',AUTHORITY_CONFLICT:'Полномочия',CONTROL_CONFLICT:'Контроль',APPROVAL_CONFLICT:'Согласование / утверждение',CONDITION_CONFLICT:'Условия применения',PROCESS_CONFLICT:'Порядок действий',SCOPE_CONFLICT:'Сфера действия',TERM_CONFLICT:'Терминология',NEGATION_CONFLICT:'Отрицание',RESULT_CONFLICT:'Результат',HIERARCHY_CONFLICT:'Иерархия норм',CONTEXT_DIFFERENCE:'Контекст различается'})[t]||t||'Потенциальное противоречие';}
function profileValue(p,k){
  if(k==='actors')return (p.actors||[]).join('; ')||'не определён';
  if(k==='actions')return (p.actions||[]).map(x=>({control:'контроль',monitoring:'мониторинг',coordinate:'координация',organize:'организация',execute:'исполнение',approve:'утверждение',agree:'согласование',decide:'принятие решения',send:'направление/предоставление',register:'регистрация/учёт',develop:'разработка/формирование',inform:'информирование',store:'хранение/архив'})[x]||x).join('; ')||'не определено';
  if(k==='objects')return (p.objects||[]).slice(0,8).join(', ')||'не определён';
  if(k==='modality')return modalityLabel(p.modality);
  if(k==='deadlines')return (p.deadlines||[]).join('; ')||'не установлен';
  if(k==='conditions')return (p.conditions||[]).slice(0,3).join('; ')||'не выделены';
  if(k==='exceptions')return (p.exceptions||[]).slice(0,3).join('; ')||'нет явных';
  if(k==='scopes')return (p.scopes||[]).join('; ')||'не определена';
  if(k==='authority')return p.authority?'есть признаки полномочия':'не выявлено явно';
  if(k==='obligation')return p.obligation?'есть обязанность':'не выявлена явно';
  if(k==='responsibilityActors')return (p.responsibilityActors||[]).join('; ')||'не определена явно';
  if(k==='basis')return (p.basis||[]).slice(0,4).join('; ')||'не выделено';
  return '';
}
function compareNormProfiles(a,b,sim){
  const pa=extractNormProfile(a.text),pb=extractNormProfile(b.text),types=[],diffs=[],same=[];
  const actionOverlap=normListOverlap(pa.actions,pb.actions),objectSim=tokenJaccard(pa.objects,pb.objects),actorOverlap=normListOverlap(pa.actors,pb.actors),scopeOverlap=normListOverlap(pa.scopes,pb.scopes);
  const sameCore=(actionOverlap>0 || sim>=.68) && (objectSim>=.16 || sim>=.68);
  if(actionOverlap>0)same.push('действие'); if(objectSim>=.25)same.push('объект регулирования');
  if(pa.deadlines.length&&pb.deadlines.length&&norm(pa.deadlines[0])!==norm(pb.deadlines[0])){types.push('DEADLINE_CONFLICT');diffs.push(`разные сроки: ${pa.deadlines[0]} ↔ ${pb.deadlines[0]}`);}
  if(pa.modality!==pb.modality&&['must','may','prohibited'].includes(pa.modality)&&['must','may','prohibited'].includes(pb.modality)){types.push(pa.modality==='prohibited'||pb.modality==='prohibited'?'PROHIBITION_CONFLICT':'MODALITY_CONFLICT');diffs.push(`разная модальность: ${modalityLabel(pa.modality)} ↔ ${modalityLabel(pb.modality)}`);}
  if(pa.negated!==pb.negated&&sameCore){types.push('NEGATION_CONFLICT');diffs.push('в одной норме присутствует отрицание, в другой — нет');}
  for(const pair of S.rules.nonEquivalentTerms||[]){const [x,y]=pair.map(norm),na=norm(a.text),nb=norm(b.text);if((na.includes(x)&&nb.includes(y))||(na.includes(y)&&nb.includes(x))){types.push('TERM_CONFLICT');diffs.push(`функционально различающиеся понятия «${pair[0]}» ↔ «${pair[1]}»`);break;}}
  const scopeDifferent=pa.scopes.length&&pb.scopes.length&&scopeOverlap===0;
  if(sameCore&&pa.actors.length&&pb.actors.length&&actorOverlap===0){
    if(scopeDifferent){types.push('CONTEXT_DIFFERENCE');diffs.push(`разные субъекты (${pa.actors.join(', ')} ↔ ${pb.actors.join(', ')}) при различающейся сфере действия`);}
    else {types.push('SUBJECT_CONFLICT');diffs.push(`одинаковая/близкая функция закреплена за разными субъектами: ${pa.actors.join(', ')} ↔ ${pb.actors.join(', ')}`);}
  }
  if(sameCore&&pa.responsibilityActors.length&&pb.responsibilityActors.length&&normListOverlap(pa.responsibilityActors,pb.responsibilityActors)===0){types.push('RESPONSIBILITY_CONFLICT');diffs.push('различаются субъекты ответственности');}
  if(sameCore&&pa.controlRole&&pb.controlRole&&pa.actors.length&&pb.actors.length&&actorOverlap===0&&!scopeDifferent){types.push('CONTROL_CONFLICT');diffs.push('контрольная функция закреплена за разными субъектами');}
  if(sameCore&&pa.approvalRole&&pb.approvalRole&&pa.actors.length&&pb.actors.length&&actorOverlap===0&&!scopeDifferent){types.push('APPROVAL_CONFLICT');diffs.push('согласование/утверждение закреплено за разными субъектами');}
  if(sameCore&&pa.authority!==pb.authority&&(pa.authority||pb.authority)){types.push('AUTHORITY_CONFLICT');diffs.push('объём явно выраженных полномочий различается');}
  if(sameCore&&pa.conditions.length&&pb.conditions.length&&normListOverlap(pa.conditions,pb.conditions)===0){types.push('CONDITION_CONFLICT');diffs.push('условия применения нормы различаются');}
  const realTypes=uniq(types.filter(x=>x!=='CONTEXT_DIFFERENCE'));
  let isConflict=realTypes.length>0;
  // Не считать один только терминологический переход monitoring/control противоречием, если роли явно различаются и нет других конфликтов.
  const monitoringVsControl=(pa.actions.includes('monitoring')&&pb.actions.includes('control'))||(pb.actions.includes('monitoring')&&pa.actions.includes('control'));
  const responsibilityAsymmetric=(pa.responsibilityActors.length>0)!==(pb.responsibilityActors.length>0);
  if(monitoringVsControl&&pa.actors.length&&pb.actors.length&&actorOverlap===0&&responsibilityAsymmetric&&!realTypes.includes('DEADLINE_CONFLICT')&&!realTypes.includes('MODALITY_CONFLICT')&&!realTypes.includes('PROHIBITION_CONFLICT'))isConflict=false;
  if(realTypes.length===1&&realTypes[0]==='TERM_CONFLICT'&&monitoringVsControl&&pa.actors.length&&pb.actors.length&&actorOverlap===0)isConflict=false;
  if(realTypes.length===1&&realTypes[0]==='SUBJECT_CONFLICT'&&scopeDifferent)isConflict=false;
  const primary=realTypes[0]||types[0]||'';
  const matrix=[
    ['Субъект / исполнитель',profileValue(pa,'actors'),profileValue(pb,'actors')],
    ['Действие / функция',profileValue(pa,'actions'),profileValue(pb,'actions')],
    ['Объект регулирования',profileValue(pa,'objects'),profileValue(pb,'objects')],
    ['Право / обязанность',profileValue(pa,'modality'),profileValue(pb,'modality')],
    ['Полномочия',profileValue(pa,'authority'),profileValue(pb,'authority')],
    ['Ответственность',profileValue(pa,'responsibilityActors'),profileValue(pb,'responsibilityActors')],
    ['Срок',profileValue(pa,'deadlines'),profileValue(pb,'deadlines')],
    ['Условия применения',profileValue(pa,'conditions'),profileValue(pb,'conditions')],
    ['Исключения',profileValue(pa,'exceptions'),profileValue(pb,'exceptions')],
    ['Сфера действия',profileValue(pa,'scopes'),profileValue(pb,'scopes')],
    ['Нормативное основание',profileValue(pa,'basis'),profileValue(pb,'basis')]
  ];
  const effect=realTypes.includes('RESPONSIBILITY_CONFLICT')||realTypes.includes('SUBJECT_CONFLICT')||realTypes.includes('CONTROL_CONFLICT')?'Возможны спор о владельце функции, двойной контроль либо перекладывание ответственности.':realTypes.includes('DEADLINE_CONFLICT')?'Один процесс может исполняться по разным срокам, что создаёт риск нарушения и спор о применимой норме.':realTypes.includes('MODALITY_CONFLICT')||realTypes.includes('PROHIBITION_CONFLICT')?'Одинаковое действие получает различный обязательный/разрешительный режим.':'Различия могут приводить к неодинаковому применению одного процесса; требуется контекстная проверка.';
  const recommendation=realTypes.includes('SUBJECT_CONFLICT')||realTypes.includes('CONTROL_CONFLICT')||realTypes.includes('RESPONSIBILITY_CONFLICT')?'Определить владельца функции и разграничить исполнение, текущий контроль, централизованный мониторинг, принятие решения и ответственность; затем синхронизировать связанные ВНД.':realTypes.includes('DEADLINE_CONFLICT')?'Определить единый применимый срок с учётом иерархии документов и области действия, затем привести связанные нормы к единой редакции.':realTypes.includes('MODALITY_CONFLICT')||realTypes.includes('PROHIBITION_CONFLICT')?'Уточнить единый правовой режим действия: обязанность, право, возможность или запрет; проверить основание в вышестоящем документе.':'Сопоставить область действия, роли, полномочия, условия и нормативное основание; при подтверждении конфликта определить основной источник нормы.';
  return {isConflict,primaryType:primary,types:realTypes,differences:diffs,same,profiles:{a:pa,b:pb},matrix,effect,recommendation,metrics:{similarity:sim,actionOverlap,objectSimilarity:objectSim,actorOverlap,scopeOverlap},contextDifferent:scopeDifferent};
}
function compareUnits(all){
  const min=S.cfg.analysis.exactDuplicateMinChars; const exact=new Map(); for(const u of all){if(u.normText.length<min)continue;const k=u.normText;if(!exact.has(k))exact.set(k,[]);exact.get(k).push(u);} for(const arr of exact.values())if(arr.length>1){for(let i=1;i<arr.length;i++)addDuplicate(arr[0],arr[i],1,'Точное совпадение текста нормы');}
  const inv=new Map(); for(let i=0;i<all.length;i++){const u=all[i]; if(u.normText.length<min||u.tokens.length<5)continue; const keys=[...u.tokens].sort((a,b)=>b.length-a.length).slice(0,10); for(const k of keys){if(!inv.has(k))inv.set(k,[]);inv.get(k).push(i);}}
  const counts=new Map(); let candidates=0; candidateLoop: for(const list of inv.values()){if(list.length>100)continue;for(let a=0;a<list.length;a++)for(let b=a+1;b<list.length;b++){const i=list[a],j=list[b];if(all[i].packageKey===all[j].packageKey && all[i].sourceFile===all[j].sourceFile)continue;const key=i<j?`${i}|${j}`:`${j}|${i}`;if(!counts.has(key)){if(candidates>=S.cfg.analysis.maxCandidatePairs){S.analysisPairLimitHit=true;break candidateLoop;}counts.set(key,1);candidates++;}else counts.set(key,counts.get(key)+1);}}S.analysisCandidatePairs=candidates;
  for(const [key,c] of counts){if(c<2)continue;const [i,j]=key.split('|').map(Number),a=all[i],b=all[j];if(a.normText===b.normText)continue;const sim=tokenJaccard(a.tokens,b.tokens);const structured=compareNormProfiles(a,b,sim);if((sim>=.34||structured.metrics.actionOverlap>0)&&structured.isConflict){addConflict(a,b,sim,structured);continue;}if(sim>=S.cfg.analysis.nearDuplicateThreshold)addDuplicate(a,b,sim,'Существенное смысловое/лексическое пересечение');}
}
function addDuplicate(a,b,sim,reason){ const pa=S.packages.find(p=>p.key===a.packageKey),pb=S.packages.find(p=>p.key===b.packageKey);addIssue({categoryId:'PM05',vndId:a.vndId,otherVndId:b.vndId,documentTitle:pa?.match?.title||pa?.containerName||a.vndId,otherDocumentTitle:pb?.match?.title||pb?.containerName||b.vndId,sourceFile:a.sourceFile,otherSourceFile:b.sourceFile,page:a.page,otherPage:b.page,point:a.point,otherPoint:b.point,snippet:a.text,otherSnippet:b.text,similarity:sim,issue:reason,recommendation:'Проверить, является ли повторение необходимой детализацией. При конкурирующей компетенции — унифицировать, разграничить или исключить дубль.',confidence:Math.min(.96,.55+sim*.4)}); }
function conflictReason(a,b){const z=compareNormProfiles({text:a},{text:b},tokenJaccard(words(a),words(b)));return z.isConflict?z.differences.join('; '):'';}
function addConflict(a,b,sim,analysis){const pa=S.packages.find(p=>p.key===a.packageKey),pb=S.packages.find(p=>p.key===b.packageKey);const reason=analysis.differences.length?analysis.differences.join('; '):'Выявлено потенциальное структурное противоречие норм.';addIssue({categoryId:'PM02',vndId:a.vndId,otherVndId:b.vndId,documentTitle:pa?.match?.title||pa?.containerName||a.vndId,otherDocumentTitle:pb?.match?.title||pb?.containerName||b.vndId,sourceFile:a.sourceFile,otherSourceFile:b.sourceFile,page:a.page,otherPage:b.page,point:a.point,otherPoint:b.point,snippet:a.text,otherSnippet:b.text,similarity:sim,issue:reason,recommendation:analysis.recommendation,confidence:Math.min(.96,.58+Math.max(sim,analysis.metrics.objectSimilarity)*.35),severity:'high',conflictType:analysis.primaryType,conflictTypes:analysis.types,conflictMatrix:analysis.matrix,conflictEffect:analysis.effect,conflictMetrics:analysis.metrics,conflictSame:analysis.same});}

function recommendationSetFor(issues){ const map=new Map(); const recText={PM01:'Устранить неработающие внутренние ссылки и уточнить механизм исполнения.',PM02:'Устранить подтверждённые противоречия и определить единый источник нормы.',PM03:'Актуализировать устаревшие нормы, реквизиты, названия и редакции.',PM04:'Проверить дискреционные формулировки и установить объективные критерии/контроль.',PM05:'Унифицировать или разграничить дублирующиеся функции и нормы.',PM06:'Заполнить выявленные пробелы регулирования и ответственности.',PM07:'Сократить неоднозначные отсылки и указать конкретные основания.',PM08:'Устранить иные структурные/терминологические недостатки.'};
  for(const i of issues||[]){const r=S.registry.find(x=>x.vndId===i.vndId);const dep=r?.department||'Не определено';const key=`${dep}|${i.categoryId}`;if(!map.has(key))map.set(key,{id:`REC-${map.size+1}`,department:dep,categoryId:i.categoryId,title:recText[i.categoryId]||'Провести актуализацию.',issueIds:[],vndIds:new Set(),priority:i.severity==='high'?'high':'medium'});const x=map.get(key);x.issueIds.push(i.id);x.vndIds.add(i.vndId);if(i.severity==='high')x.priority='high';}
  return [...map.values()].map(x=>({...x,vndIds:[...x.vndIds]})); }
function buildRecommendations(){ S.recommendations=recommendationSetFor(S.issues.filter(i=>i.reviewStatus!=='rejected')); }
function buildQuality(){ const q=[...S.ingestErrors]; for(const r of S.registry)if(r.idTemporary)q.push({level:'warn',type:'VND_ID',message:`${r.title}: отсутствует постоянный VND_ID в реестре.`});
  if(S.structure.some(x=>/^Блок ответственности:/iu.test(x.department||'')))q.push({level:'warn',type:'Структура',message:'structure.xlsx содержит стартовые блоки ответственности, а не фактические наименования структурных подразделений. Для корректного распределения рекомендаций замените содержание structure.xlsx на действующую структуру, сохранив шапки.'});
  if(S.analysisPairLimitHit)q.push({level:'warn',type:'Охват сравнения',message:`Достигнут технический лимит междокументных сравнений (${S.cfg.analysis.maxCandidatePairs}). Часть потенциальных пар могла не провериться; увеличьте лимит или анализируйте тематическими группами.`});
  for(const p of S.packages.filter(x=>x.importedAnalysis)){const m=p.reportMeta||{};if(m.completeness==='partial')q.push({level:'warn',type:'Отчёт .normatrix',message:`${p.containerName}: машинный отчёт был создан из неполного анализа (OCR/ошибки чтения). Используйте результат с осторожностью или повторно проанализируйте исходный ВНД.`});if(m.analysisRulesVersion&&S.version?.analysisRulesVersion&&m.analysisRulesVersion!==S.version.analysisRulesVersion)q.push({level:'warn',type:'Версия анализа',message:`${p.containerName}: отчёт создан по правилам ${m.analysisRulesVersion}, текущая версия правил ${S.version.analysisRulesVersion}. Нормы будут пересопоставлены текущим движком, но качество исходного извлечения зависит от версии, в которой создан отчёт.`});if(m.importSanitizedUnits)q.push({level:'warn',type:'Отчёт .normatrix',message:`${p.containerName}: при импорте очищено ${m.importSanitizedUnits} фрагментов от признаков бинарного/повреждённого хвоста.`});if(m.importExcludedUnits)q.push({level:'warn',type:'Отчёт .normatrix',message:`${p.containerName}: при импорте исключено ${m.importExcludedUnits} полностью нечитаемых текстовых фрагментов, чтобы они не создавали ложные дубли и противоречия.`});}
  for(const p of S.packages)for(const f of p.files||[])if(f.semanticExcluded)q.push({level:'warn',type:'DOC',message:`${f.name}: извлечённый из старого DOC текст признан недостаточно надёжным для смыслового сравнения и исключён из поиска дублей/противоречий; рекомендуется сверить оригинал или использовать DOCX/PDF.`});
  for(const p of S.packages){if(!p.match)q.push({level:'error',type:'Сопоставление',message:`${p.containerName}: пакет не связан однозначно с реестром.`});if(!p.importedAnalysis&&!p.files.length)q.push({level:'error',type:'Состав пакета',message:`${p.containerName}: архив не содержит поддерживаемых документов.`});if(!p.importedAnalysis&&p.files.length&&!p.primaryFiles?.length)q.push({level:'warn',type:'Основной источник',message:`${p.containerName}: основной источник не определён автоматически. Откройте пакет и выберите файл вручную.`});else if(p.primaryFiles?.length&&p.primaryConfidence<.72)q.push({level:'warn',type:'Основной источник',message:`${p.containerName}: основной источник определён с низкой уверенностью (${Math.round(p.primaryConfidence*100)}%). Рекомендуется проверить выбор.`});for(const f of p.files){if(f.e==='doc'&&f.parseStatus==='legacy-doc-partial')q.push({level:'warn',type:'DOC',message:`${f.name}: текст старого DOC извлечён локально методом best-effort (${f.legacyDocMethod||'binary'}); критические выводы рекомендуется сверять с оригиналом.`});if(f.e==='doc'&&f.parseStatus==='doc-unreadable')q.push({level:'error',type:'DOC',message:`${f.name}: старый DOC не удалось надёжно прочитать локально. Используйте исходный просмотр/конвертацию только для этого файла.`});if(f.parseStatus==='ocr-pending')q.push({level:'warn',type:'OCR',message:`${f.name}: часть страниц ожидает OCR. Нажмите «OCR: распознать ожидающие» — они будут обработаны локальной очередью с прогрессом.`});const oc=f.pages?.find(x=>x.ocrConfidence!=null)?.ocrConfidence;if(oc!=null&&oc<60)q.push({level:'warn',type:'OCR',message:`${f.name}: низкая уверенность OCR (${Math.round(oc)}%); выводы по этому фрагменту требуют проверки оригинала.`});if(f.parseStatus==='error')q.push({level:'error',type:'Чтение файла',message:`${f.name}: ${f.parseError}`});}}
  const dup=Object.entries(S.packages.reduce((m,p)=>(p.vndId&&(m[p.vndId]=(m[p.vndId]||0)+1),m),{})).filter(([,n])=>n>1);for(const [id,n] of dup)q.push({level:'error',type:'Дубли пакетов',message:`${id}: загружено ${n} пакета.`});S.quality=q; }

function renderMonitoring(){
  renderCategoryCards('#monitoringCategories');
  const arr=S.issueFilter?S.issues.filter(i=>i.categoryId===S.issueFilter):S.issues;
  const conf=arr.filter(i=>i.reviewStatus==='confirmed').length,rej=arr.filter(i=>i.reviewStatus==='rejected').length;
  $('#issuesMeta').textContent=(S.issueFilter?`${categoryName(S.issueFilter)} — ${arr.length}`:`Все виды недостатков — ${arr.length}`)+` · подтверждено ${conf}${rej?` · отклонено ${rej}`:''}`;
  const reset=$('#clearIssueFilterBtn'); if(reset){reset.disabled=!S.issueFilter;reset.onclick=()=>{S.issueFilter='';renderMonitoring();};}
  $('#issuesTable').innerHTML=arr.length?monitoringTable(arr):'<div class="empty"><strong>Замечаний нет</strong>Загрузите документы или измените фильтр.</div>';
  $('#issuesTable').querySelectorAll('[data-issue]').forEach(x=>x.onclick=e=>{if(e.target.closest('[data-vnd]'))return;openIssue(x.dataset.issue);});
  $('#issuesTable').querySelectorAll('[data-vnd]').forEach(x=>x.onclick=e=>{e.stopPropagation();openVnd(x.dataset.vnd,x.dataset.vndTab||'monitoring',x.closest('[data-issue]')?.dataset.issue||'');});
}
function monitoringActText(i,d){
  if(i.otherVndId){const od=humanDoc(i.otherVndId,i.otherDocumentTitle);return `${docRequisites(od)?docRequisites(od)+' — ':''}${od.title}`;}
  if(i.reference)return String(i.reference);
  if(i.categoryId==='PM03' && d?.change)return d.change;
  if(i.categoryId==='PM01' && /внутренн(?:яя|ей) ссылка|пункт/i.test(i.issue||''))return `Настоящий ВНД — требуется проверить корректность внутренней ссылки и целевого структурного элемента.`;
  if(i.categoryId==='PM07')return 'Конкретный акт не установлен автоматически. Требуется определить источник бланкетной/отсылочной нормы.';
  if(i.categoryId==='PM04')return 'Конкретный акт несоответствия автоматически не установлен. Требуется правовая/антикоррупционная оценка специалистом.';
  if(i.categoryId==='PM06')return 'Нормативное основание выявленного пробела требует определения специалистом.';
  if(i.categoryId==='PM05')return 'Сопоставляемый ВНД не определён автоматически.';
  if(i.categoryId==='PM02')return 'Сопоставляемый акт/ВНД не определён автоматически.';
  return 'Акт, которому не соответствует норма, автоматически не установлен; требуется уточнение специалистом.';
}
function monitoringMeasureText(i){
  if(i.reviewStatus==='rejected')return 'Замечание отклонено специалистом. Дополнительные меры по данному замечанию не требуются.';
  const rec=String(i.recommendation||'Требуется определить способ устранения после рассмотрения замечания специалистом.').trim();
  const stage=i.reviewStatus==='confirmed'
    ? 'Стадия: замечание подтверждено; наименование проекта документа и стадия разработки/согласования в текущей сессии не указаны.'
    : i.reviewStatus==='legal'
      ? 'Стадия: направлено на юридическую проверку; проект документа пока не определён.'
      : 'Стадия: предложение системы, требуется рассмотрение ответственным структурным подразделением; проект документа пока не определён.';
  return `${rec} ${stage}`;
}
function monitoringUpdateText(i,d){
  const parts=[];
  if(d?.change)parts.push(`По реестру: ${d.change}`);
  if(d?.status)parts.push(`Статус ВНД: ${d.status}`);
  if(!parts.length)parts.push('Сведения об актуализации документа в реестре не указаны.');
  if(i.reviewStatus==='confirmed')parts.push('Замечание подтверждено специалистом; сведения об утверждённом актуализирующем документе в текущей сессии не внесены.');
  return parts.join(' ');
}
function monitoringLocationText(i){
  const parts=[];
  if(i.page)parts.push(`стр. ${i.page}`);
  if(i.point)parts.push(`п. ${i.point}`);
  if(!parts.length)parts.push('Структурный элемент не определён автоматически');
  if(i.sourceFile)parts.push(`файл: ${i.sourceFile}`);
  return parts.join('; ');
}
function monitoringTable(arr){
  return `<div class="table-wrap monitoring-wrap"><table class="data-table legal-monitor-table"><thead><tr>
    <th>Вид недостатка норм поведения</th>
    <th>Наименование ВНД</th>
    <th>Номер и дата принятия ВНД</th>
    <th>Структурные элементы документа<br><small>(абзац, подпункт, пункт, статья, часть и др.)</small></th>
    <th>Номер и наименование акта, которому не соответствует норма поведения документа</th>
    <th>Предпринимаемые меры по устранению выявленных недостатков и противоречий<br><small>(проект документа, стадия разработки/согласования)</small></th>
    <th>Сведения об актуализации документа<br><small>(наименование документа, дата утверждения и пр.)</small></th>
  </tr></thead><tbody>${arr.map(i=>{const d=humanDoc(i.vndId,i.documentTitle);return `<tr class="clickable" data-issue="${i.id}">
    <td><span class="badge ${i.categoryId==='PM02'?'danger':i.categoryId==='PM05'?'warn':'info'}">${esc(categoryName(i.categoryId))}</span> ${reviewBadge(i.reviewStatus)}<div class="monitor-note">${esc(i.issue||'')}</div></td>
    <td><button class="link-btn" data-vnd="${esc(i.vndId)}" data-vnd-tab="monitoring"><strong>${esc(d.title)}</strong></button><div class="monitor-note">${esc(d.department||'Структурное подразделение не определено')}</div></td>
    <td>${esc(docRequisites(d)||'Реквизиты в реестре не указаны')}</td>
    <td>${esc(monitoringLocationText(i))}${i.snippet?`<div class="monitor-excerpt">«${esc(cleanExcerpt(i.snippet,220))}»</div>`:''}</td>
    <td>${esc(monitoringActText(i,d))}</td>
    <td>${esc(monitoringMeasureText(i))}</td>
    <td>${esc(monitoringUpdateText(i,d))}</td>
  </tr>`}).join('')}</tbody></table></div><div class="footer-note">Категория недостатка присваивается автоматически как предварительная классификация. Нажмите на строку, чтобы проверить доказательство и принять решение. В официальный отчёт и Приложение 3 попадают только замечания со статусом «Подтверждено специалистом»; отклонённые замечания исключаются из отчёта и рекомендаций СП.</div>`;
}
function issueTable(arr,compact=false){return `<div class="table-wrap"><table class="data-table"><thead><tr><th>Документ</th><th>Категория</th><th>Место</th><th>Суть</th>${compact?'':'<th>Уверенность</th><th>Статус</th>'}</tr></thead><tbody>${arr.map(i=>{const d=humanDoc(i.vndId,i.documentTitle);return `<tr class="clickable" data-issue="${i.id}"><td><button class="link-btn" data-vnd="${esc(i.vndId)}">${esc(d.title)}</button><br><small>${esc(docRequisites(d)||i.vndId)}</small></td><td><span class="badge ${i.categoryId==='PM02'?'danger':i.categoryId==='PM05'?'warn':'info'}">${esc(categoryName(i.categoryId))}</span></td><td>${esc(locationText(i))}</td><td><div class="truncate" title="${esc(i.issue)}">${esc(i.issue)}</div></td>${compact?'':`<td>${Math.round((i.confidence||0)*100)}%</td><td>${reviewBadge(i.reviewStatus)}</td>`}</tr>`}).join('')}</tbody></table></div>`;}

function bindIssueAndVnd(container,defaultTab='overview'){container.querySelectorAll('[data-issue]').forEach(x=>x.onclick=e=>{if(e.target.closest('[data-vnd]'))return;openIssue(x.dataset.issue);});container.querySelectorAll('[data-vnd]').forEach(x=>x.onclick=e=>{e.stopPropagation();openVnd(x.dataset.vnd,x.dataset.vndTab||defaultTab,x.closest('[data-issue]')?.dataset.issue||'');});}
function renderDuplicates(){const a=S.duplicates;$('#duplicatesTable').innerHTML=a.length?issueTable(a):'<div class="empty"><strong>Дубли не выявлены</strong>Или документы ещё не анализировались.</div>';bindIssueAndVnd($('#duplicatesTable'),'duplicates');}
function conflictTable(arr){return `<div class="table-wrap"><table class="data-table"><thead><tr><th>Тип противоречия</th><th>Документ А</th><th>Место</th><th>Документ Б</th><th>Место</th><th>Что различается</th><th>Уверенность</th><th>Статус</th></tr></thead><tbody>${arr.map(i=>{const a=humanDoc(i.vndId,i.documentTitle),b=humanDoc(i.otherVndId,i.otherDocumentTitle);return `<tr class="clickable" data-issue="${i.id}"><td><span class="badge danger">${esc(conflictTypeLabel(i.conflictType))}</span>${(i.conflictTypes||[]).length>1?`<div class="monitor-note">${esc(i.conflictTypes.slice(1).map(conflictTypeLabel).join('; '))}</div>`:''}</td><td><button class="link-btn" data-vnd="${esc(i.vndId)}" data-vnd-tab="conflicts">${esc(a.title)}</button><br><small>${esc(docRequisites(a)||i.vndId)}</small></td><td>${esc([i.page?`стр. ${i.page}`:'',i.point?`п. ${i.point}`:''].filter(Boolean).join(', ')||'—')}</td><td><button class="link-btn" data-vnd="${esc(i.otherVndId)}" data-vnd-tab="conflicts">${esc(b.title)}</button><br><small>${esc(docRequisites(b)||i.otherVndId||'')}</small></td><td>${esc([i.otherPage?`стр. ${i.otherPage}`:'',i.otherPoint?`п. ${i.otherPoint}`:''].filter(Boolean).join(', ')||'—')}</td><td>${esc(i.issue)}</td><td>${Math.round((i.confidence||0)*100)}%</td><td>${reviewBadge(i.reviewStatus)}</td></tr>`}).join('')}</tbody></table></div><div class="footer-note">Противоречие является предварительным выводом системы. По клику открываются обе нормы, структурная матрица различий, управленческий эффект и рекомендация. В официальный правовой мониторинг запись попадает только после подтверждения специалистом.</div>`;}
function renderConflicts(){const all=S.conflicts,types=uniq(all.map(i=>i.conflictType).filter(Boolean));const cards=$('#conflictTypeCards');if(cards)cards.innerHTML=types.length?types.map(t=>{const n=all.filter(i=>i.conflictType===t).length,c=all.filter(i=>i.conflictType===t&&i.reviewStatus==='confirmed').length;return `<button class="cat-card ${S.conflictFilter===t?'active':''}" data-conflict-type="${esc(t)}"><span>${esc(conflictTypeLabel(t))}</span><b>${n}</b><small>подтверждено ${c}</small></button>`;}).join(''):'<div class="empty">Типы появятся после анализа.</div>';cards?.querySelectorAll('[data-conflict-type]').forEach(b=>b.onclick=()=>{S.conflictFilter=S.conflictFilter===b.dataset.conflictType?'':b.dataset.conflictType;renderConflicts();});const reset=$('#clearConflictFilterBtn');if(reset){reset.disabled=!S.conflictFilter;reset.onclick=()=>{S.conflictFilter='';renderConflicts();};}const a=S.conflictFilter?all.filter(i=>i.conflictType===S.conflictFilter):all;const meta=$('#conflictsMeta');if(meta)meta.textContent=`${a.length} из ${all.length}${S.conflictFilter?` · ${conflictTypeLabel(S.conflictFilter)}`:''}`;$('#conflictsTable').innerHTML=a.length?conflictTable(a):'<div class="empty"><strong>Потенциальные противоречия не выявлены</strong>Это означает только отсутствие срабатываний текущего структурного анализа; при неполном OCR/охвате смотрите «Качество данных».</div>';bindIssueAndVnd($('#conflictsTable'),'conflicts');}

function refBadgeClass(r){return r.status==='подтверждено'?'ok':r.type==='LAW'?'warn':r.type==='UNKNOWN'?'danger':'info';}
async function verifyLegal(){
  const laws=S.legalRefs.filter(r=>r.type==='LAW');
  if(!laws.length){toast('Ссылки на государственные НПА для проверки не найдены.','warn');return;}
  const endpoint=String(S.cfg?.legalVerification?.endpoint||'').trim();
  if(!endpoint){
    for(const r of laws){if(!r.verificationStatus)r.verificationStatus='Требует проверки по официальному источнику';}
    renderLegalRefs();
    toast('Автоматический Legal Verification endpoint не настроен. НПА не помечаются как актуальные автоматически; используйте официальные источники zan.gov.kz / adilet.zan.kz.','warn');
    return;
  }
  // Внешнему сервису разрешено передавать только метаданные публичного НПА. Полный текст ВНД не отправляется.
  const payload={mode:'metadata-only',items:laws.map(r=>({title:r.title||r.label||'',number:r.number||'',date:r.date||'',article:r.article||'',part:r.part||'',sourceText:''}))};
  try{
    const res=await fetch(endpoint,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    if(!res.ok)throw new Error(`HTTP ${res.status}`);
    const data=await res.json();
    const items=Array.isArray(data?.items)?data.items:[];
    for(let k=0;k<laws.length;k++){const v=items[k];if(v)laws[k].verificationStatus=v.status||'Проверено внешним сервисом';}
    renderLegalRefs();toast(`Проверено ссылок НПА: ${items.length}`,'ok');
  }catch(e){S.ingestErrors.push({level:'error',type:'Проверка НПА',message:e.message||String(e)});buildQuality();renderQuality();toast(`Не удалось выполнить проверку НПА: ${e.message||e}`,'error');}
}

function refSourceHtml(r){if(r.type==='LAW')return (S.legalSources?.officialSources||[]).map(s=>`<a href="${s.url}" target="_blank" rel="noopener">${esc(s.name)}</a>`).join('<br>')||'Официальный источник не настроен';if(r.type==='KTZ_CORPORATE'||(r.type!=='LAW'&&/КТЖ/u.test(r.typeLabel||'')))return 'Внешняя проверка не выполняется. Требуется корпоративное подтверждение.';if(r.type==='INTERNAL_VND'||['APPROVAL_ACT','AMENDMENT_ACT','REPEAL_ACT'].includes(r.type))return 'Внешняя проверка не выполняется. Требуется внутренняя проверка/подтверждение.';return 'Требуется идентификация источника.';}
function renderLegalRefs(){ const rows=S.legalRefs;$('#legalRefsTable').innerHTML=rows.length?`<div class="table-wrap"><table class="data-table"><thead><tr><th>Тип</th><th>Ссылка / реквизиты</th><th>ВНД</th><th>Упоминаний</th><th>Статус</th><th>Как проверять</th></tr></thead><tbody>${rows.map(r=>{const vids=uniq(r.occurrences.map(o=>o.vndId)).slice(0,6);return `<tr><td><span class="badge info">${esc(r.typeLabel||r.type||'не определено')}</span></td><td>${esc(r.raw)}</td><td>${vids.map(id=>`<button class="link-btn mini" data-vnd="${esc(id)}" data-vnd-tab="refs">${esc(humanDoc(id).title)}</button>`).join('<br>')}${uniq(r.occurrences.map(o=>o.vndId)).length>6?'<br><small>и другие…</small>':''}</td><td>${r.occurrences.length}</td><td><span class="badge ${refBadgeClass(r)}">${esc(r.status)}</span></td><td>${refSourceHtml(r)}</td></tr>`}).join('')}</tbody></table></div>`:'<div class="empty"><strong>Ссылки не извлечены</strong>После анализа здесь появятся НПА РК, внутренние акты ВЖДО, корпоративные акты КТЖ и события утверждения/изменения/утраты силы.</div>';$('#legalRefsTable').querySelectorAll('[data-vnd]').forEach(x=>x.onclick=()=>openVnd(x.dataset.vnd,x.dataset.vndTab||'refs'));}

function renderDepartments(){ const depNames=uniq(S.registry.map(r=>r.department)).sort((a,b)=>a.localeCompare(b,'ru')); $('#departmentsGrid').innerHTML=depNames.map(dep=>{const docs=S.registry.filter(r=>r.department===dep),ids=new Set(docs.map(r=>r.vndId)),iss=S.issues.filter(i=>ids.has(i.vndId)&&i.reviewStatus!=='rejected'),recs=S.recommendations.filter(r=>r.department===dep),resp=uniq(docs.map(r=>r.responsible)).join('; ');return `<div class="panel clickable" data-dep="${esc(dep)}"><div class="panel-title"><h4>${esc(dep)}</h4><span class="meta">${docs.length} ВНД</span></div><div class="meta">Ответственный: ${esc(resp||'не указан')}</div><div class="stat-line"><div class="stat-pill"><b>${iss.length}</b>замечаний</div><div class="stat-pill"><b>${recs.length}</b>рекомендаций</div><div class="stat-pill"><b>${iss.filter(i=>i.categoryId==='PM02').length}</b>противоречий</div><div class="stat-pill"><b>${iss.filter(i=>i.categoryId==='PM05').length}</b>дублей</div></div></div>`}).join('')||'<div class="empty">Структура не загружена.</div>';$('#departmentsGrid').querySelectorAll('[data-dep]').forEach(x=>x.onclick=()=>openDepartment(x.dataset.dep));}
function vndLinkifiedText(text,tab='overview'){let h=esc(text||'');return h.replace(/VND-\d{6}/g,id=>`<button class="link-btn mini" data-vnd="${id}" data-vnd-tab="${tab}">${id}</button>`);}
function renderQuality(){const err=S.quality.filter(x=>x.level==='error').length,warn=S.quality.filter(x=>x.level==='warn').length,unmatched=S.packages.filter(p=>!p.match).length,loaded=loadedRegistryIds().size,total=S.registry.length,ocr=pendingOcrTasks().length;const cards=[['Критические ошибки',err,'danger'],['Требуют внимания',warn,'warn'],['Загружено для анализа',`${loaded} / ${total}`,'ok'],['Не сопоставлено ВНД',unmatched,unmatched?'danger':'ok'],['OCR ожидает',ocr,ocr?'warn':'ok'],['Временных ID',S.registry.filter(r=>r.idTemporary).length,'warn']];$('#qualityKpis').innerHTML=cards.map(c=>`<div class="kpi-card ${c[2]}"><div class="label">${c[0]}</div><div class="value">${c[1]}</div></div>`).join('');$('#qualityTable').innerHTML=S.quality.length?`<div class="table-wrap"><table class="data-table"><thead><tr><th>Уровень</th><th>Тип</th><th>Файл / источник</th><th>Что произошло</th><th>Что сделать</th></tr></thead><tbody>${S.quality.map(q=>{const z=qualityParts(q);return `<tr><td>${q.level==='error'?'<span class="badge danger">ошибка</span>':'<span class="badge warn">предупреждение</span>'}</td><td>${esc(q.type)}</td><td>${vndLinkifiedText(z.file||'—','overview')}</td><td>${vndLinkifiedText(z.problem,'overview')}</td><td>${vndLinkifiedText(actionForQuality(q),'overview')}</td></tr>`}).join('')}</tbody></table></div>`:'<div class="empty"><strong>Проблем, требующих действия, не выявлено</strong></div>';$('#qualityTable').querySelectorAll('[data-vnd]').forEach(x=>x.onclick=()=>openVnd(x.dataset.vnd,x.dataset.vndTab||'overview'));}
function renderReports(){const confirmed=confirmedIssues().length,pending=pendingIssues().length,rejected=rejectedIssues().length,officialRecs=recommendationSetFor(confirmedIssues()).length;$('#reportPreview').innerHTML=`<div class="stat-line"><div class="stat-pill"><b>${S.packages.length}</b>уникальных ВНД</div><div class="stat-pill"><b>${S.packages.filter(p=>p.parsed).length}</b>проанализировано</div><div class="stat-pill"><b>${S.issues.length}</b>выявлено системой</div><div class="stat-pill"><b>${confirmed}</b>подтверждено / входит в официальный отчёт</div><div class="stat-pill"><b>${pending}</b>на рассмотрении / не входит</div><div class="stat-pill"><b>${rejected}</b>отклонено / не входит</div><div class="stat-pill"><b>${officialRecs}</b>рекомендаций по подтверждённым замечаниям</div></div><div class="footer-note">Официальная выгрузка и Приложение 3 формируются только из замечаний со статусом «Подтверждено специалистом». Отклонённые замечания не включаются и не влияют на рекомендации структурным подразделениям в отчёте.</div>`;}

function openModal(title){$('#modalTitle').textContent=title;$('#detailModal').classList.add('open');}
function closeModal(){$('#detailModal').classList.remove('open');}
function vndTabLabel(t){return ({overview:'Обзор',document:'Документ',monitoring:'Правовой мониторинг',conflicts:'Противоречия',duplicates:'Дубли',refs:'НПА и ссылки',functions:'Функции',recommendations:'Рекомендации',versions:'Версии / изменения'})[t]||t;}
function renderVndTab(id,tab='overview',focusIssue=''){
  const r=S.registry.find(x=>x.vndId===id);if(!r)return '<div class="empty">ВНД не найден.</div>';const p=packageFor(id),iss=issuesFor(id),refs=S.legalRefs.filter(x=>x.occurrences.some(o=>o.vndId===id)),conf=iss.filter(i=>i.categoryId==='PM02'),dup=iss.filter(i=>i.categoryId==='PM05');
  if(tab==='document')return p?.units?.length?`<h3>Структурные элементы документа</h3>${p.units.slice(0,500).map(u=>`<div class="detail-item"><b>${u.point?`п. ${esc(u.point)}`:'Фрагмент'}${u.page?` · стр. ${u.page}`:''}</b><div>${esc(cleanExcerpt(u.text,520))}</div>${p.files?.length?`<button class="btn sm" data-unit-source="${esc(u.sourceFile)}" data-unit-page="${u.page||''}" data-unit-text="${esc(cleanExcerpt(u.text,150))}">Показать в документе</button>`:''}</div>`).join('')}`:'<div class="empty">Структурированные нормы недоступны.</div>';
  if(tab==='monitoring')return iss.length?monitoringTable(iss):'<div class="empty">Замечаний правового мониторинга по этому ВНД нет.</div>';
  if(tab==='conflicts')return conf.length?conflictTable(conf):'<div class="empty">Потенциальных противоречий по этому ВНД не найдено.</div>';
  if(tab==='duplicates')return dup.length?issueTable(dup):'<div class="empty">Дубли и пересечения по этому ВНД не найдены.</div>';
  if(tab==='refs')return refs.length?refs.map(x=>`<div class="detail-item"><b>${esc(x.typeLabel||x.type||'Ссылка')}</b><br>${esc(x.raw)} — <span class="badge ${refBadgeClass(x)}">${esc(x.status)}</span><div class="monitor-note">${esc(refAction(x))}</div></div>`).join(''):'<div class="empty">Ссылки и акты не найдены.</div>';
  if(tab==='functions'){const rows=(p?.units||[]).map(u=>({u,pr:extractNormProfile(u.text)})).filter(x=>x.pr.actions.length||x.pr.actors.length).slice(0,400);return rows.length?`<div class="table-wrap"><table class="data-table"><thead><tr><th>Место</th><th>Субъект</th><th>Функция</th><th>Объект</th><th>Модальность</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc([x.u.page?`стр. ${x.u.page}`:'',x.u.point?`п. ${x.u.point}`:''].filter(Boolean).join(', ')||'—')}</td><td>${esc(profileValue(x.pr,'actors'))}</td><td>${esc(profileValue(x.pr,'actions'))}</td><td>${esc(profileValue(x.pr,'objects'))}</td><td>${esc(profileValue(x.pr,'modality'))}</td></tr>`).join('')}</tbody></table></div>`:'<div class="empty">Функции автоматически не выделены.</div>';}
  if(tab==='recommendations'){const recs=S.recommendations.filter(x=>x.vndIds?.includes(id));return recs.length?recs.map(x=>`<div class="recommendation ${x.priority}"><div class="rec-title">${esc(x.title)}</div><div class="rec-meta">${esc(categoryName(x.categoryId))} · ${x.issueIds.length} замечаний</div></div>`).join(''):'<div class="empty">Рекомендаций по текущему анализу нет.</div>';}
  if(tab==='versions')return `<div class="detail-list"><div class="detail-item"><div class="k">Статус по реестру</div><div class="v">${esc(statusDerived(r))}</div></div><div class="detail-item"><div class="k">Изменения / утрата силы</div><div class="v">${esc(r.change||'В реестре не указано')}</div></div><div class="detail-item"><div class="k">Дата / год</div><div class="v">${esc(fmtDate(r.date)||r.year||'—')}</div></div></div>`;
  const confirmed=iss.filter(i=>i.reviewStatus==='confirmed').length,pending=iss.filter(i=>i.reviewStatus==='ai'||i.reviewStatus==='legal').length,rejected=iss.filter(i=>i.reviewStatus==='rejected').length;
  return `<div class="grid kpi" style="grid-template-columns:repeat(4,1fr)"><div class="kpi-card"><div class="label">Пунктов/фрагментов</div><div class="value">${p?.units?.length||0}</div></div><div class="kpi-card warn"><div class="label">Замечаний</div><div class="value">${iss.length}</div></div><div class="kpi-card danger"><div class="label">Противоречий</div><div class="value">${conf.length}</div></div><div class="kpi-card"><div class="label">Ссылок/актов</div><div class="value">${refs.length}</div></div></div><div class="detail-list"><div class="detail-item"><div class="k">Рассмотрение замечаний</div><div class="v">Подтверждено: ${confirmed} · На рассмотрении: ${pending} · Отклонено: ${rejected}</div></div><div class="detail-item"><div class="k">Источник анализа</div><div class="v">${p?.importedAnalysis?'Машинный отчёт .normatrix':'Исходные документы текущей сессии'}</div></div></div><h4>Последние замечания</h4>${iss.length?issueTable(iss.slice(0,30),true):'<div class="empty">Замечаний нет.</div>'}`;
}
function openVnd(id,tab='overview',focusIssue=''){
  const r=S.registry.find(x=>x.vndId===id);if(!r)return;const p=packageFor(id);openModal(r.title);const tabs=['overview','document','monitoring','conflicts','duplicates','refs','functions','recommendations','versions'];
  $('#modalSide').innerHTML=`<div class="detail-list"><div class="detail-item"><div class="k">Полное наименование ВНД</div><div class="v"><b>${esc(r.title)}</b></div></div><div class="detail-item"><div class="k">VND_ID</div><div class="v mono">${esc(r.vndId)}</div></div><div class="detail-item"><div class="k">Номер и дата принятия</div><div class="v">${esc(docRequisites(humanDoc(id))||'не указаны')}</div></div><div class="detail-item"><div class="k">Вид документа</div><div class="v">${esc(r.type||'—')}</div></div><div class="detail-item"><div class="k">Структурное подразделение</div><div class="v">${esc(r.department||'Не определено')}</div></div><div class="detail-item"><div class="k">Ответственный</div><div class="v">${esc(r.responsible||'не указан')}</div></div><div class="detail-item"><div class="k">Статус</div><div class="v">${esc(statusDerived(r))}</div></div><div class="detail-item"><div class="k">Изменения</div><div class="v">${esc(r.change||'не указаны')}</div></div></div>${p?.files?.length?'<hr><div class="k">Файлы текущей сессии</div>'+p.files.map(f=>`<div class="file-row ${f.isPrimary?'active':''}" data-view-file="${esc(f.name)}">${f.isPrimary?'<span class="badge ok">основной</span> ':''}${esc(f.name)}<br><small>${esc(roleLabel(f.role))} ${esc(f.lang||'')}</small></div>`).join(''):p?.importedAnalysis?'<div class="callout"><b>Загружен .normatrix.</b><br>Исходный файл в текущей сессии отсутствует.</div>':''}`;
  const renderTab=(t)=>{tab=t;$('#modalMain').innerHTML=`<div class="tabs">${tabs.map(x=>`<button data-vndcard-tab="${x}" class="${x===t?'active':''}">${vndTabLabel(x)}</button>`).join('')}</div><div id="vndTabBody">${renderVndTab(id,t,focusIssue)}</div>`;$('#modalMain').querySelectorAll('[data-vndcard-tab]').forEach(b=>b.onclick=()=>renderTab(b.dataset.vndcardTab));bindIssueAndVnd($('#modalMain'),t);$('#modalMain').querySelectorAll('[data-unit-source]').forEach(b=>b.onclick=()=>{if(!p?.files?.length)return;viewFile(p,b.dataset.unitSource,Number(b.dataset.unitPage)||null,b.dataset.unitText||'');});if(focusIssue){const row=$(`#modalMain [data-issue="${CSS.escape(focusIssue)}"]`);if(row)row.scrollIntoView({block:'center'});}};
  renderTab(tab);$('#modalSide').querySelectorAll('[data-view-file]').forEach(x=>x.onclick=()=>viewFile(p,x.dataset.viewFile,null,''));
}

function openPackage(key){const p=S.packages.find(x=>x.key===key);if(!p)return;openModal(p.containerName);$('#modalSide').innerHTML=p.importedAnalysis?'<div class="callout"><b>Загружен машинный отчёт NORMATRIX.</b><br>Исходные PDF/DOCX/JPG в текущей сессии отсутствуют. Доступны сохранённые структурированные нормы и доказательные фрагменты.</div>':p.files.map(f=>`<div class="file-row ${f.isPrimary?'active':''}" data-view-file="${esc(f.name)}">${f.isPrimary?'<span class="badge ok">основной</span> ':''}${esc(f.name)}<br><small>${esc(roleLabel(f.role))} · ${esc(f.parseStatus)}</small><br><button class="btn sm" data-set-primary="${esc(f.name)}">Сделать основным</button></div>`).join('');$('#modalMain').innerHTML=`<h3>${esc(p.match?.title||'Пакет не сопоставлен')}</h3><div class="detail-list"><div class="detail-item"><div class="k">VND_ID</div><div class="v mono">${esc(p.match?.vndId||p.vndId||'—')}</div></div><div class="detail-item"><div class="k">Точность сопоставления</div><div class="v">${Math.round(p.matchConfidence*100)}%</div></div><div class="detail-item"><div class="k">Основной источник</div><div class="v">${p.primaryFiles?.length?`${esc(p.primaryFiles.join(', '))}<br><small>${p.primarySource==='manual'?'выбран вручную':p.primarySource==='explicit'?'помечен MAIN':'определён автоматически'} · ${Math.round(p.primaryConfidence*100)}%</small>`:'Не определён'}</div></div><div class="detail-item"><div class="k">Извлечено норм/фрагментов</div><div class="v">${p.units.length}</div></div></div><div class="callout"><b>MAIN больше не обязателен.</b><br>Портал определяет основной источник по имени, типу, содержимому и совпадению с карточкой ВНД. Если выбор неверен, используйте кнопку «Сделать основным».</div>`;$('#modalSide').querySelectorAll('[data-view-file]').forEach(x=>x.onclick=e=>{if(e.target.closest('[data-set-primary]'))return;viewFile(p,x.dataset.viewFile,null,'');});$('#modalSide').querySelectorAll('[data-set-primary]').forEach(x=>x.onclick=e=>{e.stopPropagation();setPrimaryManual(p,x.dataset.setPrimary);});}
function openIssue(id){
  const i=S.issues.find(x=>x.id===id); if(!i)return;
  openModal(`${i.id} — ${categoryName(i.categoryId)}`);
  $('#modalSide').innerHTML=`<div class="detail-item"><div class="k">Категория</div><div class="v">${esc(categoryName(i.categoryId))}</div></div><div class="detail-item"><div class="k">Уверенность</div><div class="v">${Math.round((i.confidence||0)*100)}%</div></div><div class="detail-item"><div class="k">Статус</div><div class="v">${reviewBadge(i.reviewStatus)}</div></div><hr><button class="btn sm primary" data-review="confirmed">Подтвердить</button> <button class="btn sm" data-review="legal">Юр. проверка</button> <button class="btn sm" data-review="rejected">Отклонить</button>`;
  const sourceA=`<div class="norm-box"><div class="source"><button class="link-btn" data-vnd="${esc(i.vndId)}" data-vnd-tab="${i.categoryId==='PM02'?'conflicts':'monitoring'}">${esc(humanDoc(i.vndId,i.documentTitle).title)}</button> · ${esc(i.sourceFile||'')} ${i.page?`· стр. ${i.page}`:''} ${i.point?`· п. ${esc(i.point)}`:''}</div><div class="text">${esc(i.snippet||'')}</div><br><button class="btn sm" data-source="a">Показать в документе</button></div>`;
  const sourceB=i.otherSnippet
    ? `<div class="norm-box"><div class="source">${i.otherVndId?`<button class="link-btn" data-vnd="${esc(i.otherVndId)}" data-vnd-tab="conflicts">${esc(humanDoc(i.otherVndId,i.otherDocumentTitle).title)}</button>`:esc(i.otherDocumentTitle||'Сопоставляемая норма')} · ${esc(i.otherSourceFile||'')} ${i.otherPage?`· стр. ${i.otherPage}`:''} ${i.otherPoint?`· п. ${esc(i.otherPoint)}`:''}</div><div class="text">${esc(i.otherSnippet)}</div><br><button class="btn sm" data-source="b">Показать в документе</button></div>`
    : `<div class="norm-box"><div class="source">Сопоставляемый источник</div><div class="text">${esc(i.reference||'Не требуется / не установлен автоматически')}</div></div>`;
  const sim=i.similarity!=null?`<div class="detail-item"><div class="k">Смысловое/лексическое сходство</div><div class="v">${Math.round(i.similarity*100)}%</div></div>`:'';
  const conflictHtml=i.categoryId==='PM02'&&i.conflictMatrix?.length?`<h4>Структурный разбор противоречия</h4><div class="callout"><b>Тип:</b> ${esc(conflictTypeLabel(i.conflictType))}${i.conflictTypes?.length>1?`<br><b>Дополнительно:</b> ${esc(i.conflictTypes.slice(1).map(conflictTypeLabel).join('; '))}`:''}</div><div class="diff-grid"><div class="h">Элемент нормы</div><div class="h">Документ А</div><div class="h">Документ Б</div>${i.conflictMatrix.map(r=>`<div><b>${esc(r[0])}</b></div><div>${esc(r[1])}</div><div>${esc(r[2])}</div>`).join('')}</div><div class="callout warn"><b>Что совпадает:</b> ${esc((i.conflictSame||[]).join(', ')||'смысловая близость установлена по тексту')}<br><b>Что различается:</b> ${esc(i.issue)}<br><b>Управленческий эффект:</b> ${esc(i.conflictEffect||'Требуется контекстная проверка.')}</div>`:'';
  $('#modalMain').innerHTML=`<div class="issue-head">${sourceA}${sourceB}</div>${conflictHtml}<div class="callout ${i.categoryId==='PM02'?'danger':'warn'}"><b>Что обнаружено:</b><br>${esc(i.issue)}</div><div class="recommendation ${i.severity==='high'?'high':'medium'}"><div class="rec-title">Рекомендация</div><div>${esc(i.recommendation)}</div></div>${sim}<h4>Данные для официального правового мониторинга</h4><label class="field"><span>Предпринимаемые меры / проект / стадия</span><textarea id="issueMeasureNote" rows="4" placeholder="Например: подготовить проект изменений в пункт 5.1; стадия — на согласовании">${esc(i.measureNote||'')}</textarea></label><label class="field"><span>Сведения об актуализации</span><textarea id="issueUpdateNote" rows="3" placeholder="Например: изменения утверждены приказом №... от ...">${esc(i.updateNote||'')}</textarea></label><button class="btn sm primary" id="saveIssueNotesBtn">Сохранить данные мониторинга</button>`;
  $('#modalSide').querySelectorAll('[data-review]').forEach(b=>b.onclick=()=>{i.reviewStatus=b.dataset.review;S.importedReviewDecisions.set(stableIssueKey(i),i.reviewStatus);S.importedReviewData.set(stableIssueKey(i),{reviewStatus:i.reviewStatus,measureNote:i.measureNote||'',updateNote:i.updateNote||''});buildRecommendations();renderAll();openIssue(id);});
  const saveBtn=$('#saveIssueNotesBtn');if(saveBtn)saveBtn.onclick=()=>{i.measureNote=$('#issueMeasureNote')?.value||'';i.updateNote=$('#issueUpdateNote')?.value||'';S.importedReviewData.set(stableIssueKey(i),{reviewStatus:i.reviewStatus||'ai',measureNote:i.measureNote,updateNote:i.updateNote});toast('Данные правового мониторинга сохранены в текущей сессии','ok');renderAll();};
  $('#modalMain').querySelector('[data-source="a"]')?.addEventListener('click',()=>openIssueSource(i,false));
  $('#modalMain').querySelector('[data-source="b"]')?.addEventListener('click',()=>openIssueSource(i,true));
  $('#modalMain').querySelectorAll('[data-vnd]').forEach(x=>x.onclick=e=>{e.stopPropagation();openVnd(x.dataset.vnd,x.dataset.vndTab||'overview',id);});
}
function openIssueSource(i,other){const vid=other?i.otherVndId:i.vndId,fileName=other?i.otherSourceFile:i.sourceFile,page=other?i.otherPage:i.page,snip=other?i.otherSnippet:i.snippet;const p=packageFor(vid);if(!p||!p.files?.length){toast('В текущей сессии доступен только машинный отчёт ВНД. Исходный документ не загружен.','warn');return;}viewFile(p,fileName,page,snip);}
function viewFile(p,fileName,page,snippet){const f=p.files.find(x=>x.name===fileName)||p.files[0];if(!f)return;$('#modalTitle').textContent=`${p.match?.vndId||p.vndId||''} — ${f.name}`;$('#modalSide').innerHTML=p.files.map(x=>`<div class="file-row ${x===f?'active':''}" data-view-file="${esc(x.name)}">${esc(x.name)}<br><small>${esc(x.role)} · ${esc(x.parseStatus)}</small></div>`).join('');let body=snippet?`<div class="snippet"><b>Найденный фрагмент${page?` · стр. ${page}`:''}:</b><br>${esc(snippet)}</div>`:'';if(f.e==='pdf'){const frag=page?`#page=${page}${snippet?`&search=${encodeURIComponent(String(snippet).replace(/\s+/g,' ').slice(0,90))}`:''}`:'';body+=`<iframe class="viewer-frame" src="${f.objectUrl}${frag}"></iframe>`;}else if(['jpg','jpeg','png','tif','tiff','webp'].includes(f.e))body+=`<img class="viewer-img" src="${f.objectUrl}" alt="${esc(f.name)}">`;else if(f.e==='docx'&&f.html)body+=`<div class="viewer-html">${highlightHtml(f.html,snippet)}</div>`;else body+=`<pre class="viewer-text">${esc(f.text||'Предпросмотр текста недоступен для этого формата.')}</pre>`;$('#modalMain').innerHTML=body;$('#modalSide').querySelectorAll('[data-view-file]').forEach(x=>x.onclick=()=>viewFile(p,x.dataset.viewFile,null,''));}
function highlightHtml(html,snip){if(!snip)return html;const needle=esc(snip.slice(0,70));return html.replace(new RegExp(needle.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'i'),m=>`<mark>${m}</mark>`);}
function openDepartment(dep){const docs=S.registry.filter(r=>r.department===dep),ids=new Set(docs.map(r=>r.vndId)),iss=S.issues.filter(i=>ids.has(i.vndId)&&i.reviewStatus!=='rejected'),recs=S.recommendations.filter(r=>r.department===dep),resp=uniq(docs.map(r=>r.responsible)).join('; ');openModal(dep);$('#modalSide').innerHTML=`<div class="detail-item"><div class="k">Ответственный</div><div class="v">${esc(resp||'не указан')}</div></div><div class="detail-item"><div class="k">ВНД</div><div class="v">${docs.length}</div></div><div class="detail-item"><div class="k">Замечаний</div><div class="v">${iss.length}</div></div><div class="detail-item"><div class="k">Рекомендаций</div><div class="v">${recs.length}</div></div>`;$('#modalMain').innerHTML=`<h3>Рекомендации подразделению</h3>${recs.length?recs.map(r=>`<div class="recommendation ${r.priority}"><div class="rec-title">${esc(r.title)}</div><div class="rec-meta">${categoryName(r.categoryId)} · ${r.issueIds.length} замечаний · ${r.vndIds.length} ВНД</div><div style="margin-top:8px">${r.vndIds.map(id=>`<button class="btn sm" data-vnd="${id}" data-vnd-tab="recommendations">${esc(humanDoc(id).title)}</button>`).join(' ')}</div></div>`).join(''):'<div class="empty">По текущей сессии рекомендаций нет.</div>'}<h3>ВНД подразделения</h3>${docs.map(d=>`<div class="detail-item clickable" data-vnd="${d.vndId}"><b>${esc(d.title)}</b><br><small>${esc(d.vndId)} · ${esc(docRequisites(humanDoc(d.vndId)))}</small></div>`).join('')}`;$('#modalMain').querySelectorAll('[data-vnd]').forEach(x=>x.onclick=()=>openVnd(x.dataset.vnd,x.dataset.vndTab||'overview'));}

function registryItem(vndId){return S.registry.find(r=>r.vndId===vndId)||null;}
function reviewLabel(status){return ({ai:'Выявлено системой — требует рассмотрения',confirmed:'Подтверждено специалистом',legal:'Требуется юридическая проверка',rejected:'Отклонено специалистом'})[status]||'Требует рассмотрения';}
function confirmedIssues(){return S.issues.filter(i=>i.reviewStatus==='confirmed');}
function rejectedIssues(){return S.issues.filter(i=>i.reviewStatus==='rejected');}
function pendingIssues(){return S.issues.filter(i=>i.reviewStatus==='ai'||i.reviewStatus==='legal');}
function activeIssues(){return S.issues.filter(i=>i.reviewStatus!=='rejected');}
function severityLabel(v){return v==='high'?'Высокий':v==='low'?'Низкий':'Средний';}
function cleanExcerpt(v,max=260){const t=String(v||'').replace(/\s+/g,' ').trim();return t.length>max?t.slice(0,max-1)+'…':t;}
function humanDoc(vndId,fallback=''){
  const r=registryItem(vndId); if(!r)return {title:fallback||vndId||'Документ не определён',number:'',date:'',department:'Не определено',responsible:'',status:'',vndId:vndId||''};
  return {title:r.title||fallback||r.vndId,number:r.number||'',date:fmtDate(r.date),department:r.department||'Не определено',responsible:r.responsible||'',status:r.status||'',vndId:r.vndId};
}
function docRequisites(d){return [d.number?`№ ${d.number}`:'',d.date?`от ${d.date}`:''].filter(Boolean).join(' ');}
function locationText(i){return [i.page?`стр. ${i.page}`:'',i.point?`п. ${i.point}`:''].filter(Boolean).join(', ')||'место не определено автоматически';}
function actionForQuality(x){if(x.type==='OCR')return 'Запустить OCR ожидающих страниц и повторить анализ.';if(x.type==='DOC')return 'Для критических выводов сверить извлечённый текст с оригиналом; при возможности использовать DOCX/PDF.';if(x.type==='ID'||x.type==='VND_ID')return 'Присвоить постоянный VND_ID в реестре.';if(x.type==='ZIP'||x.type==='Состав пакета'||x.type==='Основной источник')return 'Проверить состав пакета и правильность основного документа.';if(x.type==='Структура')return 'Заполнить structure.xlsx фактическими подразделениями и ответственными, сохранив шапки.';if(x.type==='Охват сравнения')return 'Увеличить лимит сравнений либо запускать тематические группы ВНД; не считать аудит полным до устранения предупреждения.';if(x.type==='Отчёт .normatrix'||x.type==='Версия анализа')return 'Проверить предупреждение; при необходимости повторно проанализировать исходный ВНД и сформировать новый .normatrix.';if(x.type==='Конфликт отчётов ВНД'||x.type==='Конфликт пакетов ВНД')return 'Оставить только один корректный вариант VND_ID и повторить сводный анализ.';return 'Проверить сообщение и устранить причину перед официальным использованием результата.';}
function qualityParts(x){const m=String(x.message||'').match(/^([^:]{1,220}):\s*(.*)$/s);return m?{file:m[1],problem:m[2]}:{file:'',problem:String(x.message||'')};}
function refAction(r){if(r.type==='LAW')return r.status&&r.status!=='не проверено'?r.status:'Проверить по официальному источнику законодательства РК.';if(r.type==='INTERNAL_VND')return 'Требует внутренней проверки актуальности и применимости.';if(r.type==='KTZ_CORPORATE')return 'Требует корпоративного подтверждения актуальности и применимости.';if(/APPROVAL|AMENDMENT|REPEAL/.test(r.type||''))return 'Проверить как событие жизненного цикла документа.';return 'Требует уточнения типа и источника.';}
function setSheetWidths(ws,widths){ws['!cols']=widths.map(w=>({wch:w}));}
function addSheet(wb,name,rows,widths){const ws=XLSX.utils.aoa_to_sheet(rows);setSheetWidths(ws,widths);ws['!autofilter']={ref:ws['!ref']||'A1:A1'};XLSX.utils.book_append_sheet(wb,ws,name);return ws;}
function exportXlsx(){
  if(!window.XLSX){toast('Библиотека XLSX недоступна','error');return;}
  const wb=XLSX.utils.book_new(), date=new Date().toLocaleString('ru-RU');
  const confirmed=confirmedIssues(), pending=pendingIssues(), rejected=rejectedIssues(), officialRecs=recommendationSetFor(confirmed);
  const summary=[
    ['NORMATRIX VND Analytics — проверенный отчёт',''],
    ['Дата анализа',date],['ВНД в реестре',S.registry.length],['Загружено для анализа',S.packages.length],['Проанализировано',S.packages.filter(p=>p.parsed).length],
    ['Всего выявлено системой',S.issues.length],['Подтверждено специалистом — включено в официальный отчёт',confirmed.length],['На рассмотрении / юридической проверке — не включено',pending.length],['Отклонено специалистом — не включено',rejected.length],['Нормативных и иных ссылок',S.legalRefs.length],
    ['Правило формирования','Подтверждённые замечания включаются в официальный отчёт и Приложение 3. Отклонённые замечания не экспортируются в перечень замечаний, не попадают в Приложение 3 и не формируют рекомендации структурным подразделениям.'],
    ['Методика',S.version?.methodology||'NORMATRIX']
  ];
  addSheet(wb,'Сводка',summary,[55,105]);

  const docs=[['№','Наименование ВНД','Номер / дата','Структурное подразделение','Ответственный','Статус реестра','Источник данных','Файлов в пакете','Подтверждено','На рассмотрении','Отклонено (не включено)','Противоречия подтверждено','Дубли подтверждено','Устаревшие нормы подтверждено','Отсылочные нормы подтверждено','Рекомендуемое действие']];
  S.packages.forEach((p,n)=>{const vid=p.match?.vndId||p.vndId||'',d=humanDoc(vid,p.containerName),all=issuesFor(vid),conf=all.filter(i=>i.reviewStatus==='confirmed'),pend=all.filter(i=>i.reviewStatus==='ai'||i.reviewStatus==='legal'),rej=all.filter(i=>i.reviewStatus==='rejected');const cnt=id=>conf.filter(i=>i.categoryId===id).length;const action=conf.length?`По подтверждённым замечаниям требуется рассмотреть ${conf.length} вопрос(а/ов) и выполнить меры, указанные в листах «Подтверждено» и «Приложение 3».`:pend.length?`Подтверждённых замечаний нет; ${pend.length} замечаний ещё требуют рассмотрения и не включены в официальный отчёт.`:'По подтверждённым результатам замечания отсутствуют.';docs.push([n+1,d.title,docRequisites(d),d.department,d.responsible,d.status,p.importedAnalysis?'Отчёт .normatrix':'Исходный ВНД',p.files?.length||0,conf.length,pend.length,rej.length,cnt('PM02'),cnt('PM05'),cnt('PM03'),cnt('PM07'),action]);});
  addSheet(wb,'По документам',docs,[6,58,20,30,24,22,18,12,13,14,18,18,18,22,22,65]);

  const confirmedRows=[['№','Наименование ВНД','Номер / дата','Подразделение','Страница / пункт','Категория','Что подтверждено','Фрагмент документа','С чем сопоставляется','Рекомендация','Приоритет','Уверенность','Статус']];
  confirmed.forEach((i,n)=>{const d=humanDoc(i.vndId,i.documentTitle),od=i.otherVndId?humanDoc(i.otherVndId,i.otherDocumentTitle):null;confirmedRows.push([n+1,d.title,docRequisites(d),d.department,locationText(i),categoryName(i.categoryId),i.issue||'',cleanExcerpt(i.snippet,360),od?`${od.title}${docRequisites(od)?` (${docRequisites(od)})`:''}`:(i.reference||''),i.measureNote||i.recommendation||'',severityLabel(i.severity),`${Math.round((i.confidence||0)*100)}%`,reviewLabel(i.reviewStatus)]);});
  if(!confirmed.length)confirmedRows.push(['—','Подтверждённых специалистом замечаний нет','','','','','','','','','','','']);
  addSheet(wb,'Подтверждено',confirmedRows,[6,55,20,28,18,30,55,70,55,65,12,12,28]);

  const pendingRows=[['№','Наименование ВНД','Номер / дата','Подразделение','Страница / пункт','Категория','Что обнаружено','Фрагмент документа','Рекомендация','Статус рассмотрения']];
  pending.forEach((i,n)=>{const d=humanDoc(i.vndId,i.documentTitle);pendingRows.push([n+1,d.title,docRequisites(d),d.department,locationText(i),categoryName(i.categoryId),i.issue||'',cleanExcerpt(i.snippet,320),i.recommendation||'',reviewLabel(i.reviewStatus)]);});
  if(!pending.length)pendingRows.push(['—','Замечаний, ожидающих рассмотрения, нет','','','','','','','','']);
  addSheet(wb,'На рассмотрении',pendingRows,[6,55,20,28,18,30,55,70,65,30]);

  const app=[['№ п/п','Виды недостатков норм поведения','Наименование ВНД','Номер и дата принятия ВНД','Структурные элементы документа (абзац, подпункт, пункт, статья, часть и др.)','Номер и наименование акта, которому не соответствует норма поведения документа','Предпринимаемые меры по устранению выявленных недостатков и противоречий (наименование проекта документа, стадия разработки либо согласования)','Сведения об актуализации документа (наименование документа, дата утверждения и пр.)']];
  confirmed.forEach((i,n)=>{const d=humanDoc(i.vndId,i.documentTitle);app.push([n+1,categoryName(i.categoryId),d.title,docRequisites(d)||'Реквизиты в реестре не указаны',`${monitoringLocationText(i)}${i.snippet?`; фрагмент: «${cleanExcerpt(i.snippet,180)}»`:''}`,monitoringActText(i,d),i.measureNote||monitoringMeasureText(i),i.updateNote||monitoringUpdateText(i,d)]);});
  if(!confirmed.length)app.push(['—','Подтверждённых специалистом замечаний нет','','','','','Приложение 3 формируется только из замечаний со статусом «Подтверждено специалистом».','']);
  addSheet(wb,'Приложение 3',app,[7,32,55,24,60,58,70,55]);

  const rr=[['Структурное подразделение','Ответственный руководитель','Приоритет','Категория','Рекомендация','Количество ВНД','Документы, которых касается рекомендация','Количество подтверждённых замечаний']];
  officialRecs.forEach(r=>{const names=r.vndIds.map(id=>humanDoc(id).title),resp=uniq(r.vndIds.map(id=>humanDoc(id).responsible)).join('; ');rr.push([r.department,resp,severityLabel(r.priority),categoryName(r.categoryId),r.title,r.vndIds.length,names.join('; '),r.issueIds.length]);});
  if(!officialRecs.length)rr.push(['—','','','','Рекомендации по подтверждённым замечаниям отсутствуют','','','']);
  addSheet(wb,'Рекомендации СП',rr,[32,26,12,32,65,14,90,20]);

  const lr=[['Тип источника','Ссылка / реквизиты','Количество упоминаний','Статус','Что требуется сделать']];
  S.legalRefs.forEach(r=>lr.push([r.typeLabel||r.type||'',r.raw,r.occurrences.length,r.status||'не проверено',refAction(r)]));
  addSheet(wb,'Ссылки и акты',lr,[30,95,16,26,60]);

  const q=[['Уровень','Тип','Файл / источник','Что произошло','Что сделать']];
  S.quality.forEach(x=>{const z=qualityParts(x);q.push([x.level==='error'?'Ошибка':'Предупреждение',x.type,z.file,z.problem,actionForQuality(x)]);});
  addSheet(wb,'Качество данных',q,[15,20,70,85,65]);
  XLSX.writeFile(wb,`NORMATRIX_Проверенный_отчет_${new Date().toISOString().slice(0,10)}.xlsx`);
}
function exportJson(){const data={generatedAt:nowIso(),format:'technical-machine-readable',version:S.version,coverage:{registry:S.registry.length,loaded:S.packages.length},issues:S.issues.map(({_key,...x})=>x),legalRefs:S.legalRefs,recommendations:S.recommendations,quality:S.quality};downloadBlob(new Blob([JSON.stringify(data,null,2)],{type:'application/json'}),`NORMATRIX_TECH_${new Date().toISOString().slice(0,10)}.json`);}
function downloadBlob(blob,name){const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),1000);}
function printReport(){
  const w=window.open('','_blank');if(!w)return;
  const confirmed=confirmedIssues(),pending=pendingIssues(),rejected=rejectedIssues(),officialRecs=recommendationSetFor(confirmed);
  const byDoc=S.packages.map(p=>{const vid=p.match?.vndId||p.vndId||'',d=humanDoc(vid,p.containerName),issues=issuesFor(vid).filter(i=>i.reviewStatus==='confirmed');return {vid,d,issues};});
  const recHtml=officialRecs.length?officialRecs.map(r=>`<div class="rec"><b>${esc(r.department)}</b><div class="muted">${esc(categoryName(r.categoryId))} · ${r.issueIds.length} подтверждённых замечаний · ${r.vndIds.length} ВНД</div><div>${esc(r.title)}</div><div class="small">Документы: ${esc(r.vndIds.map(id=>humanDoc(id).title).join('; '))}</div></div>`).join(''):'<p>Рекомендации по подтверждённым замечаниям отсутствуют.</p>';
  const docsHtml=byDoc.map(({d,issues})=>`<section class="doc"><h3>${esc(d.title)}</h3><div class="meta">${esc(docRequisites(d)||'Без реквизитов')} · ${esc(d.department)}${d.responsible?` · ${esc(d.responsible)}`:''}</div><div class="chips"><span>Подтверждено: <b>${issues.length}</b></span></div>${issues.length?`<table><thead><tr><th>Место</th><th>Категория</th><th>Что подтверждено</th><th>Меры / рекомендация</th></tr></thead><tbody>${issues.map(i=>`<tr><td>${esc(locationText(i))}</td><td>${esc(categoryName(i.categoryId))}</td><td>${esc(i.issue)}${i.snippet?`<div class="quote">${esc(cleanExcerpt(i.snippet,220))}</div>`:''}</td><td>${esc(i.measureNote||i.recommendation||'')}</td></tr>`).join('')}</tbody></table>`:'<p class="oktext">Подтверждённых замечаний по этому ВНД нет.</p>'}</section>`).join('');
  w.document.write(`<html><head><meta charset="utf-8"><title>NORMATRIX — проверенный отчёт</title><style>@page{size:A4 portrait;margin:12mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#172433;font-size:10.5pt;line-height:1.35;margin:0}h1{font-size:24pt;color:#17324d;margin:0 0 6px}h2{font-size:16pt;color:#17324d;margin:22px 0 10px;border-bottom:2px solid #d9e3ec;padding-bottom:5px}h3{font-size:13pt;margin:0 0 4px;color:#17324d}.cover{padding:8mm 2mm 4mm}.lead{font-size:11.5pt}.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:14px 0}.card{border:1px solid #dbe3ea;border-radius:8px;padding:9px;background:#f8fafc}.card b{display:block;font-size:17pt;color:#17324d}.note{background:#eef8f1;border-left:4px solid #2f855a;padding:10px;margin:12px 0}.rec{border:1px solid #dbe3ea;border-radius:7px;padding:9px;margin:7px 0;page-break-inside:avoid}.muted,.meta,.small{color:#64748b}.small{font-size:9pt;margin-top:4px}.doc{page-break-before:auto;margin-top:16px}.doc+.doc{page-break-before:always}.chips{display:flex;gap:8px;flex-wrap:wrap;margin:8px 0}.chips span{background:#eef4f8;border-radius:12px;padding:4px 8px;font-size:9pt}table{border-collapse:collapse;width:100%;table-layout:fixed;margin-top:8px}th{background:#edf3f7;color:#17324d;text-align:left}td,th{border:1px solid #cfd9e2;padding:5px;vertical-align:top;word-break:break-word}th:nth-child(1){width:13%}th:nth-child(2){width:20%}th:nth-child(3){width:35%}th:nth-child(4){width:32%}.quote{margin-top:4px;padding:5px;border-left:3px solid #cbd5e1;background:#f8fafc;font-size:9pt;color:#475569}.oktext{color:#287a55}.footer{margin-top:20px;color:#64748b;font-size:9pt}</style></head><body><div class="cover"><h1>NORMATRIX VND Analytics</h1><div class="lead">Проверенный аналитический отчёт по внутренним нормативным документам</div><p><b>Дата анализа:</b> ${new Date().toLocaleString('ru-RU')}</p><div class="summary"><div class="card"><b>${S.packages.length}</b>уникальных ВНД</div><div class="card"><b>${confirmed.length}</b>подтверждено</div><div class="card"><b>${pending.length}</b>на рассмотрении — не включено</div><div class="card"><b>${rejected.length}</b>отклонено — не включено</div></div><div class="note"><b>Правило отчёта.</b> В перечень замечаний и рекомендации ниже включены только замечания со статусом «Подтверждено специалистом». Отклонённые замечания и неподтверждённые кандидаты не включены в официальный результат.</div></div><h2>Рекомендации структурным подразделениям</h2>${recHtml}<h2>Подтверждённые результаты по каждому ВНД</h2>${docsHtml}<div class="footer">Методика: ${esc(S.version?.methodology||'NORMATRIX')}. Технические идентификаторы исключены из человекочитаемого отчёта; решения пользователя сохраняются в машинном .normatrix для последующего импорта.</div></body></html>`);
  w.document.close();w.focus();setTimeout(()=>w.print(),450);
}
function clearSession(){if(!confirm('Удалить все загруженные ВНД, отчёты ВНД и результаты текущей сессии? Корневой реестр и структура останутся.'))return;for(const u of S.objectUrls)URL.revokeObjectURL(u);S.objectUrls.clear();S.packages=[];S.issues=[];S.legalRefs=[];S.duplicates=[];S.conflicts=[];S.quality=[];S.ingestErrors=[];S.recommendations=[];S.importedReviewDecisions.clear();S.importedReviewData.clear();S.issueFilter='';S.conflictFilter='';if($('#fileInput'))$('#fileInput').value='';if($('#analysisReportInput'))$('#analysisReportInput').value='';updateRunAnalysisButton();if($('#runPendingOcrBtn'))$('#runPendingOcrBtn').disabled=true;renderAll();goPage('dashboard');toast('Сессия очищена','ok');}

function bindStaticActions(){ const close=$('#modalClose'),modal=$('#detailModal');if(close)close.onclick=closeModal;if(modal)modal.addEventListener('click',e=>{if(e.target.id==='detailModal')closeModal();});const binds=[['#clearSessionBtn',clearSession],['#verifyLegalBtn',verifyLegal],['#exportXlsxBtn',exportXlsx],['#exportJsonBtn',exportJson],['#exportNormatrixBtn',exportNormatrixBundle],['#printReportBtn',printReport]];for(const [sel,fn] of binds){const el=$(sel);if(el)el.onclick=fn;}window.addEventListener('beforeunload',()=>{for(const u of S.objectUrls)URL.revokeObjectURL(u);}); }
function checkRuntimeDependencies(){
  const deps=[['SheetJS / XLSX',window.XLSX,'XLS/XLSX и корневой реестр'],['JSZip',window.JSZip,'ZIP/DOCX/.normatrix'],['Mammoth',window.mammoth,'DOCX'],['PDF.js',window.pdfjsLib,'PDF'],['Tesseract OCR',window.Tesseract,'OCR изображений/сканов']];
  const missing=deps.filter(([,v])=>!v);
  if(!missing.length)return [];
  for(const [name,,scope] of missing)S.ingestErrors.push({level:'warn',type:'Библиотека браузера',message:`${name} не загружена — недоступно: ${scope}. Остальные функции портала продолжают работать.`});
  toast(`Часть библиотек недоступна: ${missing.map(x=>x[0]).join(', ')}. Файлы всё равно можно добавить в сессию; недоступный анализ будет отмечен явно.`,'warn');
  return missing.map(x=>x[0]);
}

async function boot(){ initNav();setupUpload();bindStaticActions();renderAll();await loadOptionalDependencies();checkRuntimeDependencies();await loadStaticData(); }
document.addEventListener('DOMContentLoaded',boot);
})();
