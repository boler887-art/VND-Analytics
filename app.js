(() => {
'use strict';

const S = {
  cfg:null,rules:null,mapping:null,aliases:null,glossary:null,legalSources:null,version:null,
  registry:[],structure:[],structureMap:new Map(),packages:[],issues:[],legalRefs:[],duplicates:[],conflicts:[],quality:[],ingestErrors:[],recommendations:[],
  currentPage:'dashboard',issueFilter:'',processing:false,objectUrls:new Set(),ocrCount:0
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

async function loadStaticData(){
  try{
    [S.cfg,S.rules,S.mapping,S.aliases,S.glossary,S.legalSources,S.version]=await Promise.all([
      fetchNoCache('config.json'),fetchNoCache('rules.json'),fetchNoCache('mapping.json'),fetchNoCache('aliases.json'),fetchNoCache('glossary.json'),fetchNoCache('legal_sources.json'),fetchNoCache('system_version.json')
    ]);
    if(window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
    await Promise.all([loadRegistry(),loadStructure()]);
    buildStructureMap(); linkRegistryStructure();
    setChip('#chipRegistry',`Реестр: ${S.registry.length} ВНД`,true);
    setChip('#chipStructure',`Структура: ${S.structure.length} записей`,true);
    renderAll();
  }catch(e){ console.error(e); setChip('#chipRegistry','Ошибка источников',false); setChip('#chipStructure','Проверьте запуск через HTTP',false); toast('Не удалось прочитать корневые справочники. Открывайте портал через GitHub Pages или локальный HTTP-сервер, а не file://','error'); }
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
function setChip(sel,text,ok){ const el=$(sel); el.textContent=text; el.classList.remove('ok','err'); el.classList.add(ok?'ok':'err'); }

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
  const loaded=S.packages.length, parsed=S.packages.filter(p=>p.parsed).length, units=S.packages.reduce((a,p)=>a+(p.units?.length||0),0), refs=S.legalRefs.length;
  const confirmed=S.issues.filter(i=>i.reviewStatus==='confirmed').length; const missingIds=S.registry.filter(r=>r.idTemporary).length;
  const cards=[
    ['ВНД в реестре',S.registry.length,'registry',''],['Загружено ВНД',loaded,'upload','ok'],['Проанализировано',parsed,'upload','info'],['Выделено норм',units,'monitoring',''],['Найдено замечаний',S.issues.length,'monitoring',S.issues.length?'warn':'ok'],['Подтверждено',confirmed,'monitoring',confirmed?'danger':''],
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
  target.innerHTML=cats.map(c=>{const n=S.issues.filter(i=>i.categoryId===c.id).length;return `<div class="category-card" data-cat="${c.id}"><div class="num">${n}</div><div class="name">${esc(c.name)}</div></div>`}).join('');
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
  $('#registryTable').innerHTML=`<div class="table-wrap"><table class="data-table"><thead><tr><th>ID</th><th>ВНД</th><th>№ / дата</th><th>Вид</th><th>Структурное подразделение</th><th>Ответственный</th><th>Статус</th><th>Анализ</th></tr></thead><tbody>${rows.map(r=>{const p=packageFor(r.vndId),n=issuesFor(r.vndId).length;return `<tr class="clickable" data-vnd="${r.vndId}"><td class="mono">${esc(r.vndId)}${r.idTemporary?'<br><span class="badge warn">временный</span>':''}</td><td><div class="truncate" title="${esc(r.title)}">${esc(r.title)}</div></td><td>${esc(r.number)}<br><small>${esc(fmtDate(r.date))}</small></td><td>${esc(r.type)}</td><td>${esc(r.department)}</td><td>${esc(r.responsible)}</td><td><span class="badge ${/не актуал/i.test(statusDerived(r))?'danger':'dark'}">${esc(statusDerived(r))}</span></td><td>${p?`<span class="badge ${p.parsed?'ok':'warn'}">${p.parsed?'проверен':'загружен'}</span> ${n?`<span class="badge warn">${n}</span>`:''}`:'<span class="badge">не загружен</span>'}</td></tr>`}).join('')}</tbody></table></div><div class="footer-note">Показано ${rows.length} из ${S.registry.length} записей.</div>`;
  $('#registryTable').querySelectorAll('[data-vnd]').forEach(x=>x.onclick=()=>openVnd(x.dataset.vnd));
}

function setupUpload(){
  const dz=$('#dropZone'),inp=$('#fileInput'); ['dragenter','dragover'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add('drag')})); ['dragleave','drop'].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove('drag')}));
  dz.addEventListener('drop',e=>ingestFiles([...e.dataTransfer.files])); inp.addEventListener('change',()=>ingestFiles([...inp.files])); $('#runAnalysisBtn').onclick=processAllPackages; $('#runPendingOcrBtn').onclick=runPendingOcr;
}
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
async function ingestFiles(files){
  if(!files.length)return; const max=S.cfg.limits.maxVndPackages; if(S.packages.length+files.length>max){toast(`Максимум ${max} объектов загрузки за сессию`,'warn');return;}
  const allowed=new Set(['zip',...(JSON.parse(await fetchNoCache('package_schema.json','text'))).allowedExtensions]); let bad=files.filter(f=>!allowed.has(ext(f.name))); if(bad.length)toast(`Пропущены неподдерживаемые файлы: ${bad.map(x=>x.name).join(', ')}`,'warn');
  const existingBytes=S.packages.flatMap(p=>p.files).reduce((a,f)=>a+(f.size||0),0);
  const incomingBytes=files.reduce((a,f)=>a+(f.size||0),0);
  if(bytesMB(existingBytes+incomingBytes)>S.cfg.limits.maxSessionMB){toast(`Общий объём сессии превысит ${S.cfg.limits.maxSessionMB} MB`,'warn');return;}
  if(S.packages.reduce((a,p)=>a+p.files.length,0)+files.length>S.cfg.limits.maxPhysicalFiles){toast(`Превышен лимит ${S.cfg.limits.maxPhysicalFiles} физических файлов за сессию`,'warn');return;}
  for(const file of files.filter(f=>allowed.has(ext(f.name)))){
    if(bytesMB(file.size)>S.cfg.limits.maxSingleFileMB){S.ingestErrors.push({level:'error',type:'Размер файла',message:`${file.name}: превышен лимит ${S.cfg.limits.maxSingleFileMB} MB`});continue;}
    if(ext(file.name)==='zip') await ingestZip(file); else await addPackage(file.name,[makePhysical(file,file.name)]);
  }
  $('#runAnalysisBtn').disabled=!S.packages.length; refreshMatches(); buildQuality(); renderAll(); if(S.cfg.analysis.autoRunAfterIngest) processAllPackages();
}
function makePhysical(blob,name){ const e=ext(name); const file=blob instanceof File?blob:new File([blob],name,{type:mimeByExt(e)}); const url=URL.createObjectURL(file);S.objectUrls.add(url);return {name,e,blob:file,size:file.size,role:roleFromName(name),lang:langFromName(name),objectUrl:url,text:'',pages:[],html:'',parseStatus:'pending',parseError:''}; }
async function ingestZip(file){
  try{
    const zip=await JSZip.loadAsync(file); const entries=[]; let declaredBytes=0;
    const supported=new Set(['pdf','docx','doc','xlsx','xls','txt','jpg','jpeg','png','tif','tiff','webp']);
    const currentPhysical=S.packages.reduce((a,p)=>a+p.files.length,0);
    const currentBytes=S.packages.flatMap(p=>p.files).reduce((a,f)=>a+(f.size||0),0);
    for(const [name,z] of Object.entries(zip.files)){
      if(z.dir)continue;
      const e=ext(name); if(!supported.has(e))continue;
      if(entries.length+currentPhysical>=S.cfg.limits.maxPhysicalFiles) throw new Error(`превышен лимит ${S.cfg.limits.maxPhysicalFiles} физических файлов`);
      const declared=Number(z?._data?.uncompressedSize||0); declaredBytes+=declared;
      if(declared && bytesMB(currentBytes+declaredBytes)>S.cfg.limits.maxSessionMB) throw new Error(`распакованный объём превысит ${S.cfg.limits.maxSessionMB} MB`);
      const blob=await z.async('blob');
      if(bytesMB(blob.size)>S.cfg.limits.maxSingleFileMB) throw new Error(`${name}: после распаковки превышен лимит ${S.cfg.limits.maxSingleFileMB} MB`);
      entries.push(makePhysical(blob,name));
    }
    await addPackage(file.name,entries,true);
  }catch(e){S.ingestErrors.push({level:'error',type:'ZIP',message:`${file.name}: ${e.message}`});}
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
  $('#packageTable').innerHTML=`<div class="table-wrap"><table class="data-table"><thead><tr><th>Пакет</th><th>Связь с реестром</th><th>Состав / основной источник</th><th>Статус</th><th>Проблемы</th></tr></thead><tbody>${S.packages.map(p=>`<tr class="clickable" data-pkg="${p.key}"><td><b>${esc(p.containerName)}</b><br><span class="mono">${esc(p.vndId||'ID не указан')}</span></td><td>${p.match?`${esc(p.match.vndId)}<br><small>${esc(p.match.title)}</small><br><span class="badge ${p.matchConfidence>.94?'ok':'warn'}">${Math.round(p.matchConfidence*100)}%</span>`:'<span class="badge danger">не сопоставлен</span>'}</td><td>${p.files.length} файлов<br><small>${esc(uniq(p.files.map(f=>roleLabel(f.role))).join(', '))}</small><br>${p.primaryFiles?.length?`<span class="badge ${p.primaryConfidence>=.8?'ok':'warn'}">основной: ${esc(p.primaryFiles.length===1?p.primaryFiles[0]:`${p.primaryFiles.length} страниц/файлов`)}</span> <small>${Math.round(p.primaryConfidence*100)}%</small>`:'<span class="badge warn">основной не определён</span>'}</td><td><span class="badge ${p.parsed?'ok':p.analysisStatus==='processing'?'info':'warn'}">${esc(p.parsed?'анализ завершён':p.analysisStatus)}</span></td><td>${p.errors.length?`<span class="badge danger">${p.errors.length}</span>`:'—'}</td></tr>`).join('')}</tbody></table></div>`;
  $('#packageTable').querySelectorAll('[data-pkg]').forEach(x=>x.onclick=()=>openPackage(x.dataset.pkg));
}

async function processAllPackages(){
  if(S.processing||!S.packages.length)return; S.processing=true; $('#runAnalysisBtn').disabled=true; $('#ingestProgress').classList.remove('hidden');
  S.issues=[];S.legalRefs=[];S.duplicates=[];S.conflicts=[];S.recommendations=[];S.ocrCount=0;
  let done=0; const total=S.packages.length;
  for(const p of S.packages){ p.analysisStatus='processing'; p.units=[]; p.errors=[]; renderPackages();
    for(const f of p.files){ try{await parsePhysical(f); const inferred=inferRoleFromContent(f); if(f.role==='OTHER'&&inferred!=='OTHER')f.role=inferred; p.units.push(...unitsFromFile(f,p));}catch(e){f.parseStatus='error';f.parseError=e.message;p.errors.push(`${f.name}: ${e.message}`);} }
    detectPrimaryFiles(p,true); p.parsed=true;p.analysisStatus='done';done++; updateProgress(done,total,`Обработан ${p.containerName}`);
  }
  runLocalAnalysis(); buildRecommendations(); buildQuality(); S.processing=false; $('#runAnalysisBtn').disabled=false; renderAll(); goPage('dashboard'); toast(`Анализ завершён: ${S.packages.length} ВНД, ${S.issues.length} замечаний-кандидатов`,'ok');
}
function updateProgress(done,total,text){ const pct=Math.round(done/total*100);$('#ingestBar').style.width=`${pct}%`;$('#ingestText').textContent=`${pct}% — ${text}`; }
async function parsePhysical(f){
  f.parseStatus='processing'; const e=f.e;
  if(e==='pdf') await parsePdf(f); else if(e==='docx') await parseDocx(f); else if(e==='xlsx'||e==='xls') await parseSheet(f); else if(e==='txt') await parseTxt(f); else if(['jpg','jpeg','png','tif','tiff','webp'].includes(e)) await parseImage(f); else if(e==='doc') await parseLegacyDoc(f);
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
    if(tx.replace(/\s/g,'').length<20 && S.cfg.analysis.autoOcrImages){
      if(S.ocrCount<S.cfg.limits.maxAutoOcrImages && window.Tesseract){
        try{ S.ocrCount++; const r=await ocrPdfPage(pg); tx=r.text; ocrConfidence=r.confidence; }
        catch(e){ pending.push(i); }
      }else pending.push(i);
    }
    pages.push({number:i,text:tx,ocrConfidence,ocrPending:pending.includes(i)});
  }
  f.pages=pages; f.ocrPendingPages=pending; f.text=pages.map(p=>p.text).join('\n'); if(pending.length)f.parseStatus='ocr-pending';
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
  for(const [f,fts] of byFile){
    try{
      if(f.e==='pdf'){
        const ab=await f.blob.arrayBuffer(),pdf=await pdfjsLib.getDocument({data:ab}).promise;
        for(const t of fts){const pg=await pdf.getPage(t.page);const r=await ocrPdfPage(pg);const rec=f.pages.find(x=>x.number===t.page);if(rec){rec.text=r.text;rec.ocrConfidence=r.confidence;rec.ocrPending=false;}f.ocrPendingPages=(f.ocrPendingPages||[]).filter(x=>x!==t.page);done++;updateProgress(done,selected.length,`${f.name} · стр. ${t.page}`);await new Promise(r=>setTimeout(r,0));}
        f.text=f.pages.map(x=>x.text||'').join('\n');f.parseStatus=f.ocrPendingPages.length?'ocr-pending':'done';
      }else{
        const r=await recognizeOcr(f.blob);f.text=r.text;f.pages=[{number:pageNumFromName(f.name),text:f.text,ocrConfidence:r.confidence,ocrPending:false}];f.ocrPendingImage=false;f.parseStatus='done';done++;updateProgress(done,selected.length,f.name);await new Promise(r=>setTimeout(r,0));
      }
    }catch(e){f.parseStatus='ocr-pending';f.parseError=`OCR: ${e.message}`;}
  }
  rebuildAnalysisFromParsedFiles(); S.processing=false; $('#runAnalysisBtn').disabled=false; updateOcrButton(); renderAll(); toast(`OCR завершён: обработано ${done} из ${selected.length}`,'ok');
}
function rebuildAnalysisFromParsedFiles(){
  S.issues=[];S.legalRefs=[];S.duplicates=[];S.conflicts=[];S.recommendations=[];
  for(const p of S.packages){p.units=[];for(const f of p.files)p.units.push(...unitsFromFile(f,p));detectPrimaryFiles(p,true);}
  runLocalAnalysis();buildRecommendations();buildQuality();
}

function pageNumFromName(n){const m=n.match(/P(\d{1,4})/i);return m?Number(m[1]):null;}
function unitsFromFile(f,p){ const units=[]; for(const pg of f.pages||[]){ const text=pg.text||''; if(!text.trim())continue; const lines=text.replace(/\r/g,'\n').split(/\n+|(?<=\.)\s{2,}/).map(x=>x.trim()).filter(Boolean); let cur=null; const flush=()=>{if(cur&&cur.text.length>20){cur.normText=norm(cur.text);cur.tokens=uniq(words(cur.text));units.push(cur);}cur=null;};
    for(const line of lines){ const m=line.match(/^((?:\d{1,3}\.){0,5}\d{1,3})[.)]?\s+(.{3,})$/); if(m){flush();cur={id:`U-${p.key}-${f.name}-${pg.number||0}-${units.length}`,vndId:p.match?.vndId||p.vndId||p.key,packageKey:p.key,sourceFile:f.name,page:pg.number,point:m[1],text:m[2],file:f};} else { if(!cur)cur={id:`U-${p.key}-${f.name}-${pg.number||0}-${units.length}`,vndId:p.match?.vndId||p.vndId||p.key,packageKey:p.key,sourceFile:f.name,page:pg.number,point:'',text:'',file:f}; cur.text+=(cur.text?' ':'')+line; if(cur.text.length>1600)flush(); }
    } flush(); }
  if(!units.length && f.text?.trim()) units.push({id:`U-${p.key}-${f.name}-0`,vndId:p.match?.vndId||p.vndId||p.key,packageKey:p.key,sourceFile:f.name,page:null,point:'',text:f.text.slice(0,5000),normText:norm(f.text.slice(0,5000)),tokens:uniq(words(f.text.slice(0,5000))),file:f}); return units; }

function categoryName(id){return S.rules.legalMonitoringCategories.find(c=>c.id===id)?.name||id;}
function addIssue(obj){ const key=[obj.categoryId,obj.vndId,obj.page,obj.point,norm(obj.issue),obj.otherVndId||''].join('|'); if(S.issues.some(x=>x._key===key))return; S.issues.push({id:`PM-${String(S.issues.length+1).padStart(4,'0')}`,_key:key,severity:'medium',confidence:.65,reviewStatus:'ai',...obj}); }
function runLocalAnalysis(){
  const all=S.packages.flatMap(p=>p.units||[]); const pointSets=new Map(); for(const p of S.packages) pointSets.set(p.key,new Set((p.units||[]).map(u=>u.point).filter(Boolean)));
  for(const p of S.packages){ const rec=p.match; if(rec && /не актуал|утрат/i.test(rec.status+' '+rec.change)) addIssue({categoryId:'PM03',vndId:rec.vndId,documentTitle:rec.title,sourceFile:p.files[0]?.name||p.containerName,page:null,point:'',snippet:rec.change||rec.status,issue:'Реестр содержит признак утраты актуальности или изменения документа.',recommendation:'Проверить действующую редакцию и статус ВНД; при необходимости исключить устаревшую редакцию из действующего массива.',confidence:.9,severity:'high'});
    for(const u of p.units){ analyzeUnit(u,p,pointSets.get(p.key)); for(const ref of extractLegalRefs(u)){const k=norm(ref.raw);let ex=S.legalRefs.find(x=>x.key===k);if(!ex){ex={id:`LR-${String(S.legalRefs.length+1).padStart(4,'0')}`,key:k,raw:ref.raw,status:'не проверено',occurrences:[]};S.legalRefs.push(ex);}ex.occurrences.push({vndId:u.vndId,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text.slice(0,280)});}
    }
  }
  compareUnits(all); S.duplicates=S.issues.filter(i=>i.categoryId==='PM05'); S.conflicts=S.issues.filter(i=>i.categoryId==='PM02');
}
function analyzeUnit(u,p,pointSet){ const low=norm(u.text);
  for(const pat of S.rules.blanketPatterns||[]) if(low.includes(norm(pat))){addIssue({categoryId:'PM07',vndId:u.vndId,documentTitle:p.match?.title||p.containerName,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text,issue:`Обнаружена общая отсылочная формулировка «${pat}» без автоматически установленного конкретного источника.`,recommendation:'Уточнить необходимость отсылки и, если требуется однозначное применение, указать конкретный документ/норму.',confidence:.78});break;}
  for(const pat of S.rules.discretionPatterns||[]) if(low.includes(norm(pat))){addIssue({categoryId:'PM04',vndId:u.vndId,documentTitle:p.match?.title||p.containerName,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text,issue:`Выявлен потенциальный фактор широкого усмотрения: «${pat}».`,recommendation:'Проверить наличие объективных критериев, оснований, сроков и механизма контроля. Требуется правовая/антикоррупционная оценка специалистом.',confidence:.55,severity:'medium'});break;}
  for(const marker of S.rules.legacyMarkers||[]) if(low.includes(norm(marker))){addIssue({categoryId:'PM03',vndId:u.vndId,documentTitle:p.match?.title||p.containerName,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text,issue:`В тексте обнаружен исторический/устаревающий маркер «${marker}».`,recommendation:'Проверить актуальность наименования организации, системы или процесса и необходимость актуализации формулировки.',confidence:.55});break;}
  const refs=[...u.text.matchAll(/(?:пункт(?:а|у|ом|е)?|п\.)\s*(\d+(?:\.\d+)*)/giu)].map(m=>m[1]); if(/настоящ(?:их|его|ей)\s+(?:правил|положения|инструкции|регламента)/iu.test(u.text)){for(const r of refs)if(pointSet&&!pointSet.has(r)){addIssue({categoryId:'PM01',vndId:u.vndId,documentTitle:p.match?.title||p.containerName,sourceFile:u.sourceFile,page:u.page,point:u.point,snippet:u.text,issue:`Внутренняя ссылка на пункт ${r} не найдена среди распознанных пунктов этого документа.`,recommendation:'Проверить нумерацию и целевой пункт ссылки в утверждённом экземпляре.',confidence:.72,severity:'high'});}}
}
function extractLegalRefs(u){
  const out=[]; const t=u.text;
  const patterns=[
    /(?:ст(?:атья|атьи|атье|атью|атьей|\.)\s*\d+(?:[-.]\d+)?(?:\s*(?:част(?:ь|и|ью)|ч\.)\s*\d+)?\s+)?(?:АППК\s*РК|Административн(?:ый|ого)\s+процедурно-процессуальн(?:ый|ого)\s+кодекс(?:а)?\s+Республики Казахстан|Трудов(?:ой|ого)\s+кодекс(?:а)?\s*(?:РК|Республики Казахстан)?|Гражданск(?:ий|ого)\s+кодекс(?:а)?\s*(?:РК|Республики Казахстан)?|Кодекс[^.;\n]{0,90}?Республики Казахстан|Закон(?:а)?\s+Республики Казахстан\s+[«"][^»"]+[»"])(?:[^.;\n]{0,120})?/giu,
    /(?:АППК\s*РК|Административн(?:ый|ого)\s+процедурно-процессуальн(?:ый|ого)\s+кодекс(?:а)?\s+Республики Казахстан|Трудов(?:ой|ого)\s+кодекс(?:а)?\s*(?:РК|Республики Казахстан)?|Гражданск(?:ий|ого)\s+кодекс(?:а)?\s*(?:РК|Республики Казахстан)?|Кодекс[^.;\n]{0,90}?Республики Казахстан|Закон(?:а)?\s+Республики Казахстан\s+[«"][^»"]+[»"])(?:[^.;\n]{0,80})?ст(?:атья|атьи|атье|атью|атьей|\.)\s*\d+(?:[-.]\d+)?(?:\s*(?:част(?:ь|и|ью)|ч\.)\s*\d+)?(?:[^.;\n]{0,60})?/giu,
    /(?:пункт(?:а|у|ом|е)?|п\.)?\s*\d*(?:\.\d+)*\s*(?:приказ(?:а|ом|у)?)[^.;\n]{0,120}?№\s*[\w\-/]+(?:[^.;\n]{0,80})?/giu,
    /(?:пункт(?:а|у|ом|е)?|п\.)?\s*\d*(?:\.\d+)*\s*(?:постановлен(?:ие|ия))[^.;\n]{0,120}?№\s*[\w\-/]+(?:[^.;\n]{0,80})?/giu
  ];
  const seen=new Set(); for(const rx of patterns) for(const m of t.matchAll(rx)){const raw=m[0].trim().replace(/\s+/g,' ');const k=norm(raw);if(raw.length>3&&!seen.has(k)){seen.add(k);out.push({raw});}}
  return out;
}
function tokenJaccard(a,b){ const A=new Set(a),B=new Set(b);if(!A.size||!B.size)return 0;let n=0;for(const x of A)if(B.has(x))n++;return n/(A.size+B.size-n); }
function compareUnits(all){
  const min=S.cfg.analysis.exactDuplicateMinChars; const exact=new Map(); for(const u of all){if(u.normText.length<min)continue;const k=u.normText;if(!exact.has(k))exact.set(k,[]);exact.get(k).push(u);} for(const arr of exact.values())if(arr.length>1){for(let i=1;i<arr.length;i++)addDuplicate(arr[0],arr[i],1,'Точное совпадение текста нормы');}
  const inv=new Map(); for(let i=0;i<all.length;i++){const u=all[i]; if(u.normText.length<min||u.tokens.length<5)continue; const keys=[...u.tokens].sort((a,b)=>b.length-a.length).slice(0,8); for(const k of keys){if(!inv.has(k))inv.set(k,[]);inv.get(k).push(i);}}
  const counts=new Map(); let candidates=0; candidateLoop: for(const list of inv.values()){if(list.length>80)continue;for(let a=0;a<list.length;a++)for(let b=a+1;b<list.length;b++){if(candidates>=S.cfg.analysis.maxCandidatePairs)break candidateLoop;const i=list[a],j=list[b];if(all[i].packageKey===all[j].packageKey && all[i].sourceFile===all[j].sourceFile)continue;const key=i<j?`${i}|${j}`:`${j}|${i}`;counts.set(key,(counts.get(key)||0)+1);candidates++;}}
  for(const [key,c] of counts){if(c<2)continue;const [i,j]=key.split('|').map(Number),a=all[i],b=all[j];if(a.normText===b.normText)continue;const sim=tokenJaccard(a.tokens,b.tokens); if(sim>=S.cfg.analysis.conflictSimilarityThreshold){const reason=conflictReason(a.text,b.text);if(reason){addConflict(a,b,sim,reason);continue;}} if(sim>=S.cfg.analysis.nearDuplicateThreshold)addDuplicate(a,b,sim,'Существенное смысловое/лексическое пересечение');}
}
function addDuplicate(a,b,sim,reason){ const pa=S.packages.find(p=>p.key===a.packageKey),pb=S.packages.find(p=>p.key===b.packageKey);addIssue({categoryId:'PM05',vndId:a.vndId,otherVndId:b.vndId,documentTitle:pa?.match?.title||pa?.containerName||a.vndId,otherDocumentTitle:pb?.match?.title||pb?.containerName||b.vndId,sourceFile:a.sourceFile,otherSourceFile:b.sourceFile,page:a.page,otherPage:b.page,point:a.point,otherPoint:b.point,snippet:a.text,otherSnippet:b.text,similarity:sim,issue:reason,recommendation:'Проверить, является ли повторение необходимой детализацией. При конкурирующей компетенции — унифицировать, разграничить или исключить дубль.',confidence:Math.min(.96,.55+sim*.4)}); }
function conflictReason(a,b){ const na=norm(a),nb=norm(b); const deadline=/\b(\d{1,3})\s*(?:рабочих\s*)?(дн|дня|дней|час)/giu; const A=[...a.matchAll(deadline)].map(m=>m[1]),B=[...b.matchAll(deadline)].map(m=>m[1]); if(A.length&&B.length&&A[0]!==B[0])return`Похожие нормы содержат разные сроки (${A[0]} и ${B[0]}).`;
  const must=x=>/\b(обязан|должен|подлежит)\b/iu.test(x),may=x=>/\b(может|вправе)\b/iu.test(x); if((must(a)&&may(b))||(must(b)&&may(a)))return'Различается юридическая/управленческая модальность: обязанность сопоставлена с правом/возможностью.';
  if((/\bне\b/u.test(na))!==(/\bне\b/u.test(nb)))return'В близких по смыслу нормах обнаружено различие по отрицанию, которое может менять результат.';
  for(const pair of S.rules.nonEquivalentTerms||[]){const [x,y]=pair.map(norm);if((na.includes(x)&&nb.includes(y))||(na.includes(y)&&nb.includes(x)))return`Используются функционально различающиеся понятия «${pair[0]}» и «${pair[1]}».`;} return''; }
function addConflict(a,b,sim,reason){const pa=S.packages.find(p=>p.key===a.packageKey),pb=S.packages.find(p=>p.key===b.packageKey);addIssue({categoryId:'PM02',vndId:a.vndId,otherVndId:b.vndId,documentTitle:pa?.match?.title||pa?.containerName||a.vndId,otherDocumentTitle:pb?.match?.title||pb?.containerName||b.vndId,sourceFile:a.sourceFile,otherSourceFile:b.sourceFile,page:a.page,otherPage:b.page,point:a.point,otherPoint:b.point,snippet:a.text,otherSnippet:b.text,similarity:sim,issue:reason,recommendation:'Сопоставить полномочия, условия, сроки и результат. При подтверждении конфликта определить основной источник нормы и привести связанные ВНД в соответствие.',confidence:Math.min(.9,.48+sim*.45),severity:'high'});}

function buildRecommendations(){ const map=new Map(); const recText={PM01:'Устранить неработающие внутренние ссылки и уточнить механизм исполнения.',PM02:'Устранить подтверждённые противоречия и определить единый источник нормы.',PM03:'Актуализировать устаревшие нормы, реквизиты, названия и редакции.',PM04:'Проверить дискреционные формулировки и установить объективные критерии/контроль.',PM05:'Унифицировать или разграничить дублирующиеся функции и нормы.',PM06:'Заполнить выявленные пробелы регулирования и ответственности.',PM07:'Сократить неоднозначные отсылки и указать конкретные основания.',PM08:'Устранить иные структурные/терминологические недостатки.'};
  for(const i of S.issues){const r=S.registry.find(x=>x.vndId===i.vndId);const dep=r?.department||'Не определено';const key=`${dep}|${i.categoryId}`;if(!map.has(key))map.set(key,{id:`REC-${map.size+1}`,department:dep,categoryId:i.categoryId,title:recText[i.categoryId]||'Провести актуализацию.',issueIds:[],vndIds:new Set(),priority:i.severity==='high'?'high':'medium'});const x=map.get(key);x.issueIds.push(i.id);x.vndIds.add(i.vndId);if(i.severity==='high')x.priority='high';}
  S.recommendations=[...map.values()].map(x=>({...x,vndIds:[...x.vndIds]})); }
function buildQuality(){ const q=[...S.ingestErrors]; for(const r of S.registry)if(r.idTemporary)q.push({level:'warn',type:'VND_ID',message:`${r.title}: отсутствует постоянный VND_ID в реестре.`});
  for(const p of S.packages){if(!p.match)q.push({level:'error',type:'Сопоставление',message:`${p.containerName}: пакет не связан однозначно с реестром.`});if(!p.files.length)q.push({level:'error',type:'Состав пакета',message:`${p.containerName}: архив не содержит поддерживаемых документов.`});if(p.files.length&&!p.primaryFiles?.length)q.push({level:'warn',type:'Основной источник',message:`${p.containerName}: основной источник не определён автоматически. Откройте пакет и выберите файл вручную.`});else if(p.primaryFiles?.length&&p.primaryConfidence<.72)q.push({level:'warn',type:'Основной источник',message:`${p.containerName}: основной источник определён с низкой уверенностью (${Math.round(p.primaryConfidence*100)}%). Рекомендуется проверить выбор.`});for(const f of p.files){if(f.e==='doc'&&f.parseStatus==='legacy-doc-partial')q.push({level:'warn',type:'DOC',message:`${f.name}: текст старого DOC извлечён локально методом best-effort (${f.legacyDocMethod||'binary'}); критические выводы рекомендуется сверять с оригиналом.`});if(f.e==='doc'&&f.parseStatus==='doc-unreadable')q.push({level:'error',type:'DOC',message:`${f.name}: старый DOC не удалось надёжно прочитать локально. Используйте исходный просмотр/конвертацию только для этого файла.`});if(f.parseStatus==='ocr-pending')q.push({level:'warn',type:'OCR',message:`${f.name}: часть страниц ожидает OCR. Нажмите «OCR: распознать ожидающие» — они будут обработаны локальной очередью с прогрессом.`});const oc=f.pages?.find(x=>x.ocrConfidence!=null)?.ocrConfidence;if(oc!=null&&oc<60)q.push({level:'warn',type:'OCR',message:`${f.name}: низкая уверенность OCR (${Math.round(oc)}%); выводы по этому фрагменту требуют проверки оригинала.`});if(f.parseStatus==='error')q.push({level:'error',type:'Чтение файла',message:`${f.name}: ${f.parseError}`});}}
  const dup=Object.entries(S.packages.reduce((m,p)=>(p.vndId&&(m[p.vndId]=(m[p.vndId]||0)+1),m),{})).filter(([,n])=>n>1);for(const [id,n] of dup)q.push({level:'error',type:'Дубли пакетов',message:`${id}: загружено ${n} пакета.`});S.quality=q; }

function renderMonitoring(){ renderCategoryCards('#monitoringCategories'); const arr=S.issueFilter?S.issues.filter(i=>i.categoryId===S.issueFilter):S.issues;$('#issuesMeta').textContent=S.issueFilter?`${categoryName(S.issueFilter)} — ${arr.length}`:`Всего ${arr.length}`;$('#issuesTable').innerHTML=arr.length?issueTable(arr):'<div class="empty"><strong>Замечаний нет</strong>Загрузите документы или измените фильтр.</div>';$('#issuesTable').querySelectorAll('[data-issue]').forEach(x=>x.onclick=()=>openIssue(x.dataset.issue));}
function issueTable(arr,compact=false){return `<div class="table-wrap"><table class="data-table"><thead><tr><th>ID</th><th>Категория</th><th>ВНД / место</th><th>Суть</th>${compact?'':'<th>Уверенность</th><th>Статус</th>'}</tr></thead><tbody>${arr.map(i=>`<tr class="clickable" data-issue="${i.id}"><td class="mono">${i.id}</td><td><span class="badge ${i.categoryId==='PM02'?'danger':i.categoryId==='PM05'?'warn':'info'}">${esc(categoryName(i.categoryId))}</span></td><td>${esc(i.vndId)}<br><small>${esc(i.page?`стр. ${i.page}`:'')}${i.point?`, п. ${esc(i.point)}`:''}</small></td><td><div class="truncate" title="${esc(i.issue)}">${esc(i.issue)}</div></td>${compact?'':`<td>${Math.round((i.confidence||0)*100)}%</td><td>${reviewBadge(i.reviewStatus)}</td>`}</tr>`).join('')}</tbody></table></div>`;}
function reviewBadge(s){return s==='confirmed'?'<span class="badge ok">подтверждено</span>':s==='rejected'?'<span class="badge dark">отклонено</span>':s==='legal'?'<span class="badge purple">юридическая проверка</span>':'<span class="badge warn">выявлено ИИ</span>';}
function renderDuplicates(){const a=S.duplicates;$('#duplicatesTable').innerHTML=a.length?issueTable(a):'<div class="empty"><strong>Дубли не выявлены</strong>Или документы ещё не анализировались.</div>';$('#duplicatesTable').querySelectorAll('[data-issue]').forEach(x=>x.onclick=()=>openIssue(x.dataset.issue));}
function renderConflicts(){const a=S.conflicts;$('#conflictsTable').innerHTML=a.length?issueTable(a):'<div class="empty"><strong>Потенциальные противоречия не выявлены</strong>Или документы ещё не анализировались.</div>';$('#conflictsTable').querySelectorAll('[data-issue]').forEach(x=>x.onclick=()=>openIssue(x.dataset.issue));}
function renderLegalRefs(){ const rows=S.legalRefs;$('#legalRefsTable').innerHTML=rows.length?`<div class="table-wrap"><table class="data-table"><thead><tr><th>ID</th><th>Нормативная ссылка</th><th>Упоминаний</th><th>Статус проверки</th><th>Источник</th></tr></thead><tbody>${rows.map(r=>`<tr><td class="mono">${r.id}</td><td>${esc(r.raw)}</td><td>${r.occurrences.length}</td><td><span class="badge ${r.status==='подтверждено'?'ok':'warn'}">${esc(r.status)}</span></td><td>${(S.legalSources?.officialSources||[]).map(s=>`<a href="${s.url}" target="_blank" rel="noopener">${esc(s.name)}</a>`).join('<br>')}</td></tr>`).join('')}</tbody></table></div>`:'<div class="empty"><strong>Ссылки на НПА не извлечены</strong>После анализа здесь появятся найденные законы, кодексы, приказы и постановления.</div>';}
async function verifyLegal(){ if(!S.legalRefs.length){toast('Сначала проанализируйте документы','warn');return;}const ep=S.cfg.legalVerification.endpoint;if(!ep){toast('Защищённый Legal Verification endpoint не настроен. Полный текст ВНД в Интернет не отправляется.','warn');return;} try{const payload={references:S.legalRefs.map(r=>({id:r.id,raw:r.raw}))};const res=await fetch(ep,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});if(!res.ok)throw new Error(`HTTP ${res.status}`);const data=await res.json();for(const x of data.results||[]){const r=S.legalRefs.find(z=>z.id===x.id);if(r){r.status=x.status||'проверено';r.details=x;}}renderLegalRefs();toast('Проверка НПА завершена','ok');}catch(e){toast(`Ошибка внешней проверки: ${e.message}`,'error');}}

function renderDepartments(){ const depNames=uniq(S.registry.map(r=>r.department)).sort((a,b)=>a.localeCompare(b,'ru')); $('#departmentsGrid').innerHTML=depNames.map(dep=>{const docs=S.registry.filter(r=>r.department===dep),ids=new Set(docs.map(r=>r.vndId)),iss=S.issues.filter(i=>ids.has(i.vndId)),recs=S.recommendations.filter(r=>r.department===dep);return `<div class="panel clickable" data-dep="${esc(dep)}"><div class="panel-title"><h4>${esc(dep)}</h4><span class="meta">${docs.length} ВНД</span></div><div class="stat-line"><div class="stat-pill"><b>${iss.length}</b>замечаний</div><div class="stat-pill"><b>${recs.length}</b>рекомендаций</div><div class="stat-pill"><b>${iss.filter(i=>i.categoryId==='PM02').length}</b>противоречий</div><div class="stat-pill"><b>${iss.filter(i=>i.categoryId==='PM05').length}</b>дублей</div></div></div>`}).join('')||'<div class="empty">Структура не загружена.</div>';$('#departmentsGrid').querySelectorAll('[data-dep]').forEach(x=>x.onclick=()=>openDepartment(x.dataset.dep));}
function renderQuality(){const err=S.quality.filter(x=>x.level==='error').length,warn=S.quality.filter(x=>x.level==='warn').length,unmatched=S.packages.filter(p=>!p.match).length,missing=S.registry.length-loadedRegistryIds().size;const cards=[['Ошибки',err,'danger'],['Предупреждения',warn,'warn'],['Не сопоставлено пакетов',unmatched,unmatched?'danger':'ok'],['ВНД реестра не загружено',Math.max(0,missing),''],['Временных ID',S.registry.filter(r=>r.idTemporary).length,'warn'],['OCR ожидает (страниц)',pendingOcrTasks().length,'warn']];$('#qualityKpis').innerHTML=cards.map(c=>`<div class="kpi-card ${c[2]}"><div class="label">${c[0]}</div><div class="value">${c[1]}</div></div>`).join('');$('#qualityTable').innerHTML=S.quality.length?`<div class="table-wrap"><table class="data-table"><thead><tr><th>Уровень</th><th>Тип</th><th>Сообщение</th></tr></thead><tbody>${S.quality.map(q=>`<tr><td>${q.level==='error'?'<span class="badge danger">ошибка</span>':'<span class="badge warn">предупреждение</span>'}</td><td>${esc(q.type)}</td><td>${esc(q.message)}</td></tr>`).join('')}</tbody></table></div>`:'<div class="empty"><strong>Критических проблем данных нет</strong></div>';}
function renderReports(){const confirmed=S.issues.filter(i=>i.reviewStatus==='confirmed').length;$('#reportPreview').innerHTML=`<div class="stat-line"><div class="stat-pill"><b>${S.packages.length}</b>ВНД загружено</div><div class="stat-pill"><b>${S.packages.filter(p=>p.parsed).length}</b>проанализировано</div><div class="stat-pill"><b>${S.issues.length}</b>замечаний-кандидатов</div><div class="stat-pill"><b>${confirmed}</b>подтверждено</div><div class="stat-pill"><b>${S.recommendations.length}</b>рекомендаций СП</div></div><div class="footer-note">Методика: ${esc(S.version?.methodology||'NORMATRIX')}. Выводы локального движка требуют профессиональной проверки перед использованием как официального правового заключения.</div>`;}

function openModal(title){$('#modalTitle').textContent=title;$('#detailModal').classList.add('open');}
function closeModal(){$('#detailModal').classList.remove('open');}
function openVnd(id){const r=S.registry.find(x=>x.vndId===id);if(!r)return;const p=packageFor(id),iss=issuesFor(id),refs=S.legalRefs.filter(x=>x.occurrences.some(o=>o.vndId===id));openModal(`${id} — ${r.title}`);$('#modalSide').innerHTML=`<div class="detail-list"><div class="detail-item"><div class="k">Подразделение</div><div class="v">${esc(r.department)}</div></div><div class="detail-item"><div class="k">Ответственный</div><div class="v">${esc(r.responsible)}</div></div><div class="detail-item"><div class="k">№ / дата</div><div class="v">${esc(r.number)} · ${esc(fmtDate(r.date))}</div></div><div class="detail-item"><div class="k">Статус</div><div class="v">${esc(statusDerived(r))}</div></div></div>${p?'<hr>'+p.files.map(f=>`<div class="file-row ${f.isPrimary?'active':''}" data-view-file="${esc(f.name)}">${f.isPrimary?'<span class="badge ok">основной</span> ':''}${esc(f.name)}<br><small>${esc(roleLabel(f.role))} ${esc(f.lang)}</small></div>`).join(''):''}`;
  $('#modalMain').innerHTML=`<div class="grid kpi" style="grid-template-columns:repeat(4,1fr)"><div class="kpi-card"><div class="label">Пунктов/фрагментов</div><div class="value">${p?.units?.length||0}</div></div><div class="kpi-card warn"><div class="label">Замечаний</div><div class="value">${iss.length}</div></div><div class="kpi-card"><div class="label">НПА</div><div class="value">${refs.length}</div></div><div class="kpi-card"><div class="label">Файлов</div><div class="value">${p?.files?.length||0}</div></div></div><h4>Замечания</h4>${iss.length?issueTable(iss,true):'<div class="empty">Замечаний по этому ВНД нет.</div>'}<h4>Извлечённые нормативные ссылки</h4>${refs.length?refs.map(x=>`<div class="detail-item">${esc(x.raw)} — <span class="badge warn">${esc(x.status)}</span></div>`).join(''):'<div class="empty">Не найдено.</div>'}`;
  $('#modalMain').querySelectorAll('[data-issue]').forEach(x=>x.onclick=()=>openIssue(x.dataset.issue));$('#modalSide').querySelectorAll('[data-view-file]').forEach(x=>x.onclick=()=>viewFile(p,x.dataset.viewFile,null,''));}
function openPackage(key){const p=S.packages.find(x=>x.key===key);if(!p)return;openModal(p.containerName);$('#modalSide').innerHTML=p.files.map(f=>`<div class="file-row ${f.isPrimary?'active':''}" data-view-file="${esc(f.name)}">${f.isPrimary?'<span class="badge ok">основной</span> ':''}${esc(f.name)}<br><small>${esc(roleLabel(f.role))} · ${esc(f.parseStatus)}</small><br><button class="btn sm" data-set-primary="${esc(f.name)}">Сделать основным</button></div>`).join('');$('#modalMain').innerHTML=`<h3>${esc(p.match?.title||'Пакет не сопоставлен')}</h3><div class="detail-list"><div class="detail-item"><div class="k">VND_ID</div><div class="v mono">${esc(p.match?.vndId||p.vndId||'—')}</div></div><div class="detail-item"><div class="k">Точность сопоставления</div><div class="v">${Math.round(p.matchConfidence*100)}%</div></div><div class="detail-item"><div class="k">Основной источник</div><div class="v">${p.primaryFiles?.length?`${esc(p.primaryFiles.join(', '))}<br><small>${p.primarySource==='manual'?'выбран вручную':p.primarySource==='explicit'?'помечен MAIN':'определён автоматически'} · ${Math.round(p.primaryConfidence*100)}%</small>`:'Не определён'}</div></div><div class="detail-item"><div class="k">Извлечено норм/фрагментов</div><div class="v">${p.units.length}</div></div></div><div class="callout"><b>MAIN больше не обязателен.</b><br>Портал определяет основной источник по имени, типу, содержимому и совпадению с карточкой ВНД. Если выбор неверен, используйте кнопку «Сделать основным».</div>`;$('#modalSide').querySelectorAll('[data-view-file]').forEach(x=>x.onclick=e=>{if(e.target.closest('[data-set-primary]'))return;viewFile(p,x.dataset.viewFile,null,'');});$('#modalSide').querySelectorAll('[data-set-primary]').forEach(x=>x.onclick=e=>{e.stopPropagation();setPrimaryManual(p,x.dataset.setPrimary);});}
function openIssue(id){
  const i=S.issues.find(x=>x.id===id); if(!i)return;
  openModal(`${i.id} — ${categoryName(i.categoryId)}`);
  $('#modalSide').innerHTML=`<div class="detail-item"><div class="k">Категория</div><div class="v">${esc(categoryName(i.categoryId))}</div></div><div class="detail-item"><div class="k">Уверенность</div><div class="v">${Math.round((i.confidence||0)*100)}%</div></div><div class="detail-item"><div class="k">Статус</div><div class="v">${reviewBadge(i.reviewStatus)}</div></div><hr><button class="btn sm primary" data-review="confirmed">Подтвердить</button> <button class="btn sm" data-review="legal">Юр. проверка</button> <button class="btn sm" data-review="rejected">Отклонить</button>`;
  const sourceA=`<div class="norm-box"><div class="source">${esc(i.documentTitle||i.vndId)} · ${esc(i.sourceFile||'')} ${i.page?`· стр. ${i.page}`:''} ${i.point?`· п. ${esc(i.point)}`:''}</div><div class="text">${esc(i.snippet||'')}</div><br><button class="btn sm" data-source="a">Показать в документе</button></div>`;
  const sourceB=i.otherSnippet
    ? `<div class="norm-box"><div class="source">${esc(i.otherDocumentTitle||i.otherVndId||'Сопоставляемая норма')} · ${esc(i.otherSourceFile||'')} ${i.otherPage?`· стр. ${i.otherPage}`:''} ${i.otherPoint?`· п. ${esc(i.otherPoint)}`:''}</div><div class="text">${esc(i.otherSnippet)}</div><br><button class="btn sm" data-source="b">Показать в документе</button></div>`
    : `<div class="norm-box"><div class="source">Сопоставляемый источник</div><div class="text">${esc(i.reference||'Не требуется / не установлен автоматически')}</div></div>`;
  const sim=i.similarity!=null?`<div class="detail-item"><div class="k">Смысловое/лексическое сходство</div><div class="v">${Math.round(i.similarity*100)}%</div></div>`:'';
  $('#modalMain').innerHTML=`<div class="issue-head">${sourceA}${sourceB}</div><div class="callout ${i.categoryId==='PM02'?'danger':'warn'}"><b>Что обнаружено:</b><br>${esc(i.issue)}</div><div class="recommendation ${i.severity==='high'?'high':'medium'}"><div class="rec-title">Рекомендация</div><div>${esc(i.recommendation)}</div></div>${sim}`;
  $('#modalSide').querySelectorAll('[data-review]').forEach(b=>b.onclick=()=>{i.reviewStatus=b.dataset.review;buildRecommendations();renderAll();openIssue(id);});
  $('#modalMain').querySelector('[data-source="a"]')?.addEventListener('click',()=>openIssueSource(i,false));
  $('#modalMain').querySelector('[data-source="b"]')?.addEventListener('click',()=>openIssueSource(i,true));
}
function openIssueSource(i,other){const vid=other?i.otherVndId:i.vndId,fileName=other?i.otherSourceFile:i.sourceFile,page=other?i.otherPage:i.page,snip=other?i.otherSnippet:i.snippet;const p=packageFor(vid);if(!p){toast('Исходный файл не найден в текущей сессии','warn');return;}viewFile(p,fileName,page,snip);}
function viewFile(p,fileName,page,snippet){const f=p.files.find(x=>x.name===fileName)||p.files[0];if(!f)return;$('#modalTitle').textContent=`${p.match?.vndId||p.vndId||''} — ${f.name}`;$('#modalSide').innerHTML=p.files.map(x=>`<div class="file-row ${x===f?'active':''}" data-view-file="${esc(x.name)}">${esc(x.name)}<br><small>${esc(x.role)} · ${esc(x.parseStatus)}</small></div>`).join('');let body=snippet?`<div class="snippet"><b>Найденный фрагмент${page?` · стр. ${page}`:''}:</b><br>${esc(snippet)}</div>`:'';if(f.e==='pdf')body+=`<iframe class="viewer-frame" src="${f.objectUrl}${page?`#page=${page}`:''}"></iframe>`;else if(['jpg','jpeg','png','tif','tiff','webp'].includes(f.e))body+=`<img class="viewer-img" src="${f.objectUrl}" alt="${esc(f.name)}">`;else if(f.e==='docx'&&f.html)body+=`<div class="viewer-html">${highlightHtml(f.html,snippet)}</div>`;else body+=`<pre class="viewer-text">${esc(f.text||'Предпросмотр текста недоступен для этого формата.')}</pre>`;$('#modalMain').innerHTML=body;$('#modalSide').querySelectorAll('[data-view-file]').forEach(x=>x.onclick=()=>viewFile(p,x.dataset.viewFile,null,''));}
function highlightHtml(html,snip){if(!snip)return html;const needle=esc(snip.slice(0,70));return html.replace(new RegExp(needle.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'i'),m=>`<mark>${m}</mark>`);}
function openDepartment(dep){const docs=S.registry.filter(r=>r.department===dep),ids=new Set(docs.map(r=>r.vndId)),iss=S.issues.filter(i=>ids.has(i.vndId)),recs=S.recommendations.filter(r=>r.department===dep);openModal(dep);$('#modalSide').innerHTML=`<div class="detail-item"><div class="k">ВНД</div><div class="v">${docs.length}</div></div><div class="detail-item"><div class="k">Замечаний</div><div class="v">${iss.length}</div></div><div class="detail-item"><div class="k">Рекомендаций</div><div class="v">${recs.length}</div></div>`;$('#modalMain').innerHTML=`<h3>Рекомендации подразделению</h3>${recs.length?recs.map(r=>`<div class="recommendation ${r.priority}"><div class="rec-title">${esc(r.title)}</div><div class="rec-meta">${categoryName(r.categoryId)} · ${r.issueIds.length} замечаний · ${r.vndIds.length} ВНД</div><div style="margin-top:8px">${r.vndIds.map(id=>`<button class="btn sm" data-vnd="${id}">${id}</button>`).join(' ')}</div></div>`).join(''):'<div class="empty">По текущей сессии рекомендаций нет.</div>'}<h3>ВНД подразделения</h3>${docs.map(d=>`<div class="detail-item clickable" data-vnd="${d.vndId}"><b>${d.vndId}</b> — ${esc(d.title)}</div>`).join('')}`;$('#modalMain').querySelectorAll('[data-vnd]').forEach(x=>x.onclick=()=>openVnd(x.dataset.vnd));}

function exportXlsx(){if(!window.XLSX){toast('Библиотека XLSX недоступна','error');return;}const wb=XLSX.utils.book_new();const summary=[['Показатель','Значение'],['Дата анализа',new Date().toLocaleString('ru-RU')],['ВНД в реестре',S.registry.length],['Загружено ВНД',S.packages.length],['Проанализировано',S.packages.filter(p=>p.parsed).length],['Замечаний-кандидатов',S.issues.length],['Подтверждено специалистом',S.issues.filter(i=>i.reviewStatus==='confirmed').length],['Нормативных ссылок',S.legalRefs.length],['Методика',S.version?.methodology||'NORMATRIX']];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summary),'Свод');
  const app=[['№ п/п','Виды недостатков норм поведения','Структурные элементы документа (страница, абзац, подпункт, пункт, статья, часть и пр.)','Номер и наименование акта, которому не соответствует норма поведения документа','Предпринимаемые меры по устранению выявленных недостатков и противоречий','Сведения об актуализации документа']];S.issues.forEach((i,n)=>app.push([n+1,categoryName(i.categoryId),`${i.vndId}; ${i.page?`стр. ${i.page}; `:''}${i.point?`п. ${i.point}; `:''}${(i.snippet||'').slice(0,500)}`,i.otherDocumentTitle||i.reference||'',i.recommendation||'',i.reviewStatus==='confirmed'?'Подтверждено специалистом':'Не актуализирован / требует рассмотрения']));XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(app),'Приложение 3');
  const rr=[['Подразделение','Приоритет','Категория','Рекомендация','ВНД','Замечания']];S.recommendations.forEach(r=>rr.push([r.department,r.priority,categoryName(r.categoryId),r.title,r.vndIds.join(', '),r.issueIds.join(', ')]));XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rr),'Рекомендации СП');
  const lr=[['ID','НПА/ссылка','Упоминаний','Статус']];S.legalRefs.forEach(r=>lr.push([r.id,r.raw,r.occurrences.length,r.status]));XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(lr),'НПА');
  const q=[['Уровень','Тип','Сообщение'],...S.quality.map(x=>[x.level,x.type,x.message])];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(q),'Качество данных');XLSX.writeFile(wb,`NORMATRIX_VND_${new Date().toISOString().slice(0,10)}.xlsx`);}
function exportJson(){const data={generatedAt:nowIso(),version:S.version,coverage:{registry:S.registry.length,loaded:S.packages.length},issues:S.issues.map(({_key,...x})=>x),legalRefs:S.legalRefs,recommendations:S.recommendations,quality:S.quality};downloadBlob(new Blob([JSON.stringify(data,null,2)],{type:'application/json'}),`NORMATRIX_VND_${new Date().toISOString().slice(0,10)}.json`);}
function downloadBlob(blob,name){const u=URL.createObjectURL(blob);const a=document.createElement('a');a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),1000);}
function printReport(){const w=window.open('','_blank');if(!w)return;w.document.write(`<html><head><meta charset="utf-8"><title>NORMATRIX report</title><style>body{font-family:Arial;padding:30px}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccc;padding:6px;font-size:12px}h1{color:#17324d}</style></head><body><h1>NORMATRIX VND Analytics</h1><p>Дата: ${new Date().toLocaleString('ru-RU')}</p><p>Загружено ВНД: ${S.packages.length}; замечаний: ${S.issues.length}; подтверждено: ${S.issues.filter(i=>i.reviewStatus==='confirmed').length}</p><h2>Замечания</h2><table><tr><th>ID</th><th>Категория</th><th>ВНД / пункт</th><th>Суть</th><th>Рекомендация</th></tr>${S.issues.map(i=>`<tr><td>${i.id}</td><td>${esc(categoryName(i.categoryId))}</td><td>${esc(i.vndId)} ${i.point?`п.${esc(i.point)}`:''}</td><td>${esc(i.issue)}</td><td>${esc(i.recommendation)}</td></tr>`).join('')}</table></body></html>`);w.document.close();w.focus();setTimeout(()=>w.print(),300);}
function clearSession(){if(!confirm('Удалить все загруженные ВНД и результаты текущей сессии? Корневой реестр и структура останутся.'))return;for(const u of S.objectUrls)URL.revokeObjectURL(u);S.objectUrls.clear();S.packages=[];S.issues=[];S.legalRefs=[];S.duplicates=[];S.conflicts=[];S.quality=[];S.ingestErrors=[];S.recommendations=[];S.issueFilter='';$('#fileInput').value='';$('#runAnalysisBtn').disabled=true;$('#runPendingOcrBtn').disabled=true;renderAll();goPage('dashboard');toast('Сессия очищена','ok');}

function bindStaticActions(){ $('#modalClose').onclick=closeModal;$('#detailModal').addEventListener('click',e=>{if(e.target.id==='detailModal')closeModal();});$('#clearSessionBtn').onclick=clearSession;$('#verifyLegalBtn').onclick=verifyLegal;$('#exportXlsxBtn').onclick=exportXlsx;$('#exportJsonBtn').onclick=exportJson;$('#printReportBtn').onclick=printReport;window.addEventListener('beforeunload',()=>{for(const u of S.objectUrls)URL.revokeObjectURL(u);}); }
function checkRuntimeDependencies(){
  const deps=[['SheetJS / XLSX',window.XLSX],['JSZip',window.JSZip],['Mammoth',window.mammoth],['PDF.js',window.pdfjsLib],['Tesseract OCR',window.Tesseract]]; const missing=deps.filter(([,v])=>!v).map(([n])=>n);
  if(!missing.length)return true;
  const msg=`Не загружены браузерные библиотеки: ${missing.join(', ')}. Проверьте доступ к CDN или разместите локальные копии библиотек в vendor/.`;
  setChip('#chipRegistry','Библиотеки недоступны',false); setChip('#chipStructure','Инициализация остановлена',false); $('#coveragePanel').innerHTML=`<div class="callout danger"><b>Портал не может начать работу.</b><br>${esc(msg)}</div>`; toast(msg,'error'); return false;
}

async function boot(){ initNav();setupUpload();bindStaticActions();if(!checkRuntimeDependencies())return;await loadStaticData(); }
document.addEventListener('DOMContentLoaded',boot);
})();
