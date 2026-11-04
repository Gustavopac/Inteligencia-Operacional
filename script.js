/* Configure os IDs públicos do aplicativo SPA Microsoft Entra ID. NÃO adicione segredos. */
const CFG={clientId:'1a653739-401d-4927-bef3-69d9a878e36a',tenant:'bc0cc947-c203-43f7-b793-da63a1ff5446',sharePointHost:'arrowtransportes.sharepoint.com',sharePointSite:'/sites/OPERACAO',sharePointFile:'PROGRAMAÇÃO GERAL.xlsx',sharePointItemId:'D02CF646-0EC0-4E9E-A05A-0B6BFF2C3592',sheet:'Geral',refreshMs:60000,tempoCol:'TEMPO'};
const $=id=>document.getElementById(id), esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toUpperCase();
const today=(offset=0)=>{let d=new Date();d.setDate(d.getDate()+offset);return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`};
let scheduleSelectedDate=today(1);let rows=[],page='overview',msalApp,account,threshold=1,charts=[],mode='demo';
const displayDate=s=>/^\d{4}-\d{2}-\d{2}$/.test(String(s||''))?`${s.slice(8,10)}/${s.slice(5,7)}/${s.slice(2,4)}`:String(s||'—');
const vehicleType=r=>r['TIPOLOGIA DO VEICULO']||r['TIPOLOGIA VEICULO']||r['TIPO DE VEICULO']||r['TIPO VEICULO']||r.TIPOLOGIA||'';
const fmt=n=>`${Math.round(n||0)} min`, avg=a=>a.length?a.reduce((x,y)=>x+y,0)/a.length:0,sd=a=>Math.sqrt(avg(a.map(x=>(x-avg(a))**2))),median=a=>{if(!a.length)return 0;let b=[...a].sort((x,y)=>x-y);return (b[Math.floor((b.length-1)/2)]+b[Math.floor(b.length/2)])/2};
function dateOf(v){if(v==null||v==='')return '';if(typeof v==='number')return new Date(Date.UTC(1899,11,30)+Math.floor(v)*86400000).toISOString().slice(0,10);let s=String(v).trim();if(/^\d{4}-\d\d-\d\d/.test(s))return s.slice(0,10);let m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);if(m)return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;let d=new Date(s);return isNaN(d)?'':`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function timeOf(v){if(v==null||v==='')return null;if(typeof v==='number')return (v%1)*1440;let m=String(v).match(/(\d{1,2}):(\d{2})/);return m?+m[1]*60+ +m[2]:null}
function duration(r){let v=r[norm(CFG.tempoCol)];if(v!==''&&v!=null){if(typeof v==='number')return v<1?v*1440:v;let m=String(v).match(/^(\d+):(\d{2})(?::(\d{2}))?$/);if(m)return +m[1]*60+ +m[2]+ +(m[3]||0)/60;let n=Number(String(v).replace(',','.'));if(Number.isFinite(n))return n}let a=timeOf(r.INICIO),b=timeOf(r.TERMINNO||r.TERMINO);return a==null||b==null?null:(b-a+1440)%1440}
function process(values,source='Excel'){if(!values?.length)throw Error('Arquivo sem dados');let headers=values[0].map(norm);if(!headers.includes('DATA')||!headers.includes('CLIENTE'))throw Error('A planilha precisa conter as colunas DATA e CLIENTE.');rows=values.slice(1).filter(v=>v.some(x=>x!==''&&x!=null)).map(v=>{let r={};headers.forEach((h,i)=>r[h]=v[i]??'');r.DATA_ISO=dateOf(r.DATA);r.MIN=duration(r);return r});mode=source;populate();setDefaultDates();render();$('notice').textContent=`${rows.length.toLocaleString('pt-BR')} registros carregados de ${source}. Os filtros atualizam todas as abas.`;$('sync').textContent=`Atualizado às ${new Date().toLocaleTimeString('pt-BR')}`;$('sourceLabel').textContent=source;$('statusDot').style.background=source==='Demonstração'?'#F46E00':'#F46E00'}
function setDefaultDates(){let dates=rows.map(r=>r.DATA_ISO).filter(Boolean).sort();if(dates.length){$('from').value=dates[0];$('to').value=dates[dates.length-1]}}
function populate(){for(let [id,key] of [['client','CLIENTE'],['operation','OPERACAO'],['material','MATERIAL'],['type','TIPOLOGIA']]){let el=$(id),previous=el.value,values=[...new Set(rows.map(r=>String(id==='type'?vehicleType(r):r[key]??'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR'));el.innerHTML='<option value="">Todos</option>'+values.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');if(values.includes(previous))el.value=previous}}
function filtered(ignoreDates=false){return rows.filter(r=>(ignoreDates||(!$('from').value||r.DATA_ISO >= $('from').value)&&(!$('to').value||r.DATA_ISO <= $('to').value))&&[['client','CLIENTE'],['operation','OPERACAO'],['material','MATERIAL'],['type','TIPOLOGIA']].every(([id,k])=>!$(id).value||String(id==='type'?vehicleType(r):r[k])===$(id).value))}
const isLoad=r=>['CARGA','CARREGAMENTO'].includes(norm(r.OPERACAO)),isUnload=r=>norm(r.OPERACAO)==='DESCARGA',isDone=r=>['CARREGADO','RECEBIDO','CONCLUIDO','FINALIZADO'].includes(norm(r.STATUS)),isScheduled=r=>['SIM','S','YES','TRUE','1'].includes(norm(r['AGENDADO?']));
const card=(label,value,sub='')=>`<div class="card"><small>${esc(label)}</small><strong>${esc(value)}</strong><em>${esc(sub)}</em></div>`;
const panel=(title,body,sub='')=>`<div class="panel"><div class="panelhead"><h2>${esc(title)}</h2><small>${esc(sub)}</small></div>${body}</div>`;
const chart=(id)=>`<div class="chartbox"><canvas id="${id}"></canvas></div>`;
const fields=[['Data',r=>displayDate(r.DATA_ISO)],['Cliente',r=>r.CLIENTE],['Operação',r=>r.OPERACAO],['Material',r=>r.MATERIAL],['Tipologia do veículo',r=>vehicleType(r)],['Status',r=>r.STATUS],['Tempo',r=>r.MIN==null?'—':fmt(r.MIN)]];
function table(data,cols=fields){return `<div class="panel"><div class="tablewrap"><table><thead><tr>${cols.map(f=>`<th>${esc(f[0])}</th>`).join('')}</tr></thead><tbody>${data.length?data.map(r=>`<tr data-index="${rows.indexOf(r)}">${cols.map(([name,fn])=>`<td>${esc(fn(r))}</td>`).join('')}</tr>`).join(''):`<tr><td colspan="${cols.length}">Nenhuma operação encontrada.</td></tr>`}</tbody></table></div></div>`}
function count(rs,key){let o={};rs.forEach(r=>{let k=String(r[key]||'Não informado');o[k]=(o[k]||0)+1});return Object.entries(o).sort((a,b)=>b[1]-a[1])}
function overview(rs){let valid=rs.map(r=>r.MIN).filter(x=>x!=null),days=[...new Set(rs.map(r=>r.DATA_ISO).filter(Boolean))].length;return `<div class="cards">${card('Operações no período',rs.length.toLocaleString('pt-BR'),'Operações')}${card('Operações concluídas',rs.filter(isDone).length,rs.length?(100*rs.filter(isDone).length/rs.length).toFixed(1)+'% do total':'Sem dados')}${card('Agendados',rs.filter(isScheduled).length,'')}${card('Tempo médio',fmt(avg(valid)),`${valid.length} registros com duração`)}${card('Cargas',rs.filter(isLoad).length,'Carregamentos')}${card('Descargas',rs.filter(isUnload).length,'Recebimentos')}${card('Clientes atendidos',new Set(rs.map(r=>r.CLIENTE).filter(Boolean)).size,'Clientes distintos')}${card('Operações por dia',days?(rs.length/days).toFixed(1):'—',`${days} dias com registros`)}</div><div class="two">${panel('Evolução diária',chart('trend'),'Quantidade por data')}${panel('Distribuição de status',chart('statusChart'),'Participação por status')}</div><div class="two">${panel('Top 10 clientes',chart('clientChart'),'Volume de operações')}${panel('Cargas × descargas',chart('operationChart'),'Operacional')}</div><div class="two">${panel('Tipologia do veículo',chart('typeChart'),'Composição')}${panel('Tempo médio por material',chart('materialChart'),'Minutos por operação')}</div><div class="toolbar"><h2>Últimas operações do período</h2></div>${table([...rs].sort((a,b)=>b.DATA_ISO.localeCompare(a.DATA_ISO)).slice(0,40))}`}
function daily(rs){let valid=rs.map(r=>r.MIN).filter(v=>v!=null);return `<div class="cards">${card('Total de operações',rs.length)}${card('Agendadas',rs.filter(isScheduled).length)}${card('Cargas',rs.filter(isLoad).length)}${card('Descargas',rs.filter(isUnload).length)}${card('Concluídas',rs.filter(isDone).length)}${card('Mediana carga',fmt(median(rs.filter(isLoad).map(r=>r.MIN).filter(x=>x!=null))))}${card('Mediana descarga',fmt(median(rs.filter(isUnload).map(r=>r.MIN).filter(x=>x!=null))))}${card('Tempo médio geral',fmt(avg(valid)))}</div><div class="two">${panel('Operações por status',chart('dailyStatus'))}</div><h2>Consulta detalhada das operações</h2>${table(rs)}`}
function variation(rs){let valid=rs.map(r=>r.MIN).filter(v=>v!=null),groups={};rs.forEach(r=>{if(r.MIN==null)return;let k=`${r.MATERIAL||'—'}`;(groups[k]??=[]).push(r.MIN)});let list=Object.entries(groups).map(([k,a])=>({k,n:a.length,mean:avg(a),std:sd(a)})).sort((a,b)=>b.n-a.n);return `<div class="cards">${card('Tempo médio',fmt(avg(valid)))}${card('Mediana',fmt(median(valid)))}${card('Desvio padrão',fmt(sd(valid)))}${card('Coeficiente de variação',valid.length&&avg(valid)?(sd(valid)/avg(valid)*100).toFixed(1)+'%':'—')}${card('Maior tempo',fmt(Math.max(0,...valid)))}${card('Menor tempo',valid.length?fmt(Math.min(...valid)):'—')}${card('Amostras com tempo',valid.length)}${card('Sem tempo informado',rs.length-valid.length)}</div>${panel('Tempo médio por material',chart('variationChart'),'Minutos')}<div class="panel"><h2>Comparativo por material</h2><div class="tablewrap"><table><thead><tr><th>Material</th><th>Operações</th><th>Média</th><th>Desvio padrão</th></tr></thead><tbody>${list.map(g=>`<tr><td>${esc(g.k)}</td><td>${g.n}</td><td>${fmt(g.mean)}</td><td>${fmt(g.std)}</td></tr>`).join('')}</tbody></table></div></div>`}
function outliers(rs){let groups={};rs.forEach(r=>{if(r.MIN!=null)(groups[`${r.OPERACAO}|${r.MATERIAL}`]??=[]).push(r.MIN)});let anomalies=rs.filter(r=>{let a=groups[`${r.OPERACAO}|${r.MATERIAL}`];return a&&a.length>1&&r.MIN!=null&&r.MIN>avg(a)+threshold*sd(a)});return `<div class="toolbar"><h2>Desvios operacionais</h2><label>Limite <select id="sigma">${[0,.5,1,1.5,2,3].map(n=>`<option value="${n}" ${threshold===n?'selected':''}>${n} desvio(s)</option>`).join('')}</select></label></div><div class="cards">${card('Fora da média',anomalies.length,'Acima do limite estatístico')}${card('Taxa de desvios',rs.length?(100*anomalies.length/rs.length).toFixed(1)+'%':'—')}${card('Maior duração',fmt(Math.max(0,...rs.map(r=>r.MIN||0))))}${card('Grupos analisados',Object.keys(groups).length)}</div>${panel('Operações por dia',chart('outlierTrend'),'Dias com volume abaixo da média em destaque')}<h2>Ocorrências identificadas</h2>${table(anomalies)}`}
function tomorrow(){let date=scheduleSelectedDate,rs=filtered(true).filter(r=>r.DATA_ISO===date);return `<div class="toolbar"><h2>Programação para ${displayDate(date)}</h2><label>Selecionar data <input type="date" id="scheduleDate" value="${date}"></label><button class="primary" id="print">Imprimir programação</button></div><h2 class="printtitle">ARROW — PROGRAMAÇÃO ${displayDate(date)}</h2><div class="cards">${card('Operações previstas',rs.length)}${card('Cargas',rs.filter(isLoad).length)}${card('Descargas',rs.filter(isUnload).length)}${card('Clientes',new Set(rs.map(r=>r.CLIENTE)).size)}</div>${table(rs,[['Cliente',r=>r.CLIENTE],['Operação',r=>r.OPERACAO],['Material',r=>r.MATERIAL],['Destino',r=>r.DESTINO],['Tipologia do veículo',r=>vehicleType(r)],['Status',r=>r.STATUS],['NF',r=>r.NF],['Observações',r=>r.OBS]])}`}
function shareSchedule(){
 const date=scheduleSelectedDate;
 const dates=[...new Set([...rows.map(r=>r.DATA_ISO).filter(Boolean),today(),today(1)])].sort().reverse();
 if(!dates.includes(date))dates.unshift(date);
 const data=filtered(true).filter(r=>r.DATA_ISO===date).sort((a,b)=>String(a.CLIENTE).localeCompare(String(b.CLIENTE),'pt-BR'));
 const cols=[['CLIENTE',r=>r.CLIENTE],['OPERAÇÃO',r=>r.OPERACAO],['MATERIAL',r=>r.MATERIAL],['DESTINO',r=>r.DESTINO],['TIPOLOGIA DO VEÍCULO',r=>vehicleType(r)],['STATUS',r=>r.STATUS]];
 const trs=data.length?data.map((r,i)=>`<tr><td>${i+1}</td>${cols.map(([,fn])=>`<td>${esc(fn(r)||'—')}</td>`).join('')}</tr>`).join(''):`<tr><td colspan="7" class="share-empty">Nenhuma operação programada para esta data.</td></tr>`;
 return `<div class="share-controls"><div class="share-controls-left"><label for="shareDate">Data da programação</label><select id="shareDate">${dates.map(d=>`<option value="${d}" ${d===date?'selected':''}>${displayDate(d)}</option>`).join('')}</select><input type="date" id="shareCustomDate" value="${date}" title="Escolher outra data"></div><div class="share-buttons"><button class="ghost" id="shareImage">↓ Baixar imagem PNG</button><button class="primary" id="sharePrint">Imprimir / PDF</button></div></div>
 <div class="share-sheet" id="shareSheet"><div class="share-sheet-head"><div class="share-logo"><img src="Assets/logo-oficial.png" alt="Arrow"></div><div class="share-head-copy"><div class="share-kicker">ARROW • LOGÍSTICA INTEGRADA</div><h2>PROGRAMAÇÃO OPERACIONAL</h2><p>Programação de cargas e descargas</p></div><div class="share-date"><small>DATA DA PROGRAMAÇÃO</small><strong>${displayDate(date)}</strong></div></div>
 <div class="share-metrics"><div class="share-metric featured"><small>OPERAÇÕES</small><strong>${data.length}</strong></div><div class="share-metric"><small>CARGAS</small><strong>${data.filter(isLoad).length}</strong></div><div class="share-metric"><small>DESCARGAS</small><strong>${data.filter(isUnload).length}</strong></div></div>
 <div class="share-tablewrap"><table class="share-table"><thead><tr><th>#</th>${cols.map(([n])=>`<th>${n}</th>`).join('')}</tr></thead><tbody>${trs}</tbody></table></div>
 <div class="share-foot"><span>ARROW TRANSPORTES E LOGÍSTICA</span><span>Fonte: ${esc(mode==='demo'?'DADOS DE DEMONSTRAÇÃO':mode)} • Gerado em ${new Date().toLocaleString('pt-BR')}</span></div></div><p class="share-hint">A imagem inclui apenas o quadro acima, sem menus ou filtros. Os dados refletem a fonte e a data selecionadas.</p>`;
}
async function downloadShareImage(){
 const node=$('shareSheet');const btn=$('shareImage');if(!node)return;
 if(!window.html2canvas){$('notice').textContent='Não foi possível carregar o gerador de imagem. Utilize Imprimir / PDF ou uma captura de tela.';return}
 btn.disabled=true;btn.textContent='Gerando imagem...';
 try{
  await document.fonts.ready;
  const canvas=await html2canvas(node,{backgroundColor:'#050505',scale:2,useCORS:true,scrollX:0,scrollY:-window.scrollY,windowWidth:Math.max(document.documentElement.clientWidth,node.scrollWidth)});
  const link=document.createElement('a');link.download=`ARROW-Programacao-${scheduleSelectedDate}.png`;link.href=canvas.toDataURL('image/png');link.click();
 }catch(e){$('notice').textContent='Falha ao gerar imagem: '+e.message}
 finally{btn.disabled=false;btn.textContent='↓ Baixar imagem PNG'}
}
function quality(rs){let fields=['DATA_ISO','CLIENTE','OPERACAO','MATERIAL','TIPOLOGIA','STATUS','NF','MIN','AGENDADO?'],countMissing=k=>rs.filter(r=>r[k]===null||r[k]===undefined||r[k]==='').length;return `<div class="cards">${card('Registros analisados',rs.length)}${card('Sem data',countMissing('DATA_ISO'))}${card('Sem duração',countMissing('MIN'))}${card('Sem status',countMissing('STATUS'))}</div>${panel('Completude dos campos',chart('qualityChart'),'Percentual de preenchimento')}<div class="panel"><h2>Campos auditados</h2><div class="qualitygrid">${fields.map(k=>`<div class="qualityitem"><span class="muted">${esc(k)}</span><b>${rs.length?(100*(rs.length-countMissing(k))/rs.length).toFixed(1):'0'}%</b><small class="muted">${countMissing(k)} não preenchidos</small></div>`).join('')}</div></div>`}
// Rótulos permanentes de valores, inclusive em gráficos de rosca.
const valueLabels={
 id:'arrowValueLabels',
 afterDatasetsDraw(chart){
  const {ctx}=chart;const ds=chart.data.datasets[0];if(!ds)return;
  ctx.save();ctx.font='bold 11px Segoe UI, Arial';ctx.fillStyle='#FFFFFF';
  chart.getDatasetMeta(0).data.forEach((element,i)=>{
   const val=Number(ds.data[i]);if(!Number.isFinite(val)||val===0)return;
   let pos=element.tooltipPosition();
   if(chart.config.type==='doughnut'){
    const total=ds.data.reduce((n,v)=>n+(Number(v)||0),0);
    if(total && val/total<.055)return; // categorias pequenas permanecem na legenda
    ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(String(val),pos.x,pos.y);
   }else if(chart.options.indexAxis==='y'){
    ctx.textAlign='left';ctx.textBaseline='middle';
    ctx.fillText(String(val),Math.min(pos.x+7,chart.chartArea.right-25),pos.y);
   }else{
    ctx.textAlign='center';ctx.textBaseline='bottom';
    ctx.fillText(String(val),pos.x,Math.max(chart.chartArea.top+14,pos.y-7));
   }
  });ctx.restore();
 }
};
function draw(id,type,labels,values,opts={}){
 const el=$(id);if(!el||!window.Chart)return;
 const colors=['#F46E00','#FF861F','#FF9C47','#FFB16C','#FFC38F','#FFD4B2','#E86200','#FFAA58'];
 const dates=!!opts.dates;
 const chartObj=new Chart(el,{
  type,
  data:{labels:dates?labels.map(displayDate):labels,datasets:[{
   label:opts.label||'Operações',data:values,
   backgroundColor:opts.singleColor?'#F46E00':type==='line'?'#F46E00':labels.map((_,i)=>colors[i%colors.length]),
   borderColor:'#F46E00',borderWidth:type==='line'?3:0,
   borderRadius:type==='bar'?5:0,fill:false,tension:.3,pointRadius:4,pointHoverRadius:6
  }]},
  plugins:[valueLabels],
  options:{
   responsive:true,maintainAspectRatio:false,indexAxis:opts.horizontal?'y':'x',
   layout:{padding:{top:18,right:opts.horizontal?35:12}},
   plugins:{legend:{display:type==='doughnut',position:'bottom',labels:{color:'#ddd',font:{size:11},padding:12}},tooltip:{enabled:true}},
   scales:type==='doughnut'?{}:{
    x:{ticks:{color:'#B8B8B8',maxRotation:45,autoSkip:true,maxTicksLimit:12},grid:{display:false}},
    y:{beginAtZero:true,ticks:{color:'#B8B8B8'},grid:{color:'#303030'}}
   }
  }
 });charts.push(chartObj);
}
function drawCharts(rs){
 let top=(key,n=10)=>count(rs,key).slice(0,n);
 let days=Object.entries(rs.reduce((o,r)=>(r.DATA_ISO&&(o[r.DATA_ISO]=(o[r.DATA_ISO]||0)+1),o),{})).sort(([a],[b])=>a.localeCompare(b));
 let materials={};rs.forEach(r=>{if(r.MIN!=null)(materials[r.MATERIAL||'Não informado']??=[]).push(r.MIN)});
 let times=Object.entries(materials).map(([k,v])=>[k,Math.round(avg(v))]).sort((a,b)=>b[1]-a[1]).slice(0,10);
 const make=(id,type,data,opts={})=>draw(id,type,data.map(x=>x[0]),data.map(x=>x[1]),opts);
 make('trend','line',days,{dates:true});
 make('statusChart','doughnut',top('STATUS',8));
 make('clientChart','bar',top('CLIENTE'),{horizontal:true});
 make('operationChart','doughnut',[['Cargas',rs.filter(isLoad).length],['Descargas',rs.filter(isUnload).length]]);
 let vehicles={};rs.forEach(r=>{let v=String(vehicleType(r)||'Não informado');vehicles[v]=(vehicles[v]||0)+1});
 make('typeChart','doughnut',Object.entries(vehicles).sort((a,b)=>b[1]-a[1]).slice(0,8));
 make('materialChart','bar',times,{horizontal:true,label:'Minutos'});
 make('dailyStatus','bar',top('STATUS'));
 make('variationChart','bar',times,{horizontal:true,label:'Minutos'});
 make('outlierTrend','bar',days,{dates:true,singleColor:true});
}
function render(){charts.forEach(c=>c.destroy());charts=[];let labels={overview:['Visão executiva','Indicadores e dashboards integrados em tempo real'],daily:['Acompanhamento operacional','Controle e consulta das operações'],variation:['Tempos e variações','Análise de produtividade e dispersão'],outliers:['Fora da média','Identificação automática de desvios'],tomorrow:['Programação','Planejamento das próximas operações'],share:['Programação para WhatsApp','Imagem pronta para compartilhar com a equipe'],};$('heading').textContent=labels[page][0];$('subtitle').textContent=labels[page][1];document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('active',b.dataset.page===page));let rs=filtered();$('content').innerHTML=({overview,daily,variation,outliers,tomorrow,share:shareSchedule}[page])(rs);drawCharts(rs)}
async function initAuth(){if(!window.msal)throw Error('Biblioteca Microsoft MSAL indisponível');if(CFG.clientId.startsWith('COLOQUE_')||CFG.tenant.startsWith('COLOQUE_'))throw Error('Configure clientId e tenant no script.js antes de conectar o OneDrive.');msalApp=new msal.PublicClientApplication({auth:{clientId:CFG.clientId,authority:`https://login.microsoftonline.com/${CFG.tenant}`,redirectUri:location.origin+location.pathname},cache:{cacheLocation:'sessionStorage'}});await msalApp.initialize();await msalApp.handleRedirectPromise();account=msalApp.getAllAccounts()[0];if(account)$('account').textContent=account.username}
async function token(){if(!msalApp)await initAuth();if(!account){let r=await msalApp.loginPopup({scopes:['Files.Read.All']});account=r.account;$('account').textContent=account.username}try{return (await msalApp.acquireTokenSilent({account,scopes:['Files.Read.All']})).accessToken}catch{return (await msalApp.acquireTokenPopup({account,scopes:['Files.Read.All']})).accessToken}}
async function graphGet(path,t){
 const r=await fetch('https://graph.microsoft.com/v1.0'+path,{headers:{Authorization:`Bearer ${t}`}});
 if(!r.ok){let message=(await r.text()).slice(0,450);throw Error(`Microsoft Graph ${r.status}: ${message}`)}
 return r.json();
}
async function findSharePointFile(t){
 const site=await graphGet(`/sites/${CFG.sharePointHost}:${CFG.sharePointSite}`,t);
 const drives=await graphGet(`/sites/${site.id}/drives`,t);
 const matches=[];
 for(const drive of drives.value||[]){
   // Search each document library. Exact filename match avoids selecting similarly named files.
   const results=await graphGet(`/drives/${drive.id}/root/search(q='${encodeURIComponent(CFG.sharePointFile.replace(/'/g,"''"))}')`,t);
   for(const item of results.value||[]){
     if(item.name?.normalize('NFC').toLocaleLowerCase('pt-BR')===CFG.sharePointFile.normalize('NFC').toLocaleLowerCase('pt-BR') && item.file){
       matches.push({driveId:drive.id,itemId:item.id,path:item.webUrl});
     }
   }
 }
 if(!matches.length)throw Error('Arquivo não encontrado nas bibliotecas do site OPERACAO. Confira acesso e nome do arquivo.');
 if(matches.length>1)throw Error('Foram encontradas várias planilhas com o mesmo nome no site OPERACAO. Informe a biblioteca/pasta exata.');
 return matches[0];
}
async function load(){
 try{
   $('sync').textContent='Sincronizando SharePoint...';
   const t=await token(),file=await findSharePointFile(t);
   const path=`/drives/${file.driveId}/items/${file.itemId}/workbook/worksheets/${encodeURIComponent(CFG.sheet)}/usedRange(valuesOnly=true)`;
   const result=await graphGet(path,t);
   process(result.values,'SharePoint Arrow');
 }catch(e){
   $('sync').textContent='Falha na sincronização';
   $('notice').textContent=e.message;
   console.error(e);
 }
}
function detail(i){let r=rows[i];if(!r)return;$('detailbody').innerHTML='<div class="detailgrid">'+Object.entries(r).filter(([k])=>!['MIN','DATA_ISO'].includes(k)).map(([k,v])=>`<div><small>${esc(k)}</small>${esc(v)}</div>`).join('')+'</div>';$('detail').showModal()}
function demo(){let clients=['TSEA Energia','COPASA','CXT','Expolumen','ORIGO','ZTT'],materials=['Transformadores','Cabos','Acessórios','Bobinas','Equipamentos'],statuses=['RECEBIDO','CARREGADO','EM ANDAMENTO','AGUARDANDO','PENDENTE'],types=['Paletizado','Solto','Especial'];let base=new Date();base.setDate(base.getDate()-29);let v=[['DATA','CLIENTE','OPERAÇÃO','MATERIAL','TIPOLOGIA','STATUS','AGENDADO?','TEMPO','DESTINO','NF','OBS']];for(let i=0;i<420;i++){let d=new Date(base);d.setDate(d.getDate()+i%32);v.push([todayFrom(d),clients[(i*7+Math.floor(i/9))%6],i%3===0?'DESCARGA':'CARGA',materials[(i*3+Math.floor(i/7))%5],types[i%3],statuses[(i*7+Math.floor(i/6))%5],i%7?'SIM':'NÃO',28+(i*17%155)+(i%29===0?170:0),'Betim/MG',`NF-${12000+i}`,'Dados fictícios para demonstração'])}process(v,'Demonstração');$('from').value=today(-29);$('to').value=today(1);render()}
function todayFrom(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function exportCSV(){let data=filtered(),cols=[...new Set(data.flatMap(r=>Object.keys(r)))].filter(k=>!['DATA_ISO','MIN'].includes(k));let q=v=>'"'+String(v??'').replace(/"/g,'""')+'"';let csv='\ufeff'+[cols.join(';'),...data.map(r=>cols.map(k=>q(r[k])).join(';'))].join('\r\n');let a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8'}));a.download='arrow-operacoes-filtradas.csv';a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000)}
window.addEventListener('DOMContentLoaded',()=>{document.querySelectorAll('nav button').forEach(b=>b.onclick=()=>{page=b.dataset.page;render()});['from','to','client','operation','material','type'].forEach(id=>$(id).addEventListener('change',render));$('clear').onclick=()=>{setDefaultDates();['client','operation','material','type'].forEach(id=>$(id).value='');render()};$('login').onclick=load;$('refresh').onclick=()=>mode==='SharePoint Arrow'?load():render();$('importBtn').onclick=()=>$('fileInput').click();$('fileInput').onchange=async e=>{let file=e.target.files[0];if(!file)return;try{if(!window.XLSX)throw Error('Biblioteca de leitura de Excel não carregou.');let book=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});let name=book.SheetNames.includes(CFG.sheet)?CFG.sheet:book.SheetNames[0];process(XLSX.utils.sheet_to_json(book.Sheets[name],{header:1,defval:'',raw:true}),`Arquivo: ${file.name}`)}catch(err){$('notice').textContent=`Erro ao importar: ${err.message}`}e.target.value=''};$('export').onclick=exportCSV;$('close').onclick=()=>$('detail').close();$('content').addEventListener('click',e=>{let tr=e.target.closest('tr[data-index]');if(tr)detail(+tr.dataset.index);if(e.target.id==='print'||e.target.id==='sharePrint')window.print();if(e.target.id==='shareImage')downloadShareImage()});$('content').addEventListener('change',e=>{if(e.target.id==='shareDate'||e.target.id==='shareCustomDate'){scheduleSelectedDate=e.target.value;render();return}if(e.target.id==='sigma'){threshold=+e.target.value;render()}if(e.target.id==='scheduleDate'){scheduleSelectedDate=e.target.value;render()}});demo();setInterval(()=>{if(account&&mode==='SharePoint Arrow')load()},CFG.refreshMs);if(!CFG.clientId.startsWith('COLOQUE_'))initAuth().then(()=>{if(account)load()}).catch(e=>$('notice').textContent=e.message)});
function tomorrowWithDate(date,rs){return `<div class="toolbar"><h2>Programação para ${displayDate(date)}</h2><label>Selecionar data <input type="date" id="scheduleDate" value="${date}"></label><button class="primary" id="print">Imprimir programação</button></div><h2 class="printtitle">ARROW — PROGRAMAÇÃO ${displayDate(date)}</h2><div class="cards">${card('Operações previstas',rs.length)}${card('Cargas',rs.filter(isLoad).length)}${card('Descargas',rs.filter(isUnload).length)}${card('Clientes',new Set(rs.map(r=>r.CLIENTE)).size)}</div>${table(rs,[['Cliente',r=>r.CLIENTE],['Operação',r=>r.OPERACAO],['Material',r=>r.MATERIAL],['Destino',r=>r.DESTINO],['Tipologia do veículo',r=>vehicleType(r)],['Status',r=>r.STATUS],['NF',r=>r.NF],['Observações',r=>r.OBS]])}`}
