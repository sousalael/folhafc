/* app.js — Upload, fuzzy mapping, tab nav, rendering — v3 com upload por slot e logo corrigida */
(function(){
"use strict";
Chart.register(ChartDataLabels);

/* ===== LOGO ===== */
var LOGO='data:image/png;base64,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';
document.getElementById('logo').src=LOGO;

/* ===== STATE ===== */
var State = {files:{estoque:null,contagem:null,vendas:null,cadastro:null,exclusoes:null},mappings:{},rawData:{estoque:[],contagem:[],vendas:[],cadastro:[],exclusoes:[]},results:{},processDate:'',charts:{},info:{cliente:'',unidade:'',dataInventario:'',diasVenda:90}};

/* ===== FUZZY MAPPING — com contexto de tipo de arquivo ===== */
var SYNONYMS = {
  sku:['sku','codigo','código','cod','cod_prod','cod.prod','codigo produto','código produto','codigo de barras','código de barras','ean','gtin','cod. item','codprod','item','cod_item','codigo_produto','product_code','barcode','cod interno','cod.item','cod item'],
  descricao:['descricao','descrição','desc','produto','nome','nome produto','nome_produto','description','desc_prod','desc. produto','nome do produto','desc produto'],
  categoria:['categoria','cat','departamento','dept','depto','seção','secao','setor','grupo','family','familia','família','category','classe','tipo','segmento','sub_grupo','subgrupo'],
  qtdSistema:['qtd sistema','qtd_sistema','quantidade sistema','estoque sistema','saldo sistema','saldo_sistema','estoque_sistema','qtd anterior','quantidade anterior','estoque anterior','system_qty','stock_qty','saldo','estoque','qtd_erp','qtd erp','qtd.sistema'],
  qtdContada:['qtd contada','qtd_contada','quantidade contada','contagem','qtd fisica','qtd_fisica','quantidade fisica','quantidade física','contado','counted_qty','physical_qty','qtd.contada','qtd inventário','qtd. contada'],
  custoUnit:['custo unit','custo_unit','custo unitario','custo unitário','custo medio','custo médio','preco custo','preço custo','valor unitario','valor unitário','unit_cost','cost','pmc','custo_unitario','vlr_custo','vlr custo','custo un','preco','preço','vlr unit'],
  local:['local','localizacao','localização','deposito','depósito','area','área','setor_loja','tipo_local','location','store_area','tipo local','loc','origem','local_contagem','local contagem'],
  qtdVendida:['qtd vendida','qtd_vendida','quantidade vendida','qtd venda','qtd_venda','sold_qty','sales_qty','un vendidas','unidades vendidas','qtd.vendida','volume_vendas'],
  valorVendido:['valor vendido','valor_vendido','faturamento','receita','venda valor','venda_valor','total vendido','total_vendido','revenue','sales_value','vlr vendido','vlr_vendido','fat','total vendas','valor vendas','receita_bruta','venda r$','vendas r$','vlr venda','vlr vendas'],
  custoVendido:['custo vendido','custo_vendido','custo mercadoria','custo_mercadoria','cost_sold','cogs','custo venda','custo_venda','custo das vendas','cmv_total','cmv r$','cmv','custo merc'],
  lucro:['lucro','lucro bruto','margem','margem bruta','profit','gross_profit','lucro_bruto','resultado','margem_bruta','lucro total','contribuicao','contribuição','lucratividade','rentabilidade','lucro r$']
};
/* Mapeamento prioritário por tipo: resolve ambiguidade de "Qde" "Qtde" etc */
var TYPE_PRIORITY = {
  estoque:{qtdSistema:['qde','qtde','qt','quantidade','qtd','saldo','estoque']},
  contagem:{qtdContada:['qde','qtde','qt','quantidade','qtd']},
  vendas:{qtdVendida:['qde','qtde','qt','quantidade','qtd','vendas','venda'],valorVendido:['venda r$','vendas r$'],custoVendido:['cmv r$','cmv']},
  cadastro:{}
};
function fuzzyMatch(header,excludeFields){
  var h=String(header).toLowerCase().trim().replace(/[_\-\.]/g,' ').replace(/\s+/g,' ');
  var best=null,bestScore=0;
  Object.keys(SYNONYMS).forEach(function(field){
    if(excludeFields&&excludeFields[field])return;
    SYNONYMS[field].forEach(function(syn){
      var s=syn.toLowerCase().trim(),score=0;
      if(h===s)score=100;
      else if(h.indexOf(s)>=0||s.indexOf(h)>=0)score=80;
      else{var w=s.split(' '),m=w.filter(function(x){return h.indexOf(x)>=0}).length;if(m>0)score=50*m/w.length;}
      if(score>bestScore){bestScore=score;best=field;}
    });
  });
  return bestScore>=40?best:null;
}
function autoMapHeadersForType(headers,fileType){
  var m={},usedH={},priorities=TYPE_PRIORITY[fileType]||{};
  // 1. Prioridades do tipo
  headers.forEach(function(h){
    var hL=String(h).toLowerCase().trim().replace(/[_\-\.]/g,' ').replace(/\s+/g,' ');
    Object.keys(priorities).forEach(function(field){
      if(m[field])return;
      priorities[field].forEach(function(syn){if(!m[field]&&(hL===syn||hL.indexOf(syn)>=0||syn.indexOf(hL)>=0)){m[field]=h;usedH[h]=true;}});
    });
  });
  // 2. Fuzzy genérico pro resto
  var mapped={};Object.keys(m).forEach(function(f){mapped[f]=true;});
  headers.forEach(function(h){
    if(usedH[h])return;
    var f=fuzzyMatch(h,mapped);
    if(f&&!m[f]){m[f]=h;mapped[f]=true;usedH[h]=true;}
  });
  return m;
}
function autoMapHeaders(headers){return autoMapHeadersForType(headers,'');}
function parseCSVLine(line,sep){var r=[],cur='',inQ=false;for(var i=0;i<line.length;i++){var ch=line[i];if(inQ){if(ch==='"'&&line[i+1]==='"'){cur+='"';i++;}else if(ch==='"'){inQ=false;}else{cur+=ch;}}else{if(ch==='"'){inQ=true;}else if(ch===sep){r.push(cur.trim());cur='';}else{cur+=ch;}}}r.push(cur.trim());return r;}
function readFile(file,cb){
  var r=new FileReader();
  var isText=/\.(csv|txt)$/i.test(file.name);
  r.onload=function(e){
    var d=new Uint8Array(e.target.result);
    if(isText){
      var text=new TextDecoder('utf-8').decode(d);
      if(text.indexOf('\ufffd')>=0)text=new TextDecoder('latin1').decode(d);
      var lines=text.split(/\r?\n/).filter(function(l){return l.trim();});
      if(!lines.length){cb({headers:[],rows:[],filename:file.name,rowCount:0});return;}
      var sep=lines[0].indexOf('\t')>=0?'\t':(lines[0].indexOf(';')>=0?';':',');
      var headers=parseCSVLine(lines[0],sep);
      var rows=[];
      for(var i=1;i<lines.length;i++){var vals=parseCSVLine(lines[i],sep);var row={};for(var j=0;j<headers.length;j++){row[headers[j]]=vals[j]||'';}rows.push(row);}
      cb({headers:headers,rows:rows,filename:file.name,rowCount:rows.length});
    }else{
      var wb=XLSX.read(d,{type:'array'});var sh=wb.Sheets[wb.SheetNames[0]];
      var json=XLSX.utils.sheet_to_json(sh,{defval:''});
      cb({headers:json.length?Object.keys(json[0]):[],rows:json,filename:file.name,rowCount:json.length});
    }
  };r.readAsArrayBuffer(file);
}
/* parseNumBR: número JS passa direto; string BR converte (ponto=milhar, vírgula=decimal) */
var NUMERIC_FIELDS={qtdSistema:1,qtdContada:1,custoUnit:1,qtdVendida:1,valorVendido:1,custoVendido:1,lucro:1};
function parseNumBR(val){
  if(val===null||val===undefined||val==='')return 0;
  if(typeof val==='number')return val;
  var s=String(val).trim().replace(/\s/g,'').replace(/^R\$\s*/i,'');
  if(!s)return 0;
  s=s.replace(/\./g,'').replace(',','.');
  var n=Number(s);return isNaN(n)?0:n;
}
function normalizeSKU(val){var s=String(val||'').trim();if(/^[\d.,\s]+$/.test(s))s=s.replace(/[.,\s]/g,'');return s;}
function applyMapping(rows,mapping){return rows.map(function(row){var o={};Object.keys(mapping).forEach(function(f){var v=row[mapping[f]];if(f==='sku'){o[f]=normalizeSKU(v);}else if(NUMERIC_FIELDS[f]){o[f]=parseNumBR(v);}else{o[f]=v;}});return o;});}

/* ===== UI HELPERS ===== */
var BRL=function(v){return(v<0?'−':'')+'R$ '+Math.abs(v||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});};
var BRLi=function(v){return(v<0?'−':'')+'R$ '+Math.abs(Math.round(v||0)).toLocaleString('pt-BR');};
var PCT=function(v){return(v||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'%';};
var NUM=function(v){return(v||0).toLocaleString('pt-BR');};
var NUMBR=function(v){if(v===null||v===undefined)return'—';return Number(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});};
function $(id){return document.getElementById(id);}

/* ===== TABS ===== */
document.querySelectorAll('.tab').forEach(function(tab){
  tab.addEventListener('click',function(){
    if(this.classList.contains('disabled'))return;
    document.querySelectorAll('.tab').forEach(function(t){t.classList.remove('active')});
    document.querySelectorAll('.panel').forEach(function(p){p.classList.remove('active')});
    this.classList.add('active');
    $('panel-'+this.dataset.tab).classList.add('active');
  });
});

/* ===== UPLOAD — cada slot é clicável ===== */
var SLOT_TYPES = ['estoque','contagem','vendas','cadastro','exclusoes'];
var slotFileInputs = {};

// Drop zone geral — abre seletor de tipo depois de ler o arquivo
var dropZone=$('dropZone'),fileInput=$('fileInput');
dropZone.addEventListener('click',function(){fileInput.click();});
dropZone.addEventListener('dragover',function(e){e.preventDefault();this.classList.add('dragover');});
dropZone.addEventListener('dragleave',function(){this.classList.remove('dragover');});
dropZone.addEventListener('drop',function(e){e.preventDefault();this.classList.remove('dragover');handleFilesGeneric(e.dataTransfer.files);});
fileInput.addEventListener('change',function(){handleFilesGeneric(this.files);this.value='';});

function handleFilesGeneric(files){
  for(var i=0;i<files.length;i++){
    (function(f){
      readFile(f,function(result){
        showFileTypeSelector(result);
      });
    })(files[i]);
  }
}

// Cada slot individual — clicável pra upload direto naquele tipo
SLOT_TYPES.forEach(function(type){
  var slot = $('slot-'+type);
  var inp = document.createElement('input');
  inp.type='file';inp.accept='.xlsx,.xls,.csv,.txt';inp.style.display='none';
  inp.dataset.slotType=type;
  slot.appendChild(inp);
  slotFileInputs[type]=inp;
  slot.style.cursor='pointer';
  slot.addEventListener('click',function(e){
    e.stopPropagation();
    inp.click();
  });
  inp.addEventListener('change',function(){
    var t=this.dataset.slotType;
    if(this.files.length){
      readFile(this.files[0],function(result){
        assignFileToSlot(t,result);
      });
    }
    this.value='';
  });
});

function showFileTypeSelector(result){
  $('mappingTitle').textContent='Qual é o tipo deste arquivo?';
  var body=$('mappingBody');
  body.innerHTML='<div style="margin-bottom:12px;font-size:13px">Arquivo: <strong>'+result.filename+'</strong> ('+result.rowCount.toLocaleString('pt-BR')+' linhas)</div>';
  body.innerHTML+='<div style="margin-bottom:12px;font-size:12px;color:#888">Colunas encontradas: '+result.headers.join(', ')+'</div>';
  var types=[
    {value:'estoque',label:'Estoque Sistema',desc:'Saldo do ERP antes da contagem'},
    {value:'contagem',label:'Contagem Física',desc:'Resultado da contagem no local'},
    {value:'vendas',label:'ABC Vendas',desc:'Histórico de vendas do período'},
    {value:'cadastro',label:'Cadastro de Produtos',desc:'Descrição e categoria dos SKUs'},
    {value:'exclusoes',label:'Lista de Exclusões',desc:'SKUs a excluir de todas as análises'}
  ];
  var btnsHtml='<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">';
  types.forEach(function(t){
    var loaded=State.files[t.value]?'<span style="font-size:11px;color:#F57C00"> (substituirá o atual)</span>':'';
    btnsHtml+='<div class="file-type-btn" data-type="'+t.value+'" style="border:1px solid #e0e0e0;border-radius:8px;padding:14px;cursor:pointer;transition:all .2s"><div style="font-weight:600;font-size:14px">'+t.label+loaded+'</div><div style="font-size:12px;color:#888;margin-top:2px">'+t.desc+'</div></div>';
  });
  btnsHtml+='</div>';
  body.innerHTML+=btnsHtml;
  // Store result for use in click handler
  window._pendingFileResult=result;
  body.querySelectorAll('.file-type-btn').forEach(function(btn){
    btn.addEventListener('click',function(){
      var type=this.dataset.type;
      assignFileToSlot(type,window._pendingFileResult);
      $('mappingModal').classList.remove('active');
      window._pendingFileResult=null;
    });
    btn.addEventListener('mouseenter',function(){this.style.borderColor='#00B74A';this.style.background='#fafff6';});
    btn.addEventListener('mouseleave',function(){this.style.borderColor='#e0e0e0';this.style.background='';});
  });
  // Show cancel button, hide confirm
  $('btnMappingConfirm').style.display='none';
  $('btnMappingCancel').style.display='';
  $('mappingModal').classList.add('active');
}

function assignFileToSlot(type, result){
  // Sempre mostra o mapeamento pro usuário confirmar antes de carregar
  State.files[type]=result;
  var mapping=autoMapHeadersForType(result.headers, type);
  showMappingConfirmation(type, result, mapping);
}

function showMappingConfirmation(type, result, currentMapping){
  var typeLabels={estoque:'Estoque Sistema',contagem:'Contagem Física',vendas:'Vendas 90 dias',cadastro:'Cadastro'};
  $('mappingTitle').textContent='Confirme o mapeamento — '+typeLabels[type];
  var body=$('mappingBody');
  body.innerHTML='<div style="margin-bottom:12px;font-size:13px">Arquivo: <strong>'+result.filename+'</strong> ('+result.rowCount.toLocaleString('pt-BR')+' linhas)</div>';
  body.innerHTML+='<div style="margin-bottom:16px;font-size:12px;color:var(--fc-muted)">Verifique se as colunas foram reconhecidas corretamente. Ajuste os campos que precisar.</div>';
  var allF=Object.keys(SYNONYMS);
  var fieldLabels={sku:'SKU / Código',descricao:'Descrição',categoria:'Categoria',qtdSistema:'Qtd. Sistema',qtdContada:'Qtd. Contada',custoUnit:'Custo Unitário',local:'Local (Loja/Depósito)',qtdVendida:'Qtd. Vendida',valorVendido:'Valor Vendido (R$)',custoVendido:'Custo Vendido / CMV',lucro:'Lucro / Margem'};
  // Inverter mapping: header -> field
  var headerToField={};
  Object.keys(currentMapping).forEach(function(f){headerToField[currentMapping[f]]=f;});
  result.headers.forEach(function(h){
    var row=document.createElement('div');row.className='mapping-row';
    var mapped=headerToField[h]||'';
    var matchColor=mapped?'color:var(--fc-green);font-weight:600':'color:var(--fc-muted)';
    row.innerHTML='<div class="mapping-field" style="'+matchColor+'">'+h+'</div><div class="mapping-arrow">→</div>';
    var s=document.createElement('select');s.className='mapping-select';s.dataset.original=h;
    s.innerHTML='<option value="">(ignorar)</option>';
    allF.forEach(function(f){
      var label=fieldLabels[f]||f;
      s.innerHTML+='<option value="'+f+'"'+(mapped===f?' selected':'')+'>'+label+'</option>';
    });
    row.appendChild(s);body.appendChild(row);
  });
  $('btnMappingConfirm').style.display='';
  $('btnMappingConfirm').onclick=function(){
    var m={};
    body.querySelectorAll('select.mapping-select').forEach(function(s){
      if(s.value)m[s.value]=s.dataset.original;
    });
    State.mappings[type]=m;
    State.rawData[type]=applyMapping(result.rows,m);
    updateSlot(type,result.filename,result.rowCount);
    $('mappingModal').classList.remove('active');
    checkReady();
  };
  $('mappingModal').classList.add('active');
}

function updateSlot(type,fn,cnt){
  var s=$('slot-'+type);
  s.querySelector('.file-slot-name').textContent=fn;
  var st=s.querySelector('.file-slot-status');
  st.textContent='✓ '+cnt.toLocaleString('pt-BR')+' linhas carregadas';
  st.className='file-slot-status loaded';
}

function checkReady(){
  var hasAnyFile=State.files.estoque||State.files.contagem||State.files.vendas;
  var infoOk=$('infoCliente').value.trim()&&$('infoUnidade').value.trim()&&$('infoData').value;
  $('btnProcess').disabled=!(hasAnyFile&&infoOk);
}
['infoCliente','infoUnidade','infoData'].forEach(function(id){$(id).addEventListener('input',checkReady);});

$('btnMappingCancel').addEventListener('click',function(){$('mappingModal').classList.remove('active');});

/* ===== PROCESS ===== */
$('btnProcess').addEventListener('click',processAll);
function processAll(){
  State.info.cliente=$('infoCliente').value.trim();
  State.info.unidade=$('infoUnidade').value.trim();
  var dVal=$('infoData').value;
  State.info.dataInventario=dVal?dVal.split('-').reverse().join('/'):'';
  State.info.diasVenda=parseInt($('infoDiasVenda').value)||90;  var bar=$('progressBar');bar.style.display='block';var fill=$('progressFill');fill.style.width='10%';
  State.processDate=new Date().toLocaleString('pt-BR');
  $('headerDate').innerHTML=State.info.cliente+' — '+State.info.unidade+'<br>Inventário: '+State.info.dataInventario+' | Processado em '+State.processDate;

  var hasEstoque=State.rawData.estoque.length>0;
  var hasContagem=State.rawData.contagem.length>0;
  var hasVendas=State.rawData.vendas.length>0;
  /* Expurgar SKUs da lista de exclusão */
  var excludeSet={};
  if(State.rawData.exclusoes&&State.rawData.exclusoes.length){
    State.rawData.exclusoes.forEach(function(r){var s=String(r.sku||'').trim();if(s)excludeSet[s]=true;});
    ['estoque','contagem','vendas','cadastro'].forEach(function(tipo){
      State.rawData[tipo]=State.rawData[tipo].filter(function(r){return!excludeSet[String(r.sku||'').trim()];});
    });
    hasEstoque=State.rawData.estoque.length>0;
    hasContagem=State.rawData.contagem.length>0;
    hasVendas=State.rawData.vendas.length>0;
  }
  State.results={};
  var avail={critica:false,ruptura:false,dias:false,abc:false,perda:false};
  // Mapa global de custo derivado do CMV das vendas
  var custoMap = hasVendas ? Engine.buildCustoMap(State.rawData.vendas) : {};

  setTimeout(function(){
    fill.style.width='20%';
    if(hasEstoque&&hasContagem){
      State.results.critica=Engine.calcCritica(State.rawData.estoque,State.rawData.contagem,State.rawData.cadastro,custoMap);
      avail.critica=true;
    }
    fill.style.width='35%';
    var itemsBase=null;
    if(State.results.critica){itemsBase=State.results.critica;}
    else if(hasContagem){
      var built=Engine.buildItemsFromContagem(State.rawData.contagem,State.rawData.cadastro);
      // Aplicar custo do mapa global
      built.forEach(function(it){ if(!it.custoUnit) it.custoUnit = custoMap[it.sku]||0; });
      itemsBase={items:built,totalSKUs:built.length};
    }
    setTimeout(function(){
      fill.style.width='50%';
      if(hasContagem){
        State.results.ruptura=Engine.calcRuptura(State.rawData.contagem,hasVendas?State.rawData.vendas:[],State.rawData.cadastro,State.info.diasVenda);
        avail.ruptura=true;
      }
      fill.style.width='65%';
      setTimeout(function(){
        if(itemsBase&&hasVendas){
          State.results.dias=Engine.calcDiasEstoque(itemsBase,State.rawData.vendas,custoMap,State.info.diasVenda,State.rawData.cadastro);
          avail.dias=true;
        }
        fill.style.width='75%';
        setTimeout(function(){
          if(itemsBase&&hasVendas){
            State.results.abc=Engine.calcInvestimentoABC(itemsBase,State.rawData.vendas,custoMap,State.info.diasVenda);
            avail.abc=true;
          }
          if(hasVendas&&(hasContagem||hasEstoque)){
            State.results.perda=Engine.calcProjecaoPerda(State.rawData.vendas,State.rawData.contagem.length?State.rawData.contagem:State.rawData.estoque,State.rawData.cadastro,State.info.diasVenda);
            avail.perda=true;
          }
          fill.style.width='100%';
          setTimeout(function(){
            bar.style.display='none';
            enableTabs(avail);
            renderAll(avail);
          },300);
        },50);
      },50);
    },50);
  },50);
}

function enableTabs(avail){
  ['critica','ruptura','dias','abc','perda'].forEach(function(name){
    var tab=document.querySelector('.tab[data-tab="'+name+'"]');
    if(avail[name])tab.classList.remove('disabled');else tab.classList.add('disabled');
  });
  document.querySelector('.tab[data-tab="upload"]').classList.remove('disabled');
  var first=['critica','ruptura','dias','abc','perda'].find(function(n){return avail[n];});
  if(first)document.querySelector('.tab[data-tab="'+first+'"]').click();
}

/* ===== RENDER HELPERS ===== */
function destroyChart(id){if(State.charts[id]){State.charts[id].destroy();delete State.charts[id];}}
function renderTable(p,headers,rows,page,perPage,opts){
  opts=opts||{};page=page||1;perPage=perPage||100;
  var start=(page-1)*perPage,pageRows=rows.slice(start,start+perPage),totalPages=Math.ceil(rows.length/perPage);
  var html='<div class="table-wrap"><table class="data-table"><thead><tr>';
  headers.forEach(function(h){html+='<th>'+h.label+'</th>';});
  html+='</tr></thead><tbody>';
  pageRows.forEach(function(row){
    var cls=opts.rowClass?opts.rowClass(row):'';
    html+='<tr class="'+cls+'">';
    headers.forEach(function(h){var val=h.render?h.render(row):(row[h.field]||'');html+='<td class="'+(h.align||'')+'">'+val+'</td>';});
    html+='</tr>';});
  html+='</tbody></table></div>';
  if(totalPages>1){html+='<div class="pagination"><button onclick="changePage(\''+p.id+'\','+(page-1)+')"'+(page<=1?' disabled':'')+'>← Anterior</button><span>Pág. '+page+'/'+totalPages+' ('+rows.length.toLocaleString('pt-BR')+' itens)</span><button onclick="changePage(\''+p.id+'\','+(page+1)+')"'+(page>=totalPages?' disabled':'')+'>Próxima →</button></div>';}
  return html;
}
var renderFns={};window.changePage=function(pid,pg){if(renderFns[pid])renderFns[pid](pg);};

function renderCatCards(catList,fields){
  var html='<div class="cat-grid">';
  catList.forEach(function(cat){
    html+='<div class="cat-card"><div class="cat-name">'+cat.nome+'</div>';
    fields.forEach(function(f){
      var val=f.fmt?f.fmt(cat[f.key]):cat[f.key];
      var cls=f.color?f.color(cat[f.key]):'';
      html+='<div class="cat-row"><span class="cat-label">'+f.label+'</span><span class="cat-val '+cls+'">'+val+'</span></div>';
    });
    html+='</div>';});
  html+='</div>';return html;
}
function renderCatFilterPills(catList,currentFilter,fnName){
  var html='<span style="color:var(--fc-muted);margin:0 2px">|</span>';
  html+='<span class="pill cat-pill '+(currentFilter==='all'?'active':'')+'" onclick="App.'+fnName+'(\'all\')">Todas categ.</span>';
  catList.forEach(function(cat){
    html+='<span class="pill cat-pill '+(currentFilter===cat.nome?'active':'')+'" onclick="App.'+fnName+'(\''+cat.nome.replace(/'/g,"\'")+'\')">'+cat.nome+'</span>';
  });
  return html;
}

function renderAll(avail){
  if(!avail)avail={critica:!!State.results.critica,ruptura:!!State.results.ruptura,dias:!!State.results.dias,abc:!!State.results.abc,perda:!!State.results.perda};
  if(avail.critica)renderCritica();
  if(avail.ruptura)renderRuptura();
  if(avail.dias)renderDias();
  if(avail.abc)renderABC();
  if(avail.perda)renderPerda();
  var depMsg={critica:'Estoque Sistema + Contagem Física',ruptura:'Contagem Física (com coluna de local)',dias:'Contagem Física + Vendas 90 dias',abc:'Contagem Física + Vendas 90 dias',perda:'Contagem Física + Vendas 90 dias'};
  ['critica','ruptura','dias','abc','perda'].forEach(function(name){
    if(!avail[name]){
      $('panel-'+name).innerHTML='<div class="alert alert-info"><i class="ti ti-info-circle"></i><div>Este relatório não pôde ser gerado com os arquivos carregados.<br><strong>Arquivos necessários:</strong> '+depMsg[name]+'</div></div>';
    }
  });
}


/* ===== 1. CRITICA ===== */
function renderCritica(page){
  var c=State.results.critica, p=$('panel-critica');
  var fS=p.dataset.filterStatus||'all', fC=p.dataset.filterCat||'all', srch=p.dataset.search||'';
  var filtered=c.items.filter(function(i){if(fS!=='all'&&i.status.toLowerCase()!==fS)return false;if(fC!=='all'&&(i.categoria||'Sem categoria')!==fC)return false;if(srch&&(i.sku+' '+i.descricao).toLowerCase().indexOf(srch.toLowerCase())<0)return false;return true;});
  var html='<div class="metrics">';
  html+='<div class="metric"><div class="metric-label">Acuracidade</div><div class="metric-value text-green">'+PCT(c.acuracidade)+'</div><div class="metric-detail">'+NUM(c.okCount)+' de '+NUM(c.totalSKUs)+' SKUs</div></div>';
  html+='<div class="metric"><div class="metric-label">Valor das faltas</div><div class="metric-value text-red">'+BRLi(c.totalFaltas)+'</div><div class="metric-detail">'+NUM(c.faltaCount)+' SKUs</div></div>';
  html+='<div class="metric"><div class="metric-label">Valor das sobras</div><div class="metric-value text-amber">'+BRLi(c.totalSobras)+'</div><div class="metric-detail">'+NUM(c.sobraCount)+' SKUs</div></div>';
  html+='<div class="metric"><div class="metric-label">Saldo líquido</div><div class="metric-value text-red">'+BRLi(c.saldoLiquido)+'</div><div class="metric-detail">Faltas − sobras</div></div>';
  html+='</div>';
  if(c.hasCategorias){
    html+='<div class="section-title"><i class="ti ti-category"></i> Resultado por categoria</div>';
    html+=renderCatCards(c.categorias,[{label:'Acuracidade',key:'acuracidade',fmt:PCT,color:function(){return 'text-green';}},{label:'Faltas',key:'faltaVal',fmt:BRLi,color:function(){return 'text-red';}},{label:'Sobras',key:'sobraVal',fmt:BRLi,color:function(v){return v>0?'text-amber':'text-muted';}},{label:'Saldo',key:'saldo',fmt:BRLi,color:function(v){return v<0?'text-red':'text-green';}}]);
    html+='<div class="chart-legend"><span class="legend-item"><span class="legend-dot" style="background:#D32F2F"></span>Faltas (R$)</span><span class="legend-item"><span class="legend-dot" style="background:#F57C00"></span>Sobras (R$)</span></div>';
    html+='<div class="chart-wrap" style="height:'+Math.max(160,c.categorias.length*50)+'px"><canvas id="chartCritica"></canvas></div>';
  }
  html+='<div class="toolbar"><input class="search-input" placeholder="Buscar SKU ou descrição..." value="'+srch+'" onkeyup="App.filterCritica(this.value)">';
  html+='<span class="pill '+(fS==='all'?'active':'')+'" onclick="App.filterCriticaStatus(\'all\')">Todos</span><span class="pill '+(fS==='falta'?'active':'')+'" onclick="App.filterCriticaStatus(\'falta\')">Faltas</span><span class="pill '+(fS==='sobra'?'active':'')+'" onclick="App.filterCriticaStatus(\'sobra\')">Sobras</span><span class="pill '+(fS==='ok'?'active':'')+'" onclick="App.filterCriticaStatus(\'ok\')">Sem diverg.</span>';
  if(c.hasCategorias) html+=renderCatFilterPills(c.categorias,fC,'filterCriticaCat');
  html+='<button class="btn-export" onclick="App.openExport()"><i class="ti ti-download"></i> Excel</button><button class="btn-export btn-pdf" onclick="App.exportPDF(\'critica\')"><i class="ti ti-file-text"></i> PDF</button></div>';
  var th=[{label:'SKU',field:'sku'},{label:'Descrição',field:'descricao'},{label:'Categoria',field:'categoria'},{label:'Qtd sist.',align:'text-right',render:function(r){return NUMBR(r.qtdSistema);}},{label:'Qtd cont.',align:'text-right',render:function(r){return NUMBR(r.qtdContada);}},{label:'Dif. qtd',align:'text-right',render:function(r){return '<span class="'+(r.difQtd<0?'text-red':(r.difQtd>0?'text-green':'text-muted'))+'">'+NUMBR(r.difQtd)+'</span>';}},{label:'Dif. R$',align:'text-right',render:function(r){return '<span class="'+(r.difValor<0?'text-red':(r.difValor>0?'text-green':'text-muted'))+'">'+BRL(r.difValor)+'</span>';}},{label:'Status',align:'text-center',render:function(r){var cls=r.status==='Falta'?'badge-falta':(r.status==='Sobra'?'badge-sobra':'badge-ok');return '<span class="badge '+cls+'">'+r.status+'</span>';}}];
  html+=renderTable(p,th,filtered,page||1,100);
  p.innerHTML=html; renderFns['panel-critica']=renderCritica;
  if(c.hasCategorias){destroyChart('chartCritica');var ctx=document.getElementById('chartCritica');if(ctx){State.charts.chartCritica=new Chart(ctx,{type:'bar',data:{labels:c.categorias.map(function(c){return c.nome;}),datasets:[{label:'Faltas',data:c.categorias.map(function(c){return Math.abs(c.faltaVal);}),backgroundColor:'#D32F2F',borderRadius:4,barPercentage:.65},{label:'Sobras',data:c.categorias.map(function(c){return c.sobraVal;}),backgroundColor:'#F57C00',borderRadius:4,barPercentage:.65}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},datalabels:{anchor:'end',align:'end',color:function(c){return c.dataset.backgroundColor;},font:{size:10,weight:'bold'},formatter:function(v){return 'R$ '+v.toLocaleString('pt-BR');}}},scales:{x:{grid:{color:'#f0f0f0'},ticks:{font:{size:10},callback:function(v){return 'R$ '+(v/1000).toFixed(0)+'k';}}},y:{grid:{display:false},ticks:{font:{size:10}}}},layout:{padding:{right:80}}}});}}
}

/* ===== 2. RUPTURA ===== */
function renderRuptura(page){
  var r=State.results.ruptura, p=$('panel-ruptura');
  var fA=p.dataset.filterAbc||'all', fC=p.dataset.filterCat||'all', srch=p.dataset.search||'';
  var filtered=r.items.filter(function(i){if(fA!=='all'&&i.abc_valorVendido90!==fA)return false;if(fC!=='all'&&(i.categoria||'Sem categoria')!==fC)return false;if(srch&&(i.sku+' '+i.descricao).toLowerCase().indexOf(srch.toLowerCase())<0)return false;return true;});
  var html='<div class="metrics">';
  html+='<div class="metric"><div class="metric-label">Taxa de ruptura geral</div><div class="metric-value text-red">'+PCT(r.taxaRuptura)+'</div></div>';
  html+='<div class="metric"><div class="metric-label">SKUs em ruptura</div><div class="metric-value">'+NUM(r.totalRupturas)+'</div><div class="metric-detail">de '+NUM(r.totalComDeposito)+' com depósito</div></div>';
  html+='<div class="metric"><div class="metric-label">Ruptura curva A (fat.)</div><div class="metric-value text-red">'+PCT(r.taxaA)+'</div></div>';
  html+='<div class="metric"><div class="metric-label">Ruptura curva A (lucro)</div><div class="metric-value text-red">'+PCT(r.taxaALucro)+'</div></div>';
  html+='</div>';
  if(r.rupturaA>0) html+='<div class="alert alert-danger"><i class="ti ti-alert-circle"></i><div><strong>'+NUM(r.rupturaA)+' itens curva A</strong> estão em ruptura — produtos de maior venda ausentes no salão</div></div>';
  if(r.hasCategorias){
    html+='<div class="section-title"><i class="ti ti-category"></i> Ruptura por categoria</div>';
    html+=renderCatCards(r.categorias,[{label:'Rupturas',key:'totalRupturas',fmt:NUM,color:function(v){return v>0?'text-red':'';}},{label:'Taxa ruptura',key:'taxa',fmt:PCT,color:function(){return 'text-red';}},{label:'Rupturas curva A',key:'rupturaA',fmt:NUM,color:function(v){return v>0?'text-red':'text-muted';}},{label:'Perda fat./dia',key:'perdaDia',fmt:BRLi,color:function(){return 'text-red';}}]);
  }
  html+='<div class="loss-cards"><div class="loss-card a"><div class="loss-title">Curva A em ruptura</div><div class="loss-main">'+NUM(r.rupturaA)+' SKUs</div><div class="loss-sub">Ação imediata</div></div><div class="loss-card b"><div class="loss-title">Curva B em ruptura</div><div class="loss-main">'+NUM(r.rupturaB)+' SKUs</div><div class="loss-sub">Atenção</div></div><div class="loss-card c"><div class="loss-title">Curva C em ruptura</div><div class="loss-main">'+NUM(r.rupturaC)+' SKUs</div><div class="loss-sub">Monitorar</div></div></div>';
  html+='<div class="toolbar"><input class="search-input" placeholder="Buscar..." value="'+srch+'" onkeyup="App.filterRupturaSearch(this.value)">';
  html+='<span class="pill '+(fA==='all'?'active':'')+'" onclick="App.filterRupturaAbc(\'all\')">Todos</span><span class="pill '+(fA==='A'?'active':'')+'" onclick="App.filterRupturaAbc(\'A\')">Curva A</span><span class="pill '+(fA==='B'?'active':'')+'" onclick="App.filterRupturaAbc(\'B\')">Curva B</span><span class="pill '+(fA==='C'?'active':'')+'" onclick="App.filterRupturaAbc(\'C\')">Curva C</span>';
  if(r.hasCategorias) html+=renderCatFilterPills(r.categorias,fC,'filterRupturaCat');
  html+='<button class="btn-export" onclick="App.openExport()"><i class="ti ti-download"></i> Excel</button><button class="btn-export btn-pdf" onclick="App.exportPDF(\'ruptura\')"><i class="ti ti-file-text"></i> PDF</button></div>';
  var th=[{label:'SKU',field:'sku'},{label:'Descrição',field:'descricao'},{label:'Categoria',field:'categoria'},{label:'ABC fat.',align:'text-center',render:function(r){var c=r.abc_valorVendido90||'C';return '<span class="badge badge-'+c.toLowerCase()+'">'+c+'</span>';}},{label:'Qtd depósito',align:'text-right',render:function(r){return NUMBR(r.deposito);}},{label:'Qtd loja',align:'text-right',render:function(r){return '<span class="text-red">'+NUMBR(r.loja)+'</span>';}},{label:'Venda méd/dia',align:'text-right',render:function(r){return r.vendaMediaDia?Engine.round2(r.vendaMediaDia):'—';}},{label:'Fat. méd/dia',align:'text-right',render:function(r){return r.fatMediaDia?BRL(r.fatMediaDia):'—';}}];
  html+=renderTable(p,th,filtered,page||1,100,{rowClass:function(r){return r.abc_valorVendido90==='A'?'row-a':(r.abc_valorVendido90==='B'?'row-b':'');}});
  p.innerHTML=html; renderFns['panel-ruptura']=renderRuptura;
}

/* ===== 3. DIAS ESTOQUE ===== */
function renderDias(page){
  var d=State.results.dias, p=$('panel-dias');
  var fF=p.dataset.filterFaixa||'all', fC=p.dataset.filterCat||'all', srch=p.dataset.search||'';
  var filtered=d.items.filter(function(i){if(fF!=='all'&&i.faixa!==fF)return false;if(fC!=='all'&&(i.categoria||'Sem categoria')!==fC)return false;if(srch&&(i.sku+' '+i.descricao).toLowerCase().indexOf(srch.toLowerCase())<0)return false;return true;});
  var html='<div class="metrics" style="grid-template-columns:repeat(5,1fr)">';
  html+='<div class="metric"><div class="metric-label">Cobertura de estoque</div><div class="metric-value">'+d.coberturaGeral+' dias</div></div>';
  html+='<div class="metric"><div class="metric-label">Cobertura curva A</div><div class="metric-value">'+d.coberturaA+' dias</div></div>';
  html+='<div class="metric"><div class="metric-label">Cobertura curva B</div><div class="metric-value">'+d.coberturaB+' dias</div></div>';
  html+='<div class="metric"><div class="metric-label">Cobertura curva C</div><div class="metric-value">'+d.coberturaC+' dias</div></div>';
  html+='<div class="metric"><div class="metric-label">SKUs sem giro</div><div class="metric-value text-red">'+NUM(d.semGiro)+'</div></div>';
  html+='</div>';
  // Distribution cards - 5 faixas
  html+='<div class="dist-grid" style="grid-template-columns:repeat(5,1fr)">';
  html+='<div class="dist-card crit"><div class="dist-label">Ruptura (0-2 dias)</div><div class="dist-val">'+NUM(d.ruptura)+'</div><div class="dist-sub">'+PCT(d.total?d.ruptura/d.total*100:0)+'</div></div>';
  html+='<div class="dist-card" style="background:var(--fc-al);border:1px solid var(--fc-amb)"><div class="dist-label" style="color:var(--fc-amb)">Alto risco (3-5 dias)</div><div class="dist-val" style="color:var(--fc-amb)">'+NUM(d.altoRisco)+'</div><div class="dist-sub">'+PCT(d.total?d.altoRisco/d.total*100:0)+'</div></div>';
  html+='<div class="dist-card" style="background:#FFF8E1;border:1px solid #FBC02D"><div class="dist-label" style="color:#F57F17">Médio risco (6-15 dias)</div><div class="dist-val" style="color:#F57F17">'+NUM(d.medioRisco)+'</div><div class="dist-sub">'+PCT(d.total?d.medioRisco/d.total*100:0)+'</div></div>';
  html+='<div class="dist-card ok"><div class="dist-label">Cobertura ideal (16-30 dias)</div><div class="dist-val">'+NUM(d.coberturaIdeal)+'</div><div class="dist-sub">'+PCT(d.total?d.coberturaIdeal/d.total*100:0)+'</div></div>';
  html+='<div class="dist-card high"><div class="dist-label">Excesso (31+ dias)</div><div class="dist-val">'+NUM(d.excessos)+'</div><div class="dist-sub">'+PCT(d.total?d.excessos/d.total*100:0)+'</div></div>';
  html+='</div>';
  if(d.hasCategorias){
    html+='<div class="section-title"><i class="ti ti-category"></i> Cobertura por categoria</div>';
    html+=renderCatCards(d.categorias,[
      {label:'Cobertura média',key:'mediaCobertura',fmt:function(v){return v+' dias';},color:function(){return '';}},
      {label:'Val. estoque',key:'valorEstoque',fmt:BRLi,color:function(){return '';}},
      {label:'Ruptura + Alto risco',key:'criticos',fmt:NUM,color:function(v){return v>0?'text-red':'text-muted';}},
      {label:'Sem giro',key:'semGiro',fmt:NUM,color:function(v){return v>0?'text-red':'text-muted';}},
      {label:'Excesso (31+d)',key:'excessos',fmt:NUM,color:function(v){return v>0?'text-blue':'text-muted';}}
    ]);
  }
  html+='<div class="toolbar"><input class="search-input" placeholder="Buscar..." value="'+srch+'" onkeyup="App.filterDiasSearch(this.value)">';
  ['all','Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'].forEach(function(f){html+='<span class="pill '+(fF===f?'active':'')+'" onclick="App.filterDiasFaixa(\''+f+'\')">'+(f==='all'?'Todos':f)+'</span>';});
  if(d.hasCategorias) html+=renderCatFilterPills(d.categorias,fC,'filterDiasCat');
  html+='<button class="btn-export" onclick="App.openExport()"><i class="ti ti-download"></i> Excel</button><button class="btn-export btn-pdf" onclick="App.exportPDF(\'dias\')"><i class="ti ti-file-text"></i> PDF</button></div>';
  var th=[{label:'SKU',field:'sku'},{label:'Descrição',field:'descricao'},{label:'Categoria',field:'categoria'},{label:'Qtd estoque',align:'text-right',render:function(r){return NUMBR(r.qtdEstoque);}},{label:'Venda méd/dia',align:'text-right',render:function(r){return r.vendaMediaDia!==null?NUMBR(r.vendaMediaDia):'—';}},{label:'Dias estoque',align:'text-right',render:function(r){return r.diasEstoque!==null?NUMBR(r.diasEstoque):'—';}},{label:'Cobertura',align:'text-center',render:function(r){var cls={'Ruptura':'badge-falta','Alto risco':'badge-sobra','Médio risco':'badge-b','Cobertura ideal':'badge-ok','Excesso de cobertura':'badge-b','Sem giro':'badge-c'};return '<span class="badge '+(cls[r.faixa]||'')+'">'+r.faixa+'</span>';}},{label:'Val. estoque',align:'text-right',render:function(r){return BRL(r.valorEstoque);}},{label:'ABC fat.',field:'abcFat',align:'text-center'}];
  html+=renderTable(p,th,filtered,page||1,100);
  p.innerHTML=html; renderFns['panel-dias']=renderDias;
}

/* ===== 4. INVESTIMENTO ABC ===== */
function renderABC(page){
  var a=State.results.abc, p=$('panel-abc');
  var fC=p.dataset.filterCat||'all', srch=p.dataset.search||'';
  var filtered=a.items.filter(function(i){if(fC!=='all'&&(i.categoria||'Sem categoria')!==fC)return false;if(srch&&(i.sku+' '+i.descricao).toLowerCase().indexOf(srch.toLowerCase())<0)return false;return true;});
  var html='<div class="metrics"><div class="metric"><div class="metric-label">Valor total em estoque</div><div class="metric-value">'+BRLi(a.totalInvest)+'</div></div><div class="metric"><div class="metric-label">Faturamento 90 dias</div><div class="metric-value">'+BRLi(a.totalFat)+'</div></div><div class="metric"><div class="metric-label">Lucro 90 dias</div><div class="metric-value">'+BRLi(a.totalLucro)+'</div></div><div class="metric"><div class="metric-label">SKUs analisados</div><div class="metric-value">'+NUM(a.items.length)+'</div></div></div>';
  html+='<div class="summary-pair"><div class="summary-card"><div class="summary-header fat">Curva ABC por faturamento</div>';
  [{c:'A',d:a.fatA},{c:'B',d:a.fatB},{c:'C',d:a.fatC}].forEach(function(x){html+='<div class="summary-row"><span class="summary-class text-'+(x.c==='A'?'red':(x.c==='B'?'amber':'muted'))+'">'+x.c+'</span><div class="bar-track"><div class="bar-fill" style="width:'+x.d.pctInvest+'%;background:var(--fc-navy);opacity:.7"></div></div><span class="summary-pct">'+PCT(x.d.pctInvest)+' invest.</span></div><div class="summary-row"><span class="summary-class" style="visibility:hidden">'+x.c+'</span><div class="bar-track"><div class="bar-fill" style="width:'+x.d.pctFat+'%;background:var(--fc-green);opacity:.7"></div></div><span class="summary-pct">'+PCT(x.d.pctFat)+' fat.</span></div>';});
  html+='</div><div class="summary-card"><div class="summary-header luc">Curva ABC por lucro</div>';
  [{c:'A',d:a.lucA},{c:'B',d:a.lucB},{c:'C',d:a.lucC}].forEach(function(x){html+='<div class="summary-row"><span class="summary-class text-'+(x.c==='A'?'red':(x.c==='B'?'amber':'muted'))+'">'+x.c+'</span><div class="bar-track"><div class="bar-fill" style="width:'+x.d.pctInvest+'%;background:var(--fc-navy);opacity:.7"></div></div><span class="summary-pct">'+PCT(x.d.pctInvest)+' invest.</span></div><div class="summary-row"><span class="summary-class" style="visibility:hidden">'+x.c+'</span><div class="bar-track"><div class="bar-fill" style="width:'+x.d.pctLuc+'%;background:var(--fc-green);opacity:.7"></div></div><span class="summary-pct">'+PCT(x.d.pctLuc)+' lucro</span></div>';});
  html+='</div></div>';
  if(a.fatC.pctInvest>a.fatC.pctFat+5) html+='<div class="alert alert-warning"><i class="ti ti-bulb"></i><div>Itens curva C consomem '+PCT(a.fatC.pctInvest)+' do capital investido mas geram apenas '+PCT(a.fatC.pctFat)+' da receita.</div></div>';
  if(a.hasCategorias){
    html+='<div class="section-title"><i class="ti ti-category"></i> Investimento por categoria</div>';
    html+=renderCatCards(a.categorias,[{label:'SKUs',key:'total',fmt:NUM,color:function(){return '';}},{label:'Investimento',key:'investimento',fmt:BRLi,color:function(){return '';}},{label:'Faturamento',key:'faturamento',fmt:BRLi,color:function(){return 'text-green';}},{label:'Lucro',key:'lucro',fmt:BRLi,color:function(v){return v>=0?'text-green':'text-red';}},{label:'% do invest.',key:'pctInvest',fmt:PCT,color:function(){return '';}}]);
  }
  html+='<div class="toolbar"><input class="search-input" placeholder="Buscar..." value="'+srch+'" onkeyup="App.filterAbcSearch(this.value)">';
  if(a.hasCategorias) html+=renderCatFilterPills(a.categorias,fC,'filterAbcCat');
  html+='<button class="btn-export" onclick="App.openExport()"><i class="ti ti-download"></i> Excel</button><button class="btn-export btn-pdf" onclick="App.exportPDF(\'abc\')"><i class="ti ti-file-text"></i> PDF</button></div>';
  var th=[{label:'SKU',field:'sku'},{label:'Descrição',field:'descricao'},{label:'ABC fat.',align:'text-center',render:function(r){return '<span class="badge badge-'+r.abcFat.toLowerCase()+'">'+r.abcFat+'</span>';}},{label:'ABC lucro',align:'text-center',render:function(r){return '<span class="badge badge-'+r.abcLucro.toLowerCase()+'">'+r.abcLucro+'</span>';}},{label:'Qtd estoque',align:'text-right',render:function(r){return NUMBR(r.qtdEstoque);}},{label:'Valor Estoque',align:'text-right',render:function(r){return BRL(r.valorInvestido);}},{label:'Fat. 90d',align:'text-right',render:function(r){return BRL(r.fat90);}},{label:'Lucro 90d',align:'text-right',render:function(r){return BRL(r.lucro90);}}];
  html+=renderTable(p,th,filtered,page||1,100);
  p.innerHTML=html; renderFns['panel-abc']=renderABC;
}

/* ===== 5. PROJECAO PERDA ===== */
function renderPerda(page){
  var pe=State.results.perda, p=$('panel-perda');
  var fA=p.dataset.filterAbc||'all', fC=p.dataset.filterCat||'all', srch=p.dataset.search||'';
  var filtered=pe.items.filter(function(i){if(fA!=='all'&&i.abcFat!==fA)return false;if(fC!=='all'&&(i.categoria||'Sem categoria')!==fC)return false;if(srch&&(i.sku+' '+i.descricao).toLowerCase().indexOf(srch.toLowerCase())<0)return false;return true;});
  var html='<div class="metrics"><div class="metric"><div class="metric-label">Venda perdida / dia</div><div class="metric-value text-red">'+BRLi(pe.totalPerdaFat)+'</div></div><div class="metric"><div class="metric-label">Lucro perdido / dia</div><div class="metric-value text-red">'+BRLi(pe.totalPerdaLucro)+'</div></div><div class="metric"><div class="metric-label">Perda mensal (fat.)</div><div class="metric-value text-red">'+BRLi(pe.perdaMensal)+'</div></div><div class="metric"><div class="metric-label">SKUs em ruptura</div><div class="metric-value">'+NUM(pe.totalSKUs)+'</div></div></div>';
  html+='<div class="loss-cards"><div class="loss-card a"><div class="loss-title">Curva A — perda/dia</div><div class="loss-main">'+BRLi(pe.classA.perda)+'</div><div class="loss-sub">'+PCT(pe.classA.pct)+' — '+NUM(pe.classA.count)+' SKUs</div></div><div class="loss-card b"><div class="loss-title">Curva B — perda/dia</div><div class="loss-main">'+BRLi(pe.classB.perda)+'</div><div class="loss-sub">'+PCT(pe.classB.pct)+' — '+NUM(pe.classB.count)+' SKUs</div></div><div class="loss-card c"><div class="loss-title">Curva C — perda/dia</div><div class="loss-main">'+BRLi(pe.classC.perda)+'</div><div class="loss-sub">'+PCT(pe.classC.pct)+' — '+NUM(pe.classC.count)+' SKUs</div></div></div>';
  html+='<div class="alert alert-danger"><i class="ti ti-alert-triangle"></i><div><strong>Impacto curva A:</strong> Diário: '+BRLi(pe.classA.perda)+' | Semanal: '+BRLi(pe.classA.perda*7)+' | Mensal: '+BRLi(pe.classA.perda*30)+' | Lucro mensal perdido: '+BRLi(pe.classA.lucro*30)+'</div></div>';
  if(pe.hasCategorias){
    html+='<div class="section-title"><i class="ti ti-category"></i> Perda projetada por categoria</div>';
    html+=renderCatCards(pe.categorias,[{label:'Rupturas',key:'totalRupturas',fmt:NUM,color:function(){return '';}},{label:'Perda fat./dia',key:'perdaFatDia',fmt:BRLi,color:function(){return 'text-red';}},{label:'Perda lucro/dia',key:'perdaLucroDia',fmt:BRLi,color:function(){return 'text-red';}},{label:'Perda mensal',key:'perdaMensal',fmt:BRLi,color:function(){return 'text-red';}},{label:'Rupturas A',key:'rupturaA',fmt:NUM,color:function(v){return v>0?'text-red':'text-muted';}}]);
  }
  html+='<div class="toolbar"><input class="search-input" placeholder="Buscar..." value="'+srch+'" onkeyup="App.filterPerdaSearch(this.value)">';
  html+='<span class="pill '+(fA==='all'?'active':'')+'" onclick="App.filterPerdaAbc(\'all\')">Todos</span><span class="pill '+(fA==='A'?'active':'')+'" onclick="App.filterPerdaAbc(\'A\')">Curva A</span><span class="pill '+(fA==='B'?'active':'')+'" onclick="App.filterPerdaAbc(\'B\')">Curva B</span><span class="pill '+(fA==='C'?'active':'')+'" onclick="App.filterPerdaAbc(\'C\')">Curva C</span>';
  if(pe.hasCategorias) html+=renderCatFilterPills(pe.categorias,fC,'filterPerdaCat');
  html+='<button class="btn-export" onclick="App.openExport()"><i class="ti ti-download"></i> Excel</button><button class="btn-export btn-pdf" onclick="App.exportPDF(\'perda\')"><i class="ti ti-file-text"></i> PDF</button></div>';
  var th=[{label:'SKU',field:'sku'},{label:'Descrição',field:'descricao'},{label:'Categoria',field:'categoria'},{label:'ABC fat.',align:'text-center',render:function(r){return '<span class="badge badge-'+r.abcFat.toLowerCase()+'">'+r.abcFat+'</span>';}},{label:'Venda méd/dia',align:'text-right',render:function(r){return NUMBR(r.vendaMediaDia);}},{label:'Perda fat./dia',align:'text-right',render:function(r){return '<span class="text-red">'+BRL(r.perdaFatDia)+'</span>';}},{label:'Perda lucro/dia',align:'text-right',render:function(r){return '<span class="text-red">'+BRL(r.perdaLucroDia)+'</span>';}},{label:'Perda fat./mês',align:'text-right',render:function(r){return '<span class="text-red">'+BRL(r.perdaFatMes)+'</span>';}},{label:'Perda lucro/mês',align:'text-right',render:function(r){return '<span class="text-red">'+BRL(r.perdaLucroMes)+'</span>';}}];
  html+=renderTable(p,th,filtered,page||1,100,{rowClass:function(r){return r.abcFat==='A'?'row-a':(r.abcFat==='B'?'row-b':'');}});
  html+='<div class="note"><i class="ti ti-info-circle"></i><span>Premissa: a venda média dos últimos 90 dias representa a demanda normal. Valores projetados são estimativas.</span></div>';
  p.innerHTML=html; renderFns['panel-perda']=renderPerda;
}

/* ===== FILTER HANDLERS ===== */
window.App = {
  filterCritica:function(v){$('panel-critica').dataset.search=v;renderCritica(1);},
  filterCriticaStatus:function(v){$('panel-critica').dataset.filterStatus=v;renderCritica(1);},
  filterCriticaCat:function(v){$('panel-critica').dataset.filterCat=v;renderCritica(1);},
  filterRupturaSearch:function(v){$('panel-ruptura').dataset.search=v;renderRuptura(1);},
  filterRupturaAbc:function(v){$('panel-ruptura').dataset.filterAbc=v;renderRuptura(1);},
  filterRupturaCat:function(v){$('panel-ruptura').dataset.filterCat=v;renderRuptura(1);},
  filterDiasSearch:function(v){$('panel-dias').dataset.search=v;renderDias(1);},
  filterDiasFaixa:function(v){$('panel-dias').dataset.filterFaixa=v;renderDias(1);},
  filterDiasCat:function(v){$('panel-dias').dataset.filterCat=v;renderDias(1);},
  filterAbcSearch:function(v){$('panel-abc').dataset.search=v;renderABC(1);},
  filterAbcCat:function(v){$('panel-abc').dataset.filterCat=v;renderABC(1);},
  filterPerdaSearch:function(v){$('panel-perda').dataset.search=v;renderPerda(1);},
  filterPerdaAbc:function(v){$('panel-perda').dataset.filterAbc=v;renderPerda(1);},
  filterPerdaCat:function(v){$('panel-perda').dataset.filterCat=v;renderPerda(1);},
  openExport:function(){
    var r=State.results;
    $('exp-critica-resumo').disabled=!r.critica;$('exp-critica-resumo').checked=!!r.critica;
    $('exp-critica-detalhe').disabled=!r.critica;$('exp-critica-detalhe').checked=!!r.critica;
    $('exp-ruptura').disabled=!r.ruptura;$('exp-ruptura').checked=!!r.ruptura;
    $('exp-dias').disabled=!r.dias;$('exp-dias').checked=!!r.dias;
    $('exp-abc').disabled=!r.abc;$('exp-abc').checked=!!r.abc;
    $('exp-perda').disabled=!r.perda;$('exp-perda').checked=!!r.perda;
    $('exportModal').classList.add('active');
  },
  exportPDF:function(type){Export.generatePDF(type,State.results,State.processDate,LOGO,State.info);}
};

$('btnExportConfirm').addEventListener('click',function(){var sel={criticaResumo:$('exp-critica-resumo').checked,criticaDetalhe:$('exp-critica-detalhe').checked,ruptura:$('exp-ruptura').checked,dias:$('exp-dias').checked,abc:$('exp-abc').checked,perda:$('exp-perda').checked};Export.generateExcel(State.results,sel,State.processDate,State.info);$('exportModal').classList.remove('active');});
$('btnExportCancel').addEventListener('click',function(){$('exportModal').classList.remove('active');});
})();
