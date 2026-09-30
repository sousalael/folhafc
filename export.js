/* export.js — v3.0 — Excel formatado + Dashboard + PDF com resumo executivo */
var Export=(function(){
"use strict";
var C={navy:'001528',green:'00B74A',red:'D32F2F',amb:'F57C00',blue:'1565C0',white:'FFFFFF',light:'F5F5F5',lightG:'F0F0F0',border:'D0D0D0',text:'333333',muted:'888888'};
var BRL=function(v){return(v<0?'-':'')+'R$ '+Math.abs(v||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});};
var BRLi=function(v){return(v<0?'-':'')+'R$ '+Math.abs(Math.round(v||0)).toLocaleString('pt-BR');};
var PCT=function(v){return(v||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'%';};
var NUM=function(v){return(v||0).toLocaleString('pt-BR');};
var R2=function(v){return Engine.round2(v||0);};

/* ===== ESTILOS EXCEL ===== */
function sH(){return{font:{bold:true,color:{rgb:C.white},sz:10,name:'Arial'},fill:{fgColor:{rgb:C.navy}},alignment:{horizontal:'left',vertical:'center',wrapText:true},border:{bottom:{style:'thin',color:{rgb:C.border}},top:{style:'thin',color:{rgb:C.border}},left:{style:'thin',color:{rgb:C.border}},right:{style:'thin',color:{rgb:C.border}}}};}
function sB(a,b){return{font:{name:'Arial',sz:10,bold:!!b,color:{rgb:C.text}},alignment:{horizontal:a||'left',vertical:'center'},border:{bottom:{style:'hair',color:{rgb:C.lightG}},left:{style:'hair',color:{rgb:C.lightG}},right:{style:'hair',color:{rgb:C.lightG}}}};}
function sBA(a,b){var s=sB(a,b);s.fill={fgColor:{rgb:C.light}};return s;}
function sBr(){return{font:{bold:true,color:{rgb:C.white},sz:14,name:'Arial'},fill:{fgColor:{rgb:C.navy}},alignment:{horizontal:'left',vertical:'center'}};}
function sSub(){return{font:{color:{rgb:'B0C4DE'},sz:10,name:'Arial'},fill:{fgColor:{rgb:C.navy}},alignment:{horizontal:'left',vertical:'center'}};}
function sNF(){return{fill:{fgColor:{rgb:C.navy}}};}
function sST(){return{font:{bold:true,color:{rgb:C.navy},sz:11,name:'Arial'},border:{bottom:{style:'medium',color:{rgb:C.green}}}};}
function sKL(){return{font:{bold:true,color:{rgb:C.muted},sz:8,name:'Arial'},fill:{fgColor:{rgb:C.light}},alignment:{horizontal:'center',vertical:'center'},border:{top:{style:'thin',color:{rgb:C.border}},left:{style:'thin',color:{rgb:C.border}},right:{style:'thin',color:{rgb:C.border}}}};}
function sKV(cl){return{font:{bold:true,color:{rgb:cl||C.text},sz:14,name:'Arial'},fill:{fgColor:{rgb:C.light}},alignment:{horizontal:'center',vertical:'center'},border:{bottom:{style:'thin',color:{rgb:C.border}},left:{style:'thin',color:{rgb:C.border}},right:{style:'thin',color:{rgb:C.border}}}};}

/* ===== HELPERS PLANILHA ===== */
function cL(n){var s='';while(n>=0){s=String.fromCharCode(65+(n%26))+s;n=Math.floor(n/26)-1;}return s;}
function cR(r,c){return cL(c)+String(r+1);}
function sC(ws,r,c,v,st){var ref=cR(r,c);ws[ref]={v:v,t:typeof v==='number'?'n':'s',s:st||sB()};if(!ws['!ref'])ws['!ref']='A1:'+ref;else{var rg=XLSX.utils.decode_range(ws['!ref']);if(r>rg.e.r)rg.e.r=r;if(c>rg.e.c)rg.e.c=c;ws['!ref']=XLSX.utils.encode_range(rg);}}
function addBH(ws,r,info,pd,tc){for(var i=0;i<tc;i++)sC(ws,r,i,'',sNF());sC(ws,r,0,'FORMULA CODE — AUDITORIA DE ESTOQUE',sBr());ws['!merges']=ws['!merges']||[];ws['!merges'].push({s:{r:r,c:0},e:{r:r,c:Math.min(3,tc-1)}});for(var i=0;i<tc;i++)sC(ws,r+1,i,'',sNF());sC(ws,r+1,0,'Cliente: '+(info.cliente||'—')+' | Unidade: '+(info.unidade||'—')+' | Inventário: '+(info.dataInventario||'—')+' | Processado: '+pd,sSub());ws['!merges'].push({s:{r:r+1,c:0},e:{r:r+1,c:Math.min(5,tc-1)}});return r+3;}
function addST(ws,r,t){sC(ws,r,0,t,sST());return r+1;}
function addKR(ws,r,lb,vl,cl){for(var i=0;i<lb.length;i++){sC(ws,r,i,lb[i],sKL());sC(ws,r+1,i,vl[i],sKV(cl&&cl[i]?cl[i]:C.text));}return r+3;}
function addDT(ws,r,hd,dr,ca){for(var i=0;i<hd.length;i++)sC(ws,r,i,hd[i],sH());r++;for(var x=0;x<dr.length;x++){var alt=x%2===1;for(var c=0;c<dr[x].length;c++){var v=dr[x][c],al=(ca&&ca[c])?ca[c]:'left',st=alt?sBA(al):sB(al);if(typeof v==='string'&&v.charAt(0)==='-'&&v.indexOf('R$')>0){st=JSON.parse(JSON.stringify(st));st.font.color={rgb:C.red};}sC(ws,r+x,c,v,st);}}return r+dr.length+1;}
function fxV(items){var f={},t=0;items.forEach(function(i){var k=i.faixa||'Sem giro';f[k]=(f[k]||0)+(i.valorEstoque||0);t+=i.valorEstoque||0;});return{f:f,t:t};}
function top20Cat(items){var m={};items.forEach(function(i){var c=i.categoria||'Sem categoria';if(!m[c])m[c]={nome:c,faltas:[],sobras:[],zerados:[]};if(i.difQtd<0)m[c].faltas.push(i);else if(i.difQtd>0)m[c].sobras.push(i);if(i.qtdContada===0&&i.qtdSistema>0)m[c].zerados.push(i);});Object.keys(m).forEach(function(k){m[k].faltas.sort(function(a,b){return a.difValor-b.difValor;}).splice(20);m[k].sobras.sort(function(a,b){return b.difValor-a.difValor;}).splice(20);m[k].zerados.sort(function(a,b){return(b.qtdSistema*b.custoUnit)-(a.qtdSistema*a.custoUnit);}).splice(20);});return Object.keys(m).sort().map(function(k){return m[k];});}

/* ===== RESUMOS EXECUTIVOS ===== */
function sumCritica(c){
  var w=c.categorias.filter(function(x){return x.nome!=='Sem categoria';}).sort(function(a,b){return a.acuracidade-b.acuracidade;});
  var desvio=Math.round((100-c.acuracidade)*10)/10;
  var t='A auditoria comparou '+NUM(c.totalSKUs)+' SKUs entre o saldo do sistema e a contagem física, apurando uma acuracidade de '+PCT(c.acuracidade)+' — ou seja, '+NUM(c.okCount)+' itens sem qualquer divergência. Foram registradas '+NUM(c.faltaCount)+' faltas (itens com saldo físico menor que o sistema), somando '+BRLi(c.totalFaltas)+' em valor não localizado, e '+NUM(c.sobraCount)+' sobras, no valor de '+BRLi(c.totalSobras)+'. O resultado é um saldo líquido de '+BRLi(c.saldoLiquido)+', que representa o impacto financeiro direto das divergências sobre o estoque registrado.';
  t+='\n\nUma acuracidade de '+PCT(c.acuracidade)+' indica que a cada 100 posições, cerca de '+desvio+' apresentam algum desvio — número que resume o tamanho financeiro da divergência apurada nos processos de entrada, transformação e saída de mercadoria.';
  if(w.length){
    var c1=w[0], nomes=c1.nome+' ('+PCT(c1.acuracidade)+')';
    if(w.length>1){nomes+=' e '+w[1].nome+' ('+PCT(w[1].acuracidade)+')';}
    t+=' As categorias com menor acuracidade foram '+nomes+', que concentram a maior fragilidade nos processos de entrada, transformação e saída de mercadoria e devem ser priorizadas no reforço de controle.';
  }
  t+=' Faltas em produtos perecíveis costumam apontar para inversão de códigos no registro de vendas ou perdas não registradas (quebra, vencimento, furto, desidratação), enquanto sobras sugerem falhas de lançamento na entrada.';
  return t;
}
function metCritica(){return'O universo da crítica considera apenas produtos com estoque de sistema positivo ou negativo, contagem física maior que zero, ou venda registrada no período — itens sem estoque, sem contagem e sem venda são excluídos da análise. Cada SKU qualificado é comparado entre o saldo registrado no sistema (ERP) e a contagem física realizada no inventário. A diferença (contado - sistema) determina a classificação: Falta (negativo), Sobra (positivo) ou Sem divergência (zero). O valor financeiro da divergência é calculado multiplicando a diferença pelo custo unitário do produto. A acuracidade é calculada pela razão entre o valor do estoque contado e o valor do estoque de sistema antes da contagem (e não pela quantidade de SKUs sem divergência), e a perda de estoque (%) mede o mesmo desvio com o sinal invertido, evidenciando a parcela de valor não localizada na contagem.';}
function sumRuptura(r){
  var t='Foram identificados '+NUM(r.totalRupturas)+' itens armazenados e não expostos — produtos com saldo no depósito, porém ausentes (quantidade zero) no salão de vendas —, o que representa uma taxa de ruptura de '+PCT(r.taxaRuptura)+' sobre os '+NUM(r.totalComDeposito)+' itens disponíveis em estoque. Cada item nessa condição é uma venda potencial perdida: o produto existe na loja, mas não está acessível ao consumidor na gôndola.';
  t+='\n\nDo total, '+NUM(r.rupturaA)+' itens são curva A por faturamento e '+NUM(r.rupturaB)+' são curva B — as faixas de maior giro, cuja ausência gera perda direta e imediata de receita, além da possibilidade de gerar insatisfação nos clientes da loja.';
  var wCat=r.categorias.filter(function(x){return x.rupturaA>0;}).sort(function(a,b){return b.rupturaA-a.rupturaA;});
  if(wCat.length)t+=' A categoria '+wCat[0].nome+' concentra o maior número de rupturas curva A, o que aponta para uma falha no processo de reposição desses produtos: vale investigar se o problema está na frequência de abastecimento da gôndola, na conferência de estoque ou no ponto de pedido.';
  t+=' Reduzir a ruptura dos itens A é a ação de retorno mais rápido, pois converte estoque parado em venda sem necessidade de nova compra.';
  return t;
}
function metRuptura(dias){dias=dias||90;return'Analisa-se a contagem física por local (depósito vs. loja/salão). Itens que possuem estoque no depósito mas quantidade zero na loja são classificados como ruptura — produto disponível no estoque que não está acessível ao consumidor. A classificação ABC é aplicada com base no faturamento e no lucro dos últimos '+dias+' dias, permitindo priorizar as rupturas de maior impacto financeiro.';}
function sumDias(d,dias){
  dias=dias||90;
  var t='A cobertura geral do estoque é de '+d.coberturaGeral+' dias';
  if(d.coberturaGeral>=16&&d.coberturaGeral<=30)t+=' (considerada adequada para o varejo alimentar)';
  t+=', o que indica por quantos dias o estoque atual sustentaria a venda no ritmo médio dos últimos '+dias+' dias. Essa média, porém, esconde desequilíbrios importantes entre as curvas: os itens de curva A cobrem apenas '+d.coberturaA+' dias';
  if(d.coberturaA<10)t+=' (alto risco de desabastecimento dos produtos mais vendidos)';
  t+=', enquanto a curva C cobre '+d.coberturaC+' dias — sinal de compra equivocada, com falta no que mais vende e excesso no que menos gira.';
  t+='\n\nA distribuição por faixa revela '+NUM(d.ruptura)+' itens já em ruptura, '+NUM(d.altoRisco)+' em alto risco (3–5 dias) e '+NUM(d.medioRisco)+' em risco médio — todos candidatos a reposição prioritária. No extremo oposto, '+NUM(d.excessos)+' SKUs estão em excesso de cobertura (acima de 30 dias), imobilizando '+BRLi(d.valorExcesso||0)+' em capital de giro que poderia ser liberado.';
  if(d.semGiro>0)t+=' Há ainda '+NUM(d.semGiro)+' itens sem giro nos últimos '+dias+' dias — estoque parado, que merece ter a causa investigada e solucionada o quanto antes.';
  t+=' O equilíbrio ideal passa por realocar capital da cauda (C e sem giro) para investir nos itens A.';
  return t;
}
function metDias(dias){dias=dias||90;return'A cobertura em dias é obtida dividindo o estoque atual pela venda média diária (venda dos últimos '+dias+' dias / '+dias+'). O resultado é classificado em faixas: Ruptura (0-2 dias), Alto risco (3-5 dias), Médio risco (6-15 dias), Cobertura ideal (16-30 dias), Excesso de cobertura (31+ dias) e Sem giro (nenhuma venda em '+dias+' dias). A cobertura por curva ABC pondera o estoque e a demanda de todos os itens daquela curva.';}
function sumABC(a,dias){
  dias=dias||90;
  var t='O estoque total inventariado soma '+BRLi(a.totalInvest)+', distribuído segundo o princípio de Pareto entre as três curvas. Os itens curva A por faturamento representam '+PCT(a.fatA.pctInvest)+' do capital em estoque e respondem por '+PCT(a.fatA.pctFat)+' do faturamento dos últimos '+dias+' dias';
  if(a.fatA.pctFat>a.fatA.pctInvest)t+=' (proporção saudável, pois pouco capital gera muita receita)';
  t+='. A curva B ocupa '+PCT(a.fatB.pctInvest)+' do estoque para '+PCT(a.fatB.pctFat)+' da receita, e a curva C consome '+PCT(a.fatC.pctInvest)+' do capital gerando apenas '+PCT(a.fatC.pctFat)+' do faturamento';
  if(a.fatC.pctInvest>a.fatC.pctFat+5)t+=' (sinalizando oportunidade de redução de estoque nessa faixa)';
  t+='.';
  t+='\n\nO cruzamento faturamento × lucro é o ponto mais estratégico da análise: a curva A por lucratividade não coincide integralmente com a curva A por faturamento. Itens que vendem muito (alto faturamento) mas operam com margem baixa aparecem em A no faturamento e caem para B/C no lucro — são produtos de "giro que não paga". Na direção inversa, itens de menor volume mas alta margem sobem para A no lucro: são os que mais contribuem para o resultado e merecem proteção contra ruptura. A curva A por lucro concentra '+PCT(a.lucA.pctLuc)+' do lucro total consumindo '+PCT(a.lucA.pctInvest)+' do capital.';
  t+='\n\nLeitura gerencial: o capital deve migrar da curva C (muito investimento, pouco retorno) para sustentar os itens A — priorizando, dentro de A, aqueles que são A tanto em faturamento quanto em lucro, pois combinam volume e rentabilidade. Uma redução planejada do estoque C libera capital de giro sem afetar receita relevante, e a atenção redobrada aos itens A-lucro protege a margem do negócio.';
  return t;
}
function metABC(dias){dias=dias||90;return'A classificação ABC ordena todos os SKUs pelo valor acumulado (faturamento ou lucro dos últimos '+dias+' dias). Os itens que representam até 80% do valor acumulado são classificados como curva A, de 80% a 95% como curva B, e os demais como curva C. O valor em estoque de cada item é calculado pela quantidade em estoque multiplicada pelo custo unitário (derivado do CMV quando não informado diretamente).';}
function sumPerda(p,dias){
  dias=dias||90;
  var t='Foram identificados '+NUM(p.totalSKUs)+' itens com venda registrada nos últimos '+dias+' dias que não constam na contagem física (zerados ou ausentes) — ou seja, produtos que comprovadamente vendiam e hoje não estão disponíveis. Com base na demanda média diária de cada um, projeta-se uma perda de '+BRLi(p.totalPerdaFat)+'/dia em faturamento e '+BRLi(p.totalPerdaLucro)+'/dia em lucro bruto. Estendendo ao mês, o impacto é de '+BRLi(p.perdaMensal)+' em receita não realizada — dinheiro que o negócio deixa de faturar enquanto o abastecimento não é regularizado.';
  if(p.classA.count>0){
    t+='\n\nA perda está fortemente concentrada: os itens curva A respondem por '+PCT(p.classA.pct)+' do total';
    var topP=p.items.filter(function(i){return i.abcFat==='A';}).slice(0,2);
    if(topP.length>=2)t+=', com destaque para '+topP[0].descricao+' e '+topP[1].descricao+' entre os maiores ofensores individuais';
    t+='. Essa concentração é uma boa notícia operacional — significa que regularizar poucos itens de alto impacto recupera a maior parte da perda. A curva B contribui com '+PCT(p.classB.pct)+' e a C com '+PCT(p.classC.pct)+'. A ação de maior retorno financeiro no curto prazo é repor com urgência os itens A ausentes, seguida da investigação da causa raiz (o item não foi comprado, foi perdido, ou há erro de contagem?).';
  }
  return t;
}
function metPerda(dias){dias=dias||90;return'Para cada item com venda registrada nos últimos '+dias+' dias e que não consta na contagem física (quantidade contada igual a zero ou ausente), projeta-se a venda perdida com base na demanda média diária. A perda diária de faturamento é a venda média diária em R$; a perda de lucro é o lucro médio diário. A projeção mensal multiplica esses valores por 30 dias. A premissa é que a demanda média dos últimos '+dias+' dias representa o padrão normal de consumo.';}

/* ===== GRÁFICOS (Chart.js → PNG, para uso em PDF/PPTX) =====
   r149: todos os gráficos mostram os valores (plugin datalabels).
   Cada gráfico tem uma função cfg*() que devolve a configuração do Chart.js — a mesma
   configuração serve para a tela (app.js) e para as exportações (PNG). */
var COR={navy:'#001528',green:'#5DC500',greenDk:'#2E7D32',red:'#D32F2F',amb:'#F57C00',yel:'#FBC02D',blue:'#1565C0',grey:'#888888',greyLt:'#9E9E9E',text:'#333333'};
/* Valor abreviado para rótulos de gráfico: R$ 850 · R$ 12,3 mil · R$ 1,25 mi */
function BRLk(v){
  var a=Math.abs(v||0),s=(v<0?'−':'')+'R$ ';
  if(a>=1e6)return s+(a/1e6).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:2})+' mi';
  if(a>=1e3)return s+(a/1e3).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+' mil';
  return s+Math.round(a).toLocaleString('pt-BR');
}
function _maxAbs(arr){var m=0;arr.forEach(function(v){if(v!==null&&v!==undefined&&Math.abs(v)>m)m=Math.abs(v);});return m||1;}
function _fonte(fs,bold){return{size:fs,weight:bold?'bold':'normal',family:'Arial, Helvetica, sans-serif'};}
function _cortar(t,n){t=String(t||'');return t.length>n?t.substring(0,n-1)+'…':t;}

function _chartPNG(config,wpx,hpx){
  if(typeof Chart==='undefined')return null;
  try{
    var canvas=document.createElement('canvas');
    canvas.width=wpx;canvas.height=hpx;
    var ctx=canvas.getContext('2d');
    config.options=config.options||{};
    config.options.responsive=false;
    config.options.animation=false;
    config.options.devicePixelRatio=2;
    /* fundo branco (PNG transparente fica escuro em alguns leitores de PDF/PPTX) */
    config.plugins=(config.plugins||[]).concat([{id:'fundoBranco',beforeDraw:function(ch){var c=ch.ctx;c.save();c.globalCompositeOperation='destination-over';c.fillStyle='#FFFFFF';c.fillRect(0,0,ch.width,ch.height);c.restore();}}]);
    var chart=new Chart(ctx,config);
    var url=chart.toBase64Image('image/png',1.0);
    chart.destroy();
    return url;
  }catch(e){console.log('Erro ao gerar gráfico:',e);return null;}
}

/* --- Crítica: barras divergentes (perdas à esquerda, sobras à direita, marca do saldo) --- */
function cfgCritica(c,fs){
  fs=fs||12;
  var cats=(c.hasCategorias&&c.categorias.length)?c.categorias:[{nome:'Total',faltaVal:c.totalFaltas,sobraVal:c.totalSobras,saldo:c.saldoLiquido}];
  var perdas=cats.map(function(x){return -Math.abs(x.faltaVal||0);});
  var sobras=cats.map(function(x){return Math.max(0,x.sobraVal||0);});
  var saldo=cats.map(function(x){return x.saldo||0;});
  var mx=_maxAbs(perdas.concat(sobras).concat(saldo));
  function dentro(v){return Math.abs(v)>=mx*0.16;}
  var dl={
    display:function(ctx){return ctx.dataset.data[ctx.dataIndex]!==0;},
    anchor:function(ctx){return dentro(ctx.dataset.data[ctx.dataIndex])?'center':'end';},
    align:function(ctx){return dentro(ctx.dataset.data[ctx.dataIndex])?'center':'end';},
    offset:4,clamp:true,
    color:function(ctx){return dentro(ctx.dataset.data[ctx.dataIndex])?'#FFFFFF':ctx.dataset.backgroundColor;},
    font:_fonte(fs-1,true),
    formatter:function(v){return BRLk(v);}
  };
  return {type:'bar',data:{
    labels:cats.map(function(x){return [_cortar(x.nome,28),'Saldo: '+BRLk(x.saldo||0)];}),
    datasets:[
      {type:'bar',label:'Perdas',data:perdas,backgroundColor:COR.red,stack:'div',borderRadius:3,barPercentage:.75,categoryPercentage:.85,datalabels:dl},
      {type:'bar',label:'Sobras',data:sobras,backgroundColor:COR.greenDk,stack:'div',borderRadius:3,barPercentage:.75,categoryPercentage:.85,datalabels:dl},
      {type:'line',label:'Saldo (quebra)',data:saldo,stack:'saldo',showLine:false,pointStyle:'rectRot',pointRadius:fs*0.6,pointHoverRadius:fs*0.6,backgroundColor:COR.navy,borderColor:'#FFFFFF',borderWidth:1.5,datalabels:{display:false}}
    ]},
    options:{indexAxis:'y',maintainAspectRatio:false,
      plugins:{legend:{position:'top',labels:{font:_fonte(fs),usePointStyle:true}},tooltip:{callbacks:{label:function(ctx){return ctx.dataset.label+': '+BRLk(ctx.raw);}}}},
      scales:{x:{stacked:true,min:-mx*1.25,max:mx*1.25,grid:{color:function(ctx){return ctx.tick&&ctx.tick.value===0?'#888888':'#EEEEEE';}},ticks:{font:_fonte(fs-2),callback:function(v){return BRLk(v);}}},
        y:{stacked:true,grid:{display:false},ticks:{font:_fonte(fs-1)}}},
      layout:{padding:{left:8,right:16}}}};
}
function chartCritica(c,wpx,hpx){
  var n=(c.hasCategorias&&c.categorias.length)?c.categorias.length:1;
  return _chartPNG(cfgCritica(c,13),wpx||1000,hpx||Math.min(760,Math.max(260,n*62+70)));
}
function hCritica(c,maxMm){var n=(c.hasCategorias&&c.categorias.length)?c.categorias.length:1;return Math.min(maxMm||135,Math.max(48,(n*62+70)/5.56));}

/* --- Ruptura: SKUs em ruptura por curva, com o valor do estoque em cima de cada barra --- */
function cfgRuptura(r,fs){
  fs=fs||12;
  var vals=[r.rupturaA,r.rupturaB,r.rupturaC];
  var vEst=[r.valorRupturaA,r.valorRupturaB,r.valorRupturaC];
  return {type:'bar',data:{labels:['Curva A','Curva B','Curva C'],datasets:[{label:'SKUs em ruptura',data:vals,backgroundColor:[COR.red,COR.amb,COR.grey],borderRadius:4,barPercentage:.6,
    datalabels:{anchor:'end',align:'end',offset:2,color:COR.text,textAlign:'center',font:_fonte(fs,true),
      formatter:function(v,ctx){var ve=vEst[ctx.dataIndex];return [NUM(v)+' SKUs'].concat(ve!==undefined&&ve!==null?['Estoque: '+BRLk(ve)]:[]);}}}]},
    options:{maintainAspectRatio:false,plugins:{legend:{display:false}},layout:{padding:{top:fs*3.4}},
      scales:{y:{beginAtZero:true,grace:'5%',ticks:{font:_fonte(fs-1),precision:0},title:{display:true,text:'SKUs em ruptura',font:_fonte(fs-1)}},x:{grid:{display:false},ticks:{font:_fonte(fs)}}}}};
}
function chartRuptura(r,wpx,hpx){return _chartPNG(cfgRuptura(r,13),wpx||1000,hpx||380);}

/* --- Ruptura: valor do estoque depósito (retaguarda) x área de vendas, barras empilhadas --- */
function _estCats(r){
  var cats=(r.estoquePorCategoria||[]).slice();
  var reais=cats.filter(function(x){return x.nome!=='Sem categoria';});
  if(!reais.length){return [{nome:'Total',deposito:r.valorDeposito||0,loja:r.valorLoja||0,total:(r.valorDeposito||0)+(r.valorLoja||0)}];}
  if(cats.length>12){var top=cats.slice(0,11),resto=cats.slice(11);var o={nome:'Outras ('+resto.length+')',deposito:0,loja:0,total:0};resto.forEach(function(x){o.deposito+=x.deposito;o.loja+=x.loja;o.total+=x.total;});top.push(o);cats=top;}
  return cats;
}
function cfgRupturaEstoque(r,fs){
  fs=fs||12;
  var cats=_estCats(r);
  var tot=_maxAbs(cats.map(function(x){return x.total;}));
  var dl=function(cor){return {display:function(ctx){var v=ctx.dataset.data[ctx.dataIndex];return v>0&&v>=tot*0.09;},anchor:'center',align:'center',color:cor,font:_fonte(fs-1,true),formatter:function(v){return BRLk(v);}};};
  return {type:'bar',data:{labels:cats.map(function(x){return [_cortar(x.nome,28),'Total: '+BRLk(x.total)];}),datasets:[
      {label:'Depósito (retaguarda)',data:cats.map(function(x){return x.deposito;}),backgroundColor:COR.navy,borderRadius:3,barPercentage:.75,datalabels:dl('#FFFFFF')},
      {label:'Área de vendas',data:cats.map(function(x){return x.loja;}),backgroundColor:COR.green,borderRadius:3,barPercentage:.75,datalabels:dl(COR.navy)}]},
    options:{indexAxis:'y',maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{font:_fonte(fs)}},tooltip:{callbacks:{label:function(ctx){return ctx.dataset.label+': '+BRLk(ctx.raw);}}}},
      scales:{x:{stacked:true,beginAtZero:true,ticks:{font:_fonte(fs-2),callback:function(v){return BRLk(v);}},grid:{color:'#EEEEEE'}},y:{stacked:true,grid:{display:false},ticks:{font:_fonte(fs-1)}}},
      layout:{padding:{right:16}}}};
}
function nEstCats(r){return _estCats(r).length;}
function chartRupturaEstoque(r,wpx,hpx){var n=nEstCats(r);return _chartPNG(cfgRupturaEstoque(r,13),wpx||1000,hpx||Math.min(760,Math.max(220,n*56+70)));}
function hRupturaEstoque(r,maxMm){var n=nEstCats(r);return Math.min(maxMm||135,Math.max(40,(n*56+70)/5.56));}

/* --- Dias de estoque: valor do estoque por potencial de ruptura (faixa de cobertura), gráfico de área --- */
var FAIXAS_DIAS=[{k:'Ruptura',l:'Ruptura (0–2 d)',c:COR.red},{k:'Alto risco',l:'Alto risco (3–5 d)',c:COR.amb},{k:'Médio risco',l:'Médio risco (6–15 d)',c:COR.yel},{k:'Cobertura ideal',l:'Ideal (16–30 d)',c:'#00B74A'},{k:'Excesso de cobertura',l:'Excesso (31+ d)',c:COR.blue},{k:'Sem giro',l:'Sem giro',c:COR.grey}];
function cfgDias(d,fs){
  fs=fs||12;
  var fv=fxV(d.items);
  var vals=FAIXAS_DIAS.map(function(f){return fv.f[f.k]||0;});
  return {type:'line',data:{labels:FAIXAS_DIAS.map(function(f){return f.l;}),datasets:[{label:'Valor do estoque (R$)',data:vals,fill:'origin',backgroundColor:'rgba(0,21,40,0.16)',borderColor:COR.navy,borderWidth:2.5,cubicInterpolationMode:'monotone',
      pointRadius:fs*0.5,pointBackgroundColor:FAIXAS_DIAS.map(function(f){return f.c;}),pointBorderColor:'#FFFFFF',pointBorderWidth:1.5,
      datalabels:{anchor:'end',align:'top',offset:4,color:COR.navy,textAlign:'center',font:_fonte(fs,true),formatter:function(v){return [BRLk(v),PCT(fv.t?v/fv.t*100:0)];}}}]},
    options:{maintainAspectRatio:false,plugins:{legend:{display:false}},layout:{padding:{top:fs*3.2,left:10,right:fs*4.5}},
      scales:{y:{beginAtZero:true,grace:'8%',ticks:{font:_fonte(fs-1),callback:function(v){return BRLk(v);}},title:{display:true,text:'Valor do estoque',font:_fonte(fs-1)}},x:{grid:{display:false},ticks:{font:_fonte(fs-1)}}}}};
}
function chartDias(d,wpx,hpx){return _chartPNG(cfgDias(d,13),wpx||1000,hpx||380);}

/* --- Investimento ABC: Pareto (área) do valor do estoque, com faixas A, B, C e Sem giro --- */
var FAIXAS_ABC=[{k:'A',l:'Curva A',c:COR.navy,f:'rgba(0,21,40,0.55)'},{k:'B',l:'Curva B',c:COR.green,f:'rgba(93,197,0,0.55)'},{k:'C',l:'Curva C',c:COR.amb,f:'rgba(245,124,0,0.55)'},{k:'SG',l:'Sem giro',c:COR.greyLt,f:'rgba(158,158,158,0.55)'}];
function _grupoABC(it){return it.semGiro?'SG':(it.abcFat||'C');}
function valoresCurvasABC(a){
  var sg=a.semGiro||{invest:0,pctInvest:0},cg=a.fatCg||a.fatC;
  return {A:{v:a.fatA.invest,p:a.fatA.pctInvest},B:{v:a.fatB.invest,p:a.fatB.pctInvest},C:{v:cg.invest,p:cg.pctInvest},SG:{v:sg.invest,p:sg.pctInvest}};
}
function cfgABCPareto(a,fs){
  fs=fs||12;
  var ordem={A:0,B:1,C:2,SG:3};
  var its=a.items.filter(function(i){return i.valorInvestido>0;}).slice().sort(function(x,y){var gx=ordem[_grupoABC(x)],gy=ordem[_grupoABC(y)];if(gx!==gy)return gx-gy;if((y.fat90||0)!==(x.fat90||0))return (y.fat90||0)-(x.fat90||0);return y.valorInvestido-x.valorInvestido;});
  var N=its.length,tot=its.reduce(function(s,i){return s+i.valorInvestido;},0)||1;
  var vc=valoresCurvasABC(a);
  var passo=Math.max(1,Math.ceil(N/400));
  var cum=0,pts={A:[],B:[],C:[],SG:[]},ultimo={x:0,y:0};
  var grupoAtual=null;
  its.forEach(function(it,idx){
    var g=_grupoABC(it);
    if(g!==grupoAtual){pts[g].push({x:ultimo.x,y:ultimo.y});grupoAtual=g;}
    cum+=it.valorInvestido;
    var p={x:(idx+1)/N*100,y:cum/tot*100};
    var fimGrupo=(idx===N-1)||_grupoABC(its[idx+1])!==g;
    if(fimGrupo||((idx+1)%passo===0))pts[g].push(p);
    ultimo=p;
  });
  var datasets=[],rotulos=[];
  FAIXAS_ABC.forEach(function(f){
    var d=pts[f.k];if(d.length<2)return;
    var v=vc[f.k];
    datasets.push({type:'line',label:f.l+' — '+BRLk(v.v)+' ('+PCT(v.p)+')',data:d,borderColor:f.c,backgroundColor:f.f,fill:'origin',borderWidth:2,pointRadius:0,tension:0,datalabels:{display:false}});
    var x0=d[0].x,x1=d[d.length-1].x,xm=(x0+x1)/2,ym=0;
    for(var i=0;i<d.length;i++){if(d[i].x>=xm){ym=d[i].y;break;}}
    if(x1-x0>=9)rotulos.push({x:xm,y:Math.max(6,ym/2),t:[f.l,BRLk(v.v),PCT(v.p)],c:f.k==='B'?COR.navy:'#FFFFFF'});
  });
  datasets.push({type:'scatter',label:'_rotulos',data:rotulos.map(function(r){return{x:r.x,y:r.y};}),pointRadius:0,pointHoverRadius:0,
    datalabels:{display:true,align:'center',anchor:'center',textAlign:'center',font:_fonte(fs,true),color:function(ctx){return rotulos[ctx.dataIndex].c;},formatter:function(v,ctx){return rotulos[ctx.dataIndex].t;}}});
  return {type:'line',data:{datasets:datasets},options:{maintainAspectRatio:false,parsing:false,
    plugins:{legend:{position:'top',labels:{font:_fonte(fs-1),filter:function(it){return it.text!=='_rotulos';}}},tooltip:{enabled:false}},
    scales:{x:{type:'linear',min:0,max:100,ticks:{font:_fonte(fs-2),callback:function(v){return v+'%';}},title:{display:true,text:'% dos SKUs com estoque (ordenados por curva)',font:_fonte(fs-1)},grid:{color:'#EEEEEE'}},
      y:{min:0,max:100,ticks:{font:_fonte(fs-2),callback:function(v){return v+'%';}},title:{display:true,text:'% acumulado do valor do estoque',font:_fonte(fs-1)},grid:{color:'#EEEEEE'}}},
    layout:{padding:{right:16}}}};
}
function chartABC(a,wpx,hpx){return _chartPNG(cfgABCPareto(a,13),wpx||1000,hpx||440);}

/* --- Projeção de perda: perda projetada no mês e valor do estoque (em sistema) dos itens, por curva --- */
function estoqueSistemaPerda(p,critica){
  if(!critica||!critica.items)return null;
  var m={};critica.items.forEach(function(i){m[i.sku]=(i.qtdSistema||0)*(i.custoUnit||0);});
  var r={A:0,B:0,C:0};p.items.forEach(function(i){var v=m[i.sku];if(v)r[i.abcFat||'C']+=v;});
  return r;
}
function cfgPerda(p,critica,fs){
  fs=fs||12;
  var est=estoqueSistemaPerda(p,critica);
  var dl={anchor:'end',align:'end',offset:2,color:COR.text,font:_fonte(fs-1,true),formatter:function(v){return BRLk(v);}};
  var ds=[];
  if(est)ds.push({label:'Valor do estoque em sistema (não localizado na contagem)',data:[est.A,est.B,est.C],backgroundColor:COR.navy,borderRadius:4,datalabels:dl});
  ds.push({label:'Perda projetada no mês (faturamento)',data:[p.classA.perda*30,p.classB.perda*30,p.classC.perda*30],backgroundColor:COR.red,borderRadius:4,datalabels:dl});
  return {type:'bar',data:{labels:['Curva A','Curva B','Curva C'],datasets:ds},
    options:{maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{font:_fonte(fs-1)}}},layout:{padding:{top:fs*1.6}},
      scales:{y:{beginAtZero:true,grace:'8%',ticks:{font:_fonte(fs-1),callback:function(v){return BRLk(v);}}},x:{grid:{display:false},ticks:{font:_fonte(fs)}}}}};
}
function chartPerda(p,critica,wpx,hpx){return _chartPNG(cfgPerda(p,critica,13),wpx||1000,hpx||380);}

/* --- Comparativo: gráficos entre unidades --- */
function _compUnid(comp,campos){return comp.unidades.filter(function(u){return campos.some(function(k){return u[k]!==null&&u[k]!==undefined;});});}
function cfgCompQuebra(comp,fs){
  /* r151: colunas verticais divergentes — sobras para cima, perdas para baixo, marca da quebra */
  fs=fs||12;
  var us=_compUnid(comp,['totalFaltas','totalSobras','saldoLiquido']);
  var perdas=us.map(function(u){return -Math.abs(u.totalFaltas||0);});
  var sobras=us.map(function(u){return Math.max(0,u.totalSobras||0);});
  var saldo=us.map(function(u){return u.saldoLiquido||0;});
  var mx=_maxAbs(perdas.concat(sobras).concat(saldo));
  function dentro(v){return Math.abs(v)>=mx*0.12;}
  var dl={display:function(ctx){return ctx.dataset.data[ctx.dataIndex]!==0;},
    anchor:function(ctx){return dentro(ctx.dataset.data[ctx.dataIndex])?'center':'end';},
    align:function(ctx){return dentro(ctx.dataset.data[ctx.dataIndex])?'center':'end';},
    offset:4,clamp:true,color:function(ctx){return dentro(ctx.dataset.data[ctx.dataIndex])?'#FFFFFF':ctx.dataset.backgroundColor;},
    font:_fonte(fs,true),formatter:function(v){return BRLk(v);}};
  return {type:'bar',data:{labels:us.map(function(u){return [_cortar(u.unidade,26),'Quebra: '+BRLk(u.saldoLiquido||0)];}),datasets:[
      {type:'bar',label:'Sobras',data:sobras,backgroundColor:COR.greenDk,stack:'div',borderRadius:3,barPercentage:.6,categoryPercentage:.8,datalabels:dl},
      {type:'bar',label:'Perdas',data:perdas,backgroundColor:COR.red,stack:'div',borderRadius:3,barPercentage:.6,categoryPercentage:.8,datalabels:dl},
      {type:'line',label:'Quebra (saldo)',data:saldo,stack:'saldo',showLine:false,pointStyle:'rectRot',pointRadius:fs*0.7,pointHoverRadius:fs*0.7,backgroundColor:COR.navy,borderColor:'#FFFFFF',borderWidth:1.5,
        datalabels:{display:false}}]},
    options:{maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{font:_fonte(fs),usePointStyle:true}}},layout:{padding:{top:6,bottom:4,right:20}},
      scales:{x:{stacked:true,grid:{display:false},ticks:{font:_fonte(fs-1)}},
        y:{stacked:true,min:-mx*1.2,max:mx*1.2,ticks:{font:_fonte(fs-2),callback:function(v){return BRLk(v);}},grid:{color:function(ctx){return ctx.tick&&ctx.tick.value===0?'#888888':'#EEEEEE';}}}}}};
}
function _dlPilha(cor,fs,getTot){return {display:function(ctx){var v=ctx.dataset.data[ctx.dataIndex];var t=getTot(ctx.dataIndex);return v>0&&t>0&&v/t>=0.07;},anchor:'center',align:'center',color:cor,font:_fonte(fs-1,true),formatter:function(v){return BRLk(v);}};}
function cfgCompDepLoja(comp,fs){
  fs=fs||12;
  var us=_compUnid(comp,['valorDeposito','valorLoja']);
  var tots=us.map(function(u){return (u.valorDeposito||0)+(u.valorLoja||0);});
  var gt=function(i){return tots[i];};
  return {type:'bar',data:{labels:us.map(function(u){return [_cortar(u.unidade,26),'Total: '+BRLk((u.valorDeposito||0)+(u.valorLoja||0))];}),datasets:[
      {label:'Depósito (retaguarda)',data:us.map(function(u){return u.valorDeposito||0;}),backgroundColor:COR.navy,datalabels:_dlPilha('#FFFFFF',fs,gt)},
      {label:'Área de vendas (loja)',data:us.map(function(u){return u.valorLoja||0;}),backgroundColor:COR.green,datalabels:_dlPilha(COR.navy,fs,gt)}]},
    options:{maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{font:_fonte(fs)}}},
      scales:{y:{stacked:true,beginAtZero:true,ticks:{font:_fonte(fs-2),callback:function(v){return BRLk(v);}},grid:{color:'#EEEEEE'}},x:{stacked:true,grid:{display:false},ticks:{font:_fonte(fs-1)}}}}};
}
function cfgCompPotencial(comp,fs){
  fs=fs||12;
  var us=_compUnid(comp,['potRuptura','potAltoRisco','potMedioRisco']);
  var tots=us.map(function(u){return (u.potRuptura||0)+(u.potAltoRisco||0)+(u.potMedioRisco||0);});
  var gt=function(i){return tots[i];};
  var defs=[{k:'potRuptura',l:'Ruptura (0–2 dias)',c:COR.red,t:'#FFFFFF'},{k:'potAltoRisco',l:'Alto risco (3–5 dias)',c:COR.amb,t:'#FFFFFF'},{k:'potMedioRisco',l:'Médio risco (6–15 dias)',c:COR.yel,t:COR.navy}];
  return {type:'bar',data:{labels:us.map(function(u,i){return [_cortar(u.unidade,26),'Total: '+BRLk(tots[i])+'/mês'];}),datasets:defs.map(function(d){return {label:d.l,data:us.map(function(u){return u[d.k]||0;}),backgroundColor:d.c,datalabels:_dlPilha(d.t,fs,gt)};})},
    options:{maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{font:_fonte(fs)}}},
      scales:{y:{stacked:true,beginAtZero:true,ticks:{font:_fonte(fs-2),callback:function(v){return BRLk(v);}},grid:{color:'#EEEEEE'},title:{display:true,text:'Faturamento mensal em risco',font:_fonte(fs-1)}},x:{stacked:true,grid:{display:false},ticks:{font:_fonte(fs-1)}}}}};
}
function cfgCompCurvas(comp,fs){
  fs=fs||12;
  var us=_compUnid(comp,['estA','estB','estC','estSemGiro']);
  var tots=us.map(function(u){return (u.estA||0)+(u.estB||0)+(u.estC||0)+(u.estSemGiro||0);});
  var gt=function(i){return tots[i];};
  var defs=[{k:'estA',l:'Curva A',c:COR.navy,t:'#FFFFFF'},{k:'estB',l:'Curva B',c:COR.green,t:COR.navy},{k:'estC',l:'Curva C',c:COR.amb,t:'#FFFFFF'},{k:'estSemGiro',l:'Sem giro',c:COR.greyLt,t:'#FFFFFF'}];
  return {type:'bar',data:{labels:us.map(function(u,i){return [_cortar(u.unidade,26),'Total: '+BRLk(tots[i])];}),datasets:defs.map(function(d){return {label:d.l,data:us.map(function(u){return u[d.k]||0;}),backgroundColor:d.c,datalabels:_dlPilha(d.t,fs,gt)};})},
    options:{maintainAspectRatio:false,plugins:{legend:{position:'top',labels:{font:_fonte(fs)}}},
      scales:{y:{stacked:true,beginAtZero:true,ticks:{font:_fonte(fs-2),callback:function(v){return BRLk(v);}},grid:{color:'#EEEEEE'}},x:{stacked:true,grid:{display:false},ticks:{font:_fonte(fs-1)}}}}};
}

/* ===== GERAR EXCEL ===== */
/* Monta o workbook (sem baixar) — usado pelo download e pela gravacao no Drive. */
function buildExcelWorkbook(data,sel,pd,info){
  info=info||{};var wb=XLSX.utils.book_new();
  return _buildWbBody(wb,data,sel,pd,info);
}
/* Retorna o Excel como Blob (para salvar no Drive). Inclui TODAS as analises. */
function generateExcelBlob(data,pd,info){
  var selTudo={criticaResumo:true,criticaDetalhe:true,ruptura:true,dias:true,abc:true,perda:true};
  var wb=buildExcelWorkbook(data,selTudo,pd,info);
  var out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  return new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function generateExcel(data,sel,pd,info){
  info=info||{};var wb=XLSX.utils.book_new();
  _buildWbBody(wb,data,sel,pd,info);
  var out=XLSX.write(wb,{bookType:'xlsx',type:'array'});var blob=new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});var url=URL.createObjectURL(blob);var a=document.createElement('a');a.href=url;a.download='auditoria_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.xlsx';a.click();URL.revokeObjectURL(url);
}
function _buildWbBody(wb,data,sel,pd,info){
  /* DASHBOARD */
  var ws={},R=0,DC=8;
  R=addBH(ws,R,info,pd,DC);
  if(data.critica){var c=data.critica;R=addST(ws,R,'CRÍTICA DO INVENTÁRIO');R=addKR(ws,R,['ACURACIDADE','VALOR ESTOQUE','VALOR ESTOQUE CONTADO','VALOR DAS FALTAS','VALOR DAS SOBRAS','SALDO LÍQUIDO','PERDA DE ESTOQUE (%)'],[PCT(c.acuracidade),BRLi(c.valorEstoque),BRLi(c.valorEstoqueContado),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido),PCT(c.perdaEstoquePct)],[C.green,C.text,C.text,C.red,C.amb,C.red,c.perdaEstoquePct<0?C.red:C.green]);if(c.hasCategorias)R=addDT(ws,R,['Categoria','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{0:'left',1:'right',2:'right',3:'right',4:'right'});}
  if(data.ruptura){var r=data.ruptura;R=addST(ws,R,'RUPTURA LOJA X DEPÓSITO');R=addKR(ws,R,['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CURVA A (FAT.)','RUPTURA CURVA A (LUCRO)'],[PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],[C.red,C.text,C.red,C.red]);}
  if(data.dias){var d=data.dias,fv=fxV(d.items);R=addST(ws,R,'DIAS DE ESTOQUE');R=addKR(ws,R,['COBERTURA GERAL','CURVA A','CURVA B','CURVA C','SEM GIRO'],[d.coberturaGeral+' dias',d.coberturaA+' dias',d.coberturaB+' dias',d.coberturaC+' dias',NUM(d.semGiro)],[C.text,C.text,C.text,C.text,C.red]);var fo=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];R=addDT(ws,R,['Faixa','SKUs','% SKUs','Valor Estoque (R$)','% do Valor'],fo.map(function(f){var cn=d.items.filter(function(i){return i.faixa===f;}).length;var vl=fv.f[f]||0;return[f,cn,PCT(d.total?cn/d.total*100:0),BRLi(vl),PCT(fv.t?vl/fv.t*100:0)];}),{0:'left',1:'right',2:'right',3:'right',4:'right'});}
  if(data.abc){var a=data.abc;R=addST(ws,R,'INVESTIMENTO ABC');R=addKR(ws,R,['VALOR TOTAL EM ESTOQUE','FATURAMENTO 90D','LUCRO 90D','SKUS'],[BRLi(a.totalInvest),BRLi(a.totalFat),BRLi(a.totalLucro),NUM(a.items.length)],[C.text,C.green,C.green,C.text]);R=addDT(ws,R,['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],{0:'center',1:'right',2:'right',3:'right',4:'right'});}
  if(data.perda){var pe=data.perda;R=addST(ws,R,'PROJEÇÃO DE PERDA');R=addKR(ws,R,['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS EM RUPTURA'],[BRLi(pe.totalPerdaFat),BRLi(pe.totalPerdaLucro),BRLi(pe.perdaMensal),NUM(pe.totalSKUs)],[C.red,C.red,C.red,C.text]);R=addDT(ws,R,['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',pe.classA.count,BRLi(pe.classA.perda),BRLi(pe.classA.lucro),PCT(pe.classA.pct),BRLi(pe.classA.perda*30)],['B',pe.classB.count,BRLi(pe.classB.perda),BRLi(pe.classB.lucro),PCT(pe.classB.pct),BRLi(pe.classB.perda*30)],['C',pe.classC.count,BRLi(pe.classC.perda),BRLi(pe.classC.lucro),PCT(pe.classC.pct),BRLi(pe.classC.perda*30)]],{0:'center',1:'right',2:'right',3:'right',4:'right',5:'right'});}
  ws['!cols']=[{wch:28},{wch:18},{wch:16},{wch:18},{wch:16},{wch:18},{wch:16},{wch:16}];ws['!rows']=[{hpt:28},{hpt:20}];
  XLSX.utils.book_append_sheet(wb,ws,'Dashboard');

  /* CRITICA RESUMO + TOP20 */
  if(sel.criticaResumo&&data.critica){var wsC={},rw=0,c=data.critica;rw=addBH(wsC,rw,info,pd,6);rw=addST(wsC,rw,'RESUMO DA CRÍTICA');var sL=['ACURACIDADE','Valor estoque (sistema)','Valor estoque contado','Valor das faltas','Valor das sobras','Saldo líquido','Perda de estoque (%)'],sV=[PCT(c.acuracidade),BRLi(c.valorEstoque),BRLi(c.valorEstoqueContado),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido),PCT(c.perdaEstoquePct)];for(var i=0;i<sL.length;i++){sC(wsC,rw+i,0,sL[i],sB('left',true));sC(wsC,rw+i,1,sV[i],sB('right'));}rw+=sL.length+1;
  if(c.hasCategorias){rw=addST(wsC,rw,'RESULTADO POR CATEGORIA');rw=addDT(wsC,rw,['Categoria','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{0:'left',1:'right',2:'right',3:'right',4:'right'});}
  var ct=top20Cat(c.items);var tH=['SKU','Descrição','Qtd Sist','Qtd Contada','Dif. Qtd','Dif. R$'],tA={0:'left',1:'left',2:'right',3:'right',4:'right',5:'right'};
  ct.forEach(function(cat){if(cat.faltas.length){rw=addST(wsC,rw,'TOP '+cat.faltas.length+' FALTAS — '+cat.nome);rw=addDT(wsC,rw,tH,cat.faltas.map(function(i){return[i.sku,i.descricao,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor)];}),tA);}if(cat.sobras.length){rw=addST(wsC,rw,'TOP '+cat.sobras.length+' SOBRAS — '+cat.nome);rw=addDT(wsC,rw,tH,cat.sobras.map(function(i){return[i.sku,i.descricao,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor)];}),tA);}if(cat.zerados.length){rw=addST(wsC,rw,'TOP '+cat.zerados.length+' ZERADOS — '+cat.nome);rw=addDT(wsC,rw,['SKU','Descrição','Qtd Sistema','Valor Perdido'],cat.zerados.map(function(i){return[i.sku,i.descricao,i.qtdSistema,BRL(i.qtdSistema*i.custoUnit)];}),{0:'left',1:'left',2:'right',3:'right'});}});
  wsC['!cols']=[{wch:16},{wch:32},{wch:14},{wch:14},{wch:12},{wch:16}];wsC['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsC,'Crítica - Resumo');}

  if(sel.criticaDetalhe&&data.critica){var wsCD={},rw=0;rw=addBH(wsCD,rw,info,pd,8);rw=addDT(wsCD,rw,['SKU','Descrição','Categoria','Qtd Sistema','Qtd Contada','Dif. Qtd','Dif. R$','Status'],data.critica.items.map(function(i){return[i.sku,i.descricao,i.categoria,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor),i.status];}),{0:'left',1:'left',2:'left',3:'right',4:'right',5:'right',6:'right',7:'center'});wsCD['!cols']=[{wch:14},{wch:32},{wch:18},{wch:12},{wch:12},{wch:10},{wch:14},{wch:10}];wsCD['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsCD,'Crítica - Detalhado');}

  if(sel.ruptura&&data.ruptura){var wsR={},rw=0;rw=addBH(wsR,rw,info,pd,9);rw=addDT(wsR,rw,['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Depósito','Qtd Loja','Venda Méd/Dia','Fat. Méd/Dia'],data.ruptura.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abc_valorVendido90||'C',i.abc_lucro90||'C',i.deposito,i.loja,R2(i.vendaMediaDia),BRL(i.fatMediaDia||0)];}),{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right'});wsR['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:14},{wch:10},{wch:14},{wch:14}];wsR['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsR,'Ruptura Loja x Depósito');}

  if(sel.dias&&data.dias){var wsD={},rw=0;rw=addBH(wsD,rw,info,pd,9);rw=addDT(wsD,rw,['SKU','Descrição','Categoria','Qtd Estoque','Venda Méd/Dia','Dias Estoque','Cobertura','Valor Estoque','ABC Fat.'],data.dias.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.qtdEstoque,R2(i.vendaMediaDia),i.diasEstoque!==null?R2(i.diasEstoque):'—',i.faixa,BRL(i.valorEstoque),i.abcFat];}),{0:'left',1:'left',2:'left',3:'right',4:'right',5:'right',6:'left',7:'right',8:'center'});wsD['!cols']=[{wch:14},{wch:32},{wch:18},{wch:12},{wch:14},{wch:12},{wch:18},{wch:16},{wch:10}];wsD['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsD,'Dias de Estoque');}

  if(sel.abc&&data.abc){var wsA={},rw=0;rw=addBH(wsA,rw,info,pd,10);rw=addDT(wsA,rw,['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Estoque','Custo Unit.','Valor Estoque','Fat. 90 dias','Lucro 90 dias'],data.abc.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abcFat,i.abcLucro,i.qtdEstoque,BRL(i.custoUnit),BRL(i.valorInvestido),BRL(i.fat90),BRL(i.lucro90)];}),{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right',9:'right'});wsA['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:12},{wch:14},{wch:16},{wch:16},{wch:16}];wsA['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsA,'Investimento ABC');}

  if(sel.perda&&data.perda){var wsP={},rw=0;rw=addBH(wsP,rw,info,pd,11);rw=addDT(wsP,rw,['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Venda 90d','Venda Méd/Dia','Perda Fat./Dia','Perda Lucro/Dia','Perda Fat./Mês','Perda Lucro/Mês'],data.perda.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abcFat,i.abcLucro||'C',R2(i.qtdVendida||0),R2(i.vendaMediaDia),BRL(i.perdaFatDia),BRL(i.perdaLucroDia),BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];}),{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right',9:'right',10:'right'});wsP['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:12},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14}];wsP['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsP,'Projeção de Perda');}

  return wb;
}

/* ========== PDF COM IA ========== */
function generatePDF(rt,data,pd,logo,info){
  // Tentar obter resumo via IA antes de gerar
  var st=window.App?window.App.getState():null;
  var apiUrl=st&&st.apiUrl?st.apiUrl:null;
  var apiToken=st&&st.apiToken?st.apiToken:null;
  window._iaResumos=window._iaResumos||{};
  var secaoIA=(rt==='dias')?'dias_estoque':rt; /* a IA usa a chave "dias_estoque"; internamente usamos "dias" */
  /* Reaproveita o texto já buscado (em segundo plano) pela própria tela, para qualquer dimensão — evita nova chamada. */
  if(st&&st.iaAnalise&&st.iaAnalise[rt]){
    window._iaResumos[secaoIA]=st.iaAnalise[rt];
  }
  if(apiUrl&&apiToken&&data[rt]&&!window._iaResumos[secaoIA]){
    var metricas=Engine.buildMetricasIA(rt,data[rt]);
    try{
      var xhr=new XMLHttpRequest();
      xhr.open('POST',apiUrl,false);// síncrono
      xhr.send(JSON.stringify({action:'gerarResumoInventarioIA',token:apiToken,cliente:info.cliente||'',unidade:info.unidade||'',data:info.dataInventario||'',diasVenda:info.diasVenda||90,metricas:metricas,secao:secaoIA}));
      var resp=JSON.parse(xhr.responseText);
      if(resp.ok&&resp.resumos&&resp.resumos[secaoIA]){
        window._iaResumos[secaoIA]=resp.resumos[secaoIA];
      }
    }catch(e){console.log('IA fallback:',e);}
  }
  _generatePDFInternal(rt,data,pd,logo,info);
}
function _generatePDFInternal(rt,data,pd,logo,info){
  info=info||{};var jsPDF=window.jspdf.jsPDF;var doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});var W=210,H=297,M=15,y=0;
  function hdr(){doc.setFillColor(0,21,40);doc.rect(0,0,W,22,'F');if(logo){try{doc.addImage(logo,'PNG',M,7,32,8);}catch(e){}}doc.setFontSize(9);doc.setTextColor(255,255,255);doc.text((info.cliente||'')+' — '+(info.unidade||''),W-M,7,{align:'right'});doc.setFontSize(7);doc.setTextColor(200,220,255);doc.text('Inventário: '+(info.dataInventario||'—'),W-M,12,{align:'right'});doc.setTextColor(180,180,200);doc.text('Processado em '+pd,W-M,17,{align:'right'});y=28;}
  function ftr(pg){doc.setFontSize(7);doc.setTextColor(150,150,150);doc.text('Formula Code Tecnologia, Gestão e Automação',M,H-6);doc.text('Página '+pg,W-M,H-6,{align:'right'});doc.setDrawColor(200,200,200);doc.line(M,H-10,W-M,H-10);}
  function chk(n){if(y+n>H-18){doc.addPage();hdr();ftr(doc.getNumberOfPages());}}
  function ttl(t){chk(12);doc.setFontSize(14);doc.setTextColor(0,21,40);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFontSize(8);doc.setTextColor(150,150,150);doc.setFont(undefined,'normal');doc.text('Relatório gerado automaticamente pelo sistema Formula Code',M,y);y+=8;}
  function sec(t){chk(26);doc.setFontSize(11);doc.setTextColor(0,21,40);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFont(undefined,'normal');}
  /* r149: tabelas que continuam na página seguinte respeitam as margens e ganham cabeçalho/rodapé */
  function novaPagTab(dd){if(dd.pageNumber>1){var yy=y;hdr();ftr(doc.getNumberOfPages());y=yy;}}
  function aT(h,b,o){chk(20);doc.autoTable({startY:y,head:[h],body:b,margin:{left:M,right:M,top:28,bottom:18},didDrawPage:novaPagTab,headStyles:{fillColor:[0,21,40],fontSize:7,fontStyle:'bold',halign:'left'},bodyStyles:{fontSize:7,halign:'left'},alternateRowStyles:{fillColor:[245,245,245]},styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},columnStyles:o||{}});y=doc.lastAutoTable.finalY+6;}
  function kpi(lb,vl,cl){chk(18);var cw=(W-2*M)/lb.length;doc.setFillColor(245,245,245);doc.roundedRect(M,y-2,W-2*M,16,2,2,'F');for(var i=0;i<lb.length;i++){var x=M+i*cw+4;doc.setFontSize(7);doc.setTextColor(150,150,150);doc.setFont(undefined,'bold');doc.text(lb[i],x,y+3);doc.setFontSize(11);doc.setFont(undefined,'bold');var cc=cl[i]||[51,51,51];doc.setTextColor(cc[0],cc[1],cc[2]);doc.text(String(vl[i]),x,y+10);}doc.setFont(undefined,'normal');y+=20;}
  function bloco(txt){chk(16);doc.setFontSize(8);doc.setTextColor(80,80,80);doc.setFont(undefined,'normal');var lines=doc.splitTextToSize(txt,W-2*M);doc.text(lines,M,y);y+=lines.length*3.5+4;}
  function img(url,h){if(!url)return;chk(h+8);try{doc.addImage(url,'PNG',M,y,W-2*M,h);}catch(e){}y+=h+8;}
  /* r151: soma de coluna e formato de quantidade para as linhas de total */
  function soma(arr,fn){return arr.reduce(function(t,i){return t+(Number(fn(i))||0);},0);}
  function Q(v){return (Math.round((v||0)*100)/100).toLocaleString('pt-BR',{maximumFractionDigits:2});}
  /* r149: título + gráfico sempre na mesma página */
  function secImg(t,url,h){if(!url)return;chk(h+16);sec(t);img(url,h);}
  /* r149: tabela com linha de totalização (rodapé em destaque) */
  function aTT(h,b,tot,o){chk(20);doc.autoTable({startY:y,head:[h],body:b,foot:[tot],showFoot:'lastPage',margin:{left:M,right:M,top:28,bottom:18},didDrawPage:novaPagTab,headStyles:{fillColor:[0,21,40],fontSize:7,fontStyle:'bold',halign:'left'},footStyles:{fillColor:[232,238,245],textColor:[0,21,40],fontSize:7,fontStyle:'bold'},bodyStyles:{fontSize:7,halign:'left'},alternateRowStyles:{fillColor:[245,245,245]},styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},columnStyles:o||{},didParseCell:function(dd){if(dd.section==='foot'&&o&&o[dd.column.index]&&o[dd.column.index].halign)dd.cell.styles.halign=o[dd.column.index].halign;}});y=doc.lastAutoTable.finalY+6;}

  hdr();ftr(1);

  if(rt==='critica'){
    var c=data.critica;
    ttl('Crítica do inventário — Resumo executivo');
    var iaR=window._iaResumos&&window._iaResumos.critica;
    sec('Análise');bloco(iaR||Engine.gerarAnaliseCritica(c,info));
    sec('Metodologia');bloco(metCritica());
    sec('Indicadores gerais');
    kpi(['ACURACIDADE','VALOR ESTOQUE','VALOR ESTOQUE CONTADO','PERDA DE ESTOQUE'],[PCT(c.acuracidade),BRLi(c.valorEstoque),BRLi(c.valorEstoqueContado),PCT(c.perdaEstoquePct)],[[0,183,74],[51,51,51],[51,51,51],c.perdaEstoquePct<0?[211,47,47]:[0,183,74]]);
    kpi(['VALOR DAS FALTAS','VALOR DAS SOBRAS','SALDO LÍQUIDO'],[BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido)],[[211,47,47],[245,124,0],[211,47,47]]);
    var hCr=hCritica(c,150);secImg('Gráfico — Perdas × sobras'+(c.hasCategorias?' por categoria':'')+' (com saldo)',chartCritica(c,1000,Math.round(hCr*5.56)),hCr);
    if(c.hasCategorias){sec('Resultado por categoria');aTT(['Categoria','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),['TOTAL',PCT(c.acuracidade),BRLi(soma(c.categorias,function(x){return x.faltaVal;})),BRLi(soma(c.categorias,function(x){return x.sobraVal;})),BRLi(soma(c.categorias,function(x){return x.saldo;}))],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});}
    /* r152: sem listagem de produtos no PDF (o Excel traz todos os itens); fica só o resumo por categoria acima */
  }
  else if(rt==='ruptura'){
    var r=data.ruptura;
    ttl('Ruptura Loja x Depósito — Resumo executivo');
    var iaRup=window._iaResumos&&window._iaResumos.ruptura;
    sec('Análise');bloco(iaRup||Engine.gerarAnaliseRuptura(r,info));
    sec('Metodologia');bloco(metRuptura(info.diasVenda));
    sec('Indicadores gerais');kpi(['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CURVA A (FAT.)','RUPTURA CURVA A (LUCRO)'],[PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],[[211,47,47],[51,51,51],[211,47,47],[211,47,47]]);
    /* r149: valor do estoque por local */
    if(r.valorDeposito!==undefined)kpi(['VALOR DO ESTOQUE (DEPÓSITO)','VALOR DO ESTOQUE (ÁREA DE VENDAS)','VALOR DO ESTOQUE EM RUPTURA'],[BRLi(r.valorDeposito),BRLi(r.valorLoja),BRLi(r.valorRuptura)+' ('+PCT(r.pctValorRuptura)+')'],[[0,21,40],[46,125,50],[211,47,47]]);
    secImg('Gráfico — Rupturas por curva ABC',chartRuptura(r,1000,380),68);
    if(r.valorDeposito!==undefined){var hEs=hRupturaEstoque(r,150);secImg('Gráfico — Valor do estoque: depósito (retaguarda) × área de vendas',chartRupturaEstoque(r,1000,Math.round(hEs*5.56)),hEs);}
    /* r152: sem listagem de produtos no PDF (o Excel traz todos os itens) */
  }
  else if(rt==='dias'){
    var d=data.dias,fv=fxV(d.items);ttl('Dias de estoque — Resumo executivo');
    var iaDias=window._iaResumos&&window._iaResumos.dias_estoque;
    sec('Análise');bloco(iaDias||Engine.gerarAnaliseDias(d,info));sec('Metodologia');bloco(metDias(info.diasVenda));
    sec('Indicadores gerais');kpi(['COBERTURA GERAL','CURVA A','CURVA B','CURVA C'],[d.coberturaGeral+' dias',d.coberturaA+' dias',d.coberturaB+' dias',d.coberturaC+' dias'],[[51,51,51],[211,47,47],[245,124,0],[136,136,136]]);
    secImg('Gráfico — Valor do estoque por potencial de ruptura',chartDias(d,1000,380),68);
    var fo=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];
    /* r149: sem colunas de SKUs, com linha de total e ordenado do maior para o menor valor de estoque */
    var faixasOrd=fo.map(function(f){return{f:f,v:fv.f[f]||0};}).sort(function(a,b){return b.v-a.v;});
    sec('Distribuição por faixa');
    aTT(['Faixa','Valor Estoque (R$)','% do Valor'],faixasOrd.map(function(x){return[x.f,BRLi(x.v),PCT(fv.t?x.v/fv.t*100:0)];}),['TOTAL',BRLi(fv.t),PCT(fv.t?100:0)],{1:{halign:'right'},2:{halign:'right'}});
    if(d.hasCategorias){
      var catsOrd=d.categorias.slice().sort(function(a,b){return (b.valorEstoque||0)-(a.valorEstoque||0);});
      var totCat={cr:0,sg:0,ex:0,v:0};catsOrd.forEach(function(x){totCat.cr+=x.criticos;totCat.sg+=x.semGiro;totCat.ex+=x.excessos;totCat.v+=x.valorEstoque;});
      sec('Cobertura por categoria');
      aTT(['Categoria','Cobertura média','Rupt+Alto risco','Sem giro','Excessos','Val. estoque'],catsOrd.map(function(x){return[x.nome,x.mediaCobertura+' dias',x.criticos,x.semGiro,x.excessos,BRLi(x.valorEstoque)];}),['TOTAL',d.coberturaGeral+' dias',totCat.cr,totCat.sg,totCat.ex,BRLi(totCat.v)],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});
    }
    /* r151: sem listagem de produtos no PDF (o Excel traz todos os itens) */
  }
  else if(rt==='abc'){
    var a=data.abc;ttl('Investimento por curva ABC — Resumo executivo');
    var iaABC=window._iaResumos&&window._iaResumos.abc;
    sec('Análise');bloco(iaABC||Engine.gerarAnaliseABC(a,info));sec('Metodologia');bloco(metABC(info.diasVenda));
    /* r149: cards com o valor do estoque por curva e Sem giro (C sem os itens sem giro) */
    var vcA=valoresCurvasABC(a);
    sec('Indicadores gerais');kpi(['VALOR TOTAL EM ESTOQUE','CURVA A ('+PCT(vcA.A.p)+')','CURVA B ('+PCT(vcA.B.p)+')'],[BRLi(a.totalInvest),BRLi(vcA.A.v),BRLi(vcA.B.v)],[[51,51,51],[0,21,40],[46,125,50]]);
    kpi(['CURVA C ('+PCT(vcA.C.p)+')','SEM GIRO ('+PCT(vcA.SG.p)+')'],[BRLi(vcA.C.v),BRLi(vcA.SG.v)],[[245,124,0],[136,136,136]]);
    secImg('Gráfico — Pareto do valor do estoque por curva',chartABC(a,1000,440),79);
    var sgA=a.semGiro||{invest:0,pctInvest:0},cgF=a.fatCg||a.fatC,cgL=a.lucCg||a.lucC;
    sec('Curva ABC por faturamento');aTT(['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(cgF.invest),PCT(cgF.pctInvest),BRLi(cgF.fat),PCT(cgF.pctFat)],['Sem giro',BRLi(sgA.invest),PCT(sgA.pctInvest),BRLi(0),PCT(0)]],['TOTAL',BRLi(a.totalInvest),PCT(a.totalInvest?100:0),BRLi(a.totalFat),PCT(a.totalFat?100:0)],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
    sec('Curva ABC por lucro');aTT(['Curva','Valor Estoque (R$)','% Estoque','Lucro (R$)','% Lucro'],[['A',BRLi(a.lucA.invest),PCT(a.lucA.pctInvest),BRLi(a.lucA.luc),PCT(a.lucA.pctLuc)],['B',BRLi(a.lucB.invest),PCT(a.lucB.pctInvest),BRLi(a.lucB.luc),PCT(a.lucB.pctLuc)],['C',BRLi(cgL.invest),PCT(cgL.pctInvest),BRLi(cgL.luc),PCT(cgL.pctLuc)],['Sem giro',BRLi(sgA.invest),PCT(sgA.pctInvest),BRLi(0),PCT(0)]],['TOTAL',BRLi(a.totalInvest),PCT(a.totalInvest?100:0),BRLi(a.totalLucro),PCT(a.totalLucro?100:0)],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
  }
  else if(rt==='perda'){
    var p=data.perda;ttl('Projeção de venda perdida — Resumo executivo');
    var iaPerda=window._iaResumos&&window._iaResumos.perda;
    sec('Análise');bloco(iaPerda||Engine.gerarAnalisePerda(p,info));sec('Metodologia');bloco(metPerda(info.diasVenda));
    sec('Indicadores gerais');kpi(['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS EM RUPTURA'],[BRLi(p.totalPerdaFat),BRLi(p.totalPerdaLucro),BRLi(p.perdaMensal),NUM(p.totalSKUs)],[[211,47,47],[211,47,47],[211,47,47],[51,51,51]]);
    sec('Projeção de Perda');aTT(['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',p.classA.count,BRLi(p.classA.perda),BRLi(p.classA.lucro),PCT(p.classA.pct),BRLi(p.classA.perda*30)],['B',p.classB.count,BRLi(p.classB.perda),BRLi(p.classB.lucro),PCT(p.classB.pct),BRLi(p.classB.perda*30)],['C',p.classC.count,BRLi(p.classC.perda),BRLi(p.classC.lucro),PCT(p.classC.pct),BRLi(p.classC.perda*30)]],['TOTAL',NUM(p.totalSKUs),BRLi(p.totalPerdaFat),BRLi(p.totalPerdaLucro),PCT(p.totalPerdaFat?100:0),BRLi(p.perdaMensal)],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});
    secImg('Gráfico — Valor do estoque e perda projetada por curva',chartPerda(p,data.critica,1000,380),68);
    /* r150: a listagem dos produtos em ruptura fica só no Excel */
  }
  chk(12);doc.setFontSize(7);doc.setTextColor(150,150,150);
  doc.text('Nota: relatório baseado em dados processados em '+pd+'. Valores projetados são estimativas.',M,y);
  doc.save('resumo_'+rt+'_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.pdf');
}
/* ========== PDF COMPARATIVO ========== */
function generateComparativoPDF(comp, units, info, iaTextos, logo){
  iaTextos=iaTextos||{};info=info||{};
  var jsPDF=window.jspdf.jsPDF;var doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  var W=210,H=297,M=15,y=0;
  var pd=new Date().toLocaleString('pt-BR');

  function hdr(){doc.setFillColor(0,21,40);doc.rect(0,0,W,22,'F');if(logo){try{doc.addImage(logo,'PNG',M,7,32,8);}catch(e){}}doc.setFontSize(9);doc.setTextColor(255,255,255);doc.text('COMPARATIVO — '+(info.cliente||''),W-M,7,{align:'right'});doc.setFontSize(7);doc.setTextColor(200,220,255);doc.text(comp.unidades.map(function(u){return u.unidade;}).join(' × '),W-M,12,{align:'right'});doc.setTextColor(180,180,200);doc.text('Inventário: '+(info.dataInventario||'—')+' | Gerado em '+pd,W-M,17,{align:'right'});y=28;}
  function ftr(pg){doc.setFontSize(7);doc.setTextColor(150,150,150);doc.text('Formula Code Tecnologia, Gestão e Automação',M,H-6);doc.text('Página '+pg,W-M,H-6,{align:'right'});doc.setDrawColor(200,200,200);doc.line(M,H-10,W-M,H-10);}
  function chk(n){if(y+n>H-18){doc.addPage();hdr();ftr(doc.getNumberOfPages());}}
  function ttl(t){chk(12);doc.setFontSize(14);doc.setTextColor(0,21,40);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFontSize(8);doc.setTextColor(150,150,150);doc.setFont(undefined,'normal');doc.text('Relatório comparativo gerado pelo Sistema Formula Code',M,y);y+=8;}
  function sec(t){chk(10);doc.setFontSize(11);doc.setTextColor(0,21,40);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFont(undefined,'normal');}
  function bloco(txt){if(!txt)return;chk(16);doc.setFontSize(8);doc.setTextColor(80,80,80);doc.setFont(undefined,'normal');var lines=doc.splitTextToSize(txt,W-2*M);doc.text(lines,M,y);y+=lines.length*3.5+4;}
  function aT(h,b,o){chk(20);doc.autoTable({startY:y,head:[h],body:b,margin:{left:M,right:M},headStyles:{fillColor:[0,21,40],fontSize:7,fontStyle:'bold',halign:'left'},bodyStyles:{fontSize:7,halign:'left'},alternateRowStyles:{fillColor:[245,245,245]},styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},columnStyles:o||{}});y=doc.lastAutoTable.finalY+6;}
  function kpi(lb,vl,cl){chk(18);var cw=(W-2*M)/lb.length;doc.setFillColor(245,245,245);doc.roundedRect(M,y-2,W-2*M,16,2,2,'F');for(var i=0;i<lb.length;i++){var x=M+i*cw+4;doc.setFontSize(7);doc.setTextColor(150,150,150);doc.setFont(undefined,'bold');doc.text(lb[i],x,y+3);doc.setFontSize(11);doc.setFont(undefined,'bold');var cc=cl[i]||[51,51,51];doc.setTextColor(cc[0],cc[1],cc[2]);doc.text(String(vl[i]),x,y+10);}doc.setFont(undefined,'normal');y+=20;}

  hdr();ftr(1);
  ttl('Comparativo entre unidades — '+comp.unidades.length+' unidades');

  /* KPIs globais */
  if(comp.skuOverlap){
    sec('Sobreposição de SKUs');
    kpi(['SKUS ÚNICOS (TOTAL)','EM COMUM','SOBREPOSIÇÃO'],
      [NUM(comp.skuOverlap.totalUnique),NUM(comp.skuOverlap.emComum),PCT(comp.skuOverlap.pctComum)],
      [[51,51,51],[0,183,74],[0,183,74]]);
  }

  /* IA resumo comparativo */
  sec('Análise comparativa');
  bloco(iaTextos.resumo_comparativo||'Análise comparativa entre '+comp.unidades.length+' unidades do cliente '+(info.cliente||'')+' referente ao inventário de '+(info.dataInventario||'')+'.');

  /* Tabela de rankings */
  sec('Ranking por métrica');
  var tH=['Métrica'];comp.unidades.forEach(function(u){tH.push(u.unidade);});
  var tB=[];
  comp.rankings.forEach(function(r){
    var row=[r.label];
    r.valores.forEach(function(v){
      row.push(fmtRank(r,v)+(r.melhor&&v.unidade===r.melhor?' ★':'')+(r.pior&&v.unidade===r.pior?' ▼':''));
    });
    tB.push(row);
  });
  var colStyles={0:{fontStyle:'bold'}};
  for(var ci=1;ci<=comp.unidades.length;ci++)colStyles[ci]={halign:'right'};
  aT(tH,tB,colStyles);

  /* IA por dimensão */
  if(iaTextos.analise_critica){sec('Análise — Crítica do inventário');bloco(iaTextos.analise_critica);}
  if(iaTextos.analise_ruptura){sec('Análise — Ruptura Loja x Depósito');bloco(iaTextos.analise_ruptura);}
  if(iaTextos.analise_cobertura){sec('Análise — Cobertura de estoque');bloco(iaTextos.analise_cobertura);}
  if(iaTextos.analise_perda){sec('Análise — Projeção de perda');bloco(iaTextos.analise_perda);}

  /* Recomendações */
  if(iaTextos.recomendacoes){
    sec('Recomendações');
    bloco(iaTextos.recomendacoes);
  }

  /* Legenda */
  chk(12);doc.setFontSize(7);doc.setTextColor(150,150,150);
  doc.text('★ = melhor desempenho na métrica · ▼ = pior. Relatório gerado em '+pd+'.',M,y);
  doc.save('comparativo_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.pdf');
}

/* ========== EXCEL COMPARATIVO ========== */
function generateComparativoExcel(comp, units, info){
  info=info||{};var pd=new Date().toLocaleString('pt-BR');
  var wb=XLSX.utils.book_new();
  var ws={},R=0;
  var nUnits=comp.unidades.length;
  var TC=nUnits+2;
  R=addBH(ws,R,info,pd,TC);
  R=addST(ws,R,'COMPARATIVO — '+nUnits+' UNIDADES');

  /* r149: cards de SKUs/Sobreposição removidos do comparativo (tela e exportações) */

  /* Rankings */
  R=addST(ws,R,'RANKING POR MÉTRICA');
  var hdr=['Métrica'];comp.unidades.forEach(function(u){hdr.push(u.unidade);});
  var rows=comp.rankings.map(function(r){
    var row=[r.label];
    r.valores.forEach(function(v){
      row.push(fmtRank(r,v));
    });
    return row;
  });
  var ca={0:'left'};for(var i=1;i<=nUnits;i++)ca[i]='right';
  R=addDT(ws,R,hdr,rows,ca);

  var colWidths=[{wch:36}];for(var j=0;j<nUnits;j++)colWidths.push({wch:24});
  ws['!cols']=colWidths;
  ws['!rows']=[{hpt:28},{hpt:20}];
  XLSX.utils.book_append_sheet(wb,ws,'Comparativo');

  var out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  var blob=new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  var url=URL.createObjectURL(blob);var a=document.createElement('a');a.href=url;
  a.download='comparativo_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.xlsx';
  a.click();URL.revokeObjectURL(url);
}

/* (r149: o return do módulo foi movido para o final do arquivo, depois das variáveis dos slides) */

/* ===== r68: RESUMO EXECUTIVO — PDF Documento ===== */
function generateResumoPDF(results,recs,info,unidade,logo){
  var jsPDF=window.jspdf.jsPDF;
  var doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  var W=210,H=297,M=20,cw=W-2*M;
  var y=M;
  var navy=[0,21,40],green=[93,197,0],dark=[0,21,40];

  /* Cover */
  doc.setFillColor.apply(doc,dark);doc.rect(0,0,W,H,'F');
  doc.setFillColor.apply(doc,green);doc.rect(0,0,W,2,'F');
  if(logo){try{doc.addImage(logo,'PNG',60,30,90,22.46);}catch(e){}}
  doc.setFontSize(28);doc.setTextColor(255,255,255);
  doc.text('Resumo Executivo',W/2,80,{align:'center'});
  doc.setFontSize(14);doc.setTextColor.apply(doc,green);
  doc.text('Análise de Inventário',W/2,92,{align:'center'});
  doc.setFontSize(12);doc.setTextColor(136,153,170);
  var sub=info.cliente||'';
  if(unidade)sub+=' — '+unidade;
  sub+='  ·  '+(info.dataInventario||'');
  doc.text(sub,W/2,106,{align:'center'});
  doc.addPage();

  /* Content page */
  y=M;
  doc.setFillColor.apply(doc,navy);doc.rect(0,0,W,12,'F');
  doc.setFontSize(10);doc.setTextColor(255,255,255);
  doc.text('RESUMO EXECUTIVO — '+(info.cliente||'').toUpperCase(),M,8);
  y=20;

  doc.setFontSize(18);doc.setTextColor.apply(doc,navy);
  doc.text('Indicadores-Chave',M,y);y+=10;

  var red=[211,47,47],amb=[245,124,0];
  var dims=[];
  if(results.critica)dims.push({label:'Acuracidade',valor:results.critica.acuracidade+'%',detalhe:results.critica.totalSKUs+' SKUs · '+results.critica.faltaCount+' faltas · '+results.critica.sobraCount+' sobras',cor:results.critica.acuracidade<90?red:(results.critica.acuracidade<95?amb:green)});
  if(results.ruptura)dims.push({label:'Ruptura',valor:results.ruptura.taxaRuptura+'%',detalhe:results.ruptura.totalRupturas+' SKUs em falta · Curva A: '+results.ruptura.rupturaA,cor:results.ruptura.taxaRuptura>10?red:(results.ruptura.taxaRuptura>5?amb:green)});
  if(results.dias)dims.push({label:'Cobertura',valor:Engine.round2(results.dias.coberturaGeral)+' dias',detalhe:'Sem giro: '+results.dias.semGiro+' · Excesso: '+results.dias.excessos,cor:(results.dias.coberturaGeral<15||results.dias.coberturaGeral>60)?red:green});
  if(results.abc)dims.push({label:'Investimento',valor:'R$ '+Engine.formatNum(results.abc.totalInvest),detalhe:'Faturamento: R$ '+Engine.formatNum(results.abc.totalFat),cor:navy});
  if(results.perda)dims.push({label:'Perda Mensal',valor:'R$ '+Engine.formatNum(results.perda.perdaMensal),detalhe:results.perda.totalSKUs+' SKUs identificados',cor:results.perda.perdaMensal>50000?red:(results.perda.perdaMensal>10000?amb:green)});

  /* Cards em grade de 2 colunas */
  var cardGap=6,cardW=(cw-cardGap)/2,cardH=32;
  dims.forEach(function(d,i){
    var col=i%2,row=Math.floor(i/2);
    var cx=M+col*(cardW+cardGap),cy=y+row*(cardH+cardGap);
    doc.setFillColor(245,245,245);doc.roundedRect(cx,cy,cardW,cardH,2,2,'F');
    doc.setFillColor.apply(doc,d.cor);doc.rect(cx,cy,1.3,cardH,'F');
    doc.setFontSize(8);doc.setTextColor.apply(doc,green);
    doc.text(d.label.toUpperCase(),cx+7,cy+8);
    doc.setFontSize(16);doc.setTextColor.apply(doc,navy);
    doc.text(d.valor,cx+7,cy+17);
    doc.setFontSize(8);doc.setTextColor(85,102,119);
    doc.text(doc.splitTextToSize(d.detalhe,cardW-12),cx+7,cy+24);
  });
  y+=Math.ceil(dims.length/2)*(cardH+cardGap)+6;

  /* r149: um gráfico por página, com os valores */
  var grR=graficosResumo(results);
  grR.forEach(function(g){
    doc.addPage();
    doc.setFillColor.apply(doc,navy);doc.rect(0,0,W,12,'F');
    doc.setFontSize(10);doc.setTextColor(255,255,255);
    doc.text('RESUMO EXECUTIVO — GRÁFICOS',M,8);
    y=24;
    doc.setFontSize(13);doc.setTextColor.apply(doc,navy);
    var tl=doc.splitTextToSize(g.titulo,cw);doc.text(tl,M,y);y+=tl.length*6+4;
    var hImg=Math.min(150,297-M-y);
    var url=_chartPNG(g.cfg(13),Math.round(cw*5.56),Math.round(hImg*5.56));
    if(url){try{doc.addImage(url,'PNG',M,y,cw,hImg);}catch(e){}}
    y+=hImg+8;
  });
  if(recs.length>0&&grR.length){doc.addPage();y=M;}

  /* Recomendações */
  if(recs.length>0){
    if(y>240){doc.addPage();y=M;}
    doc.setFontSize(18);doc.setTextColor.apply(doc,navy);
    doc.text('Recomendações',M,y);y+=10;

    recs.forEach(function(rec){
      if(y>265){doc.addPage();y=M;}
      var icon=rec.prioridade===1?'[!]':(rec.prioridade===2?'[▲]':'[✓]');
      var color=rec.prioridade===1?[204,51,51]:(rec.prioridade===2?[232,135,43]:[34,139,34]);
      doc.setFontSize(10);doc.setTextColor.apply(doc,color);
      doc.text(icon,M,y);
      doc.setTextColor(68,85,102);
      var lines=doc.splitTextToSize(rec.texto,cw-10);
      doc.text(lines,M+10,y);
      y+=lines.length*5+6;
    });
  }

  doc.save('resumo_executivo_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'.pdf');
}

/* ===================================================================================
   r149: APRESENTAÇÕES (PPTX + PDF) a partir de um único roteiro de slides.
   Cada slide é descrito uma vez (retângulos, textos, imagens, tabelas) e desenhado
   igual no PowerPoint (PptxGenJS) e no PDF (jsPDF, 16:9). Área útil respeita a
   margem mínima de 0,5 pol em todos os lados (só fundos e faixas decorativas vão
   até a borda).
   =================================================================================== */
var SL={W:10,H:5.625,M:0.5};
SL.R=SL.W-SL.M;SL.B=SL.H-SL.M;
function _hexRGB(h){h=String(h||'000000').replace('#','');return[parseInt(h.substr(0,2),16),parseInt(h.substr(2,2),16),parseInt(h.substr(4,2),16)];}
function novoDeck(){
  var deck={slides:[]};
  deck.add=function(bg){
    var sl={bg:bg||'FFFFFF',items:[]};deck.slides.push(sl);
    var api={
      rect:function(x,y,w,h,cor){sl.items.push({t:'rect',x:x,y:y,w:w,h:h,cor:cor});return api;},
      ellipse:function(x,y,w,h,cor){sl.items.push({t:'ellipse',x:x,y:y,w:w,h:h,cor:cor});return api;},
      text:function(txt,x,y,w,h,o){sl.items.push({t:'text',txt:txt,x:x,y:y,w:w,h:h,o:o||{}});return api;},
      img:function(data,x,y,w,h){if(data)sl.items.push({t:'img',data:data,x:x,y:y,w:w,h:h});return api;},
      /* cabecalho: array de textos; corpo: array de linhas; célula = texto ou {t,cor,bold,fill} */
      table:function(cab,corpo,x,y,w,o){sl.items.push({t:'table',cab:cab,corpo:corpo,x:x,y:y,w:w,o:o||{}});return api;}
    };
    return api;
  };
  return deck;
}
/* medidor de texto (mesmas métricas Helvetica/Arial usadas no PDF) */
var _medidor=null;
function linhasTexto(txt,w,fs,bold){
  if(!_medidor){_medidor=new window.jspdf.jsPDF({orientation:'landscape',unit:'in',format:[SL.H,SL.W]});}
  _medidor.setFont('helvetica',bold?'bold':'normal');_medidor.setFontSize(fs);
  return _medidor.splitTextToSize(String(txt||''),w);
}
function _celTxt(c){return (c&&typeof c==='object')?String(c.t===undefined?'':c.t):String(c===undefined||c===null?'':c);}

function deckParaPPTX(deck){
  var pres=new PptxGenJS();pres.layout='LAYOUT_16x9';
  deck.slides.forEach(function(sl){
    var s=pres.addSlide();s.background={color:sl.bg};
    sl.items.forEach(function(it){
      if(it.t==='rect'){s.addShape(pres.ShapeType.rect,{x:it.x,y:it.y,w:it.w,h:it.h,fill:{color:it.cor},line:{color:it.cor,width:0}});}
      else if(it.t==='ellipse'){s.addShape(pres.ShapeType.ellipse,{x:it.x,y:it.y,w:it.w,h:it.h,fill:{color:it.cor},line:{color:it.cor,width:0}});}
      else if(it.t==='img'){s.addImage({data:it.data,x:it.x,y:it.y,w:it.w,h:it.h});}
      else if(it.t==='text'){
        var o=it.o;
        var conteudo=Array.isArray(it.txt)?it.txt.map(function(r){return{text:r.text,options:{color:r.color||o.color||'333333',bold:r.bold!==undefined?r.bold:!!o.bold}};}):String(it.txt);
        var op={x:it.x,y:it.y,w:it.w,h:it.h,fontSize:o.fs||12,fontFace:'Arial',color:o.color||'333333',bold:!!o.bold,align:o.align||'left',valign:o.valign||'top',margin:0,isTextBox:true,fit:'none',lineSpacingMultiple:o.ls||1.0};
        if(o.cs)op.charSpacing=o.cs;
        s.addText(conteudo,op);
      }
      else if(it.t==='table'){
        var o2=it.o,fs=o2.fs||9,cols=it.cab.length;
        var colW=o2.colW||it.cab.map(function(){return it.w/cols;});
        var al=o2.align||[];
        var linhas=[it.cab.map(function(c,ci){return{text:_celTxt(c),options:{bold:true,color:'FFFFFF',fill:{color:o2.headFill||'001528'},align:al[ci]||'left'}};})];
        it.corpo.forEach(function(row,ri){
          linhas.push(row.map(function(c,ci){var ob=(c&&typeof c==='object')?c:{};return{text:_celTxt(c),options:{bold:!!ob.bold,color:ob.cor||'333333',fill:{color:ob.fill||(ri%2?'F5F5F5':'FFFFFF')},align:al[ci]||'left'}};}));
        });
        s.addTable(linhas,{x:it.x,y:it.y,w:it.w,colW:colW,rowH:o2.rowH||0.3,fontSize:fs,fontFace:'Arial',valign:'middle',margin:[0.02,0.06,0.02,0.06],border:{type:'solid',pt:0.5,color:'DDDDDD'},autoPage:false});
      }
    });
  });
  return pres;
}
function deckParaPDF(deck){
  var jsPDF=window.jspdf.jsPDF;
  var doc=new jsPDF({orientation:'landscape',unit:'in',format:[SL.H,SL.W]});
  deck.slides.forEach(function(sl,si){
    if(si>0)doc.addPage([SL.H,SL.W],'landscape');
    var b=_hexRGB(sl.bg);doc.setFillColor(b[0],b[1],b[2]);doc.rect(0,0,SL.W,SL.H,'F');
    sl.items.forEach(function(it){
      if(it.t==='rect'){var c=_hexRGB(it.cor);doc.setFillColor(c[0],c[1],c[2]);doc.rect(it.x,it.y,it.w,it.h,'F');}
      else if(it.t==='ellipse'){var c2=_hexRGB(it.cor);doc.setFillColor(c2[0],c2[1],c2[2]);doc.ellipse(it.x+it.w/2,it.y+it.h/2,it.w/2,it.h/2,'F');}
      else if(it.t==='img'){try{doc.addImage(it.data,'PNG',it.x,it.y,it.w,it.h);}catch(e){}}
      else if(it.t==='text'){
        var o=it.o,fs=o.fs||12,lh=fs/72*1.2*(o.ls||1.0);
        if(o.cs)doc.setCharSpace(o.cs/72);else doc.setCharSpace(0);
        if(Array.isArray(it.txt)){
          /* trechos coloridos numa linha só */
          var larg=0;it.txt.forEach(function(r){doc.setFont('helvetica',(r.bold!==undefined?r.bold:o.bold)?'bold':'normal');doc.setFontSize(fs);larg+=doc.getTextWidth(r.text);});
          var x0=o.align==='center'?it.x+(it.w-larg)/2:(o.align==='right'?it.x+it.w-larg:it.x);
          var yT=o.valign==='middle'?it.y+(it.h-lh)/2:(o.valign==='bottom'?it.y+it.h-lh:it.y);
          it.txt.forEach(function(r){doc.setFont('helvetica',(r.bold!==undefined?r.bold:o.bold)?'bold':'normal');doc.setFontSize(fs);var cc=_hexRGB(r.color||o.color||'333333');doc.setTextColor(cc[0],cc[1],cc[2]);doc.text(r.text,x0,yT+(lh-fs/72)/2,{baseline:'top'});x0+=doc.getTextWidth(r.text);});
        }else{
          doc.setFont('helvetica',o.bold?'bold':'normal');doc.setFontSize(fs);
          var ct=_hexRGB(o.color||'333333');doc.setTextColor(ct[0],ct[1],ct[2]);
          var ls=[];String(it.txt).split('\n').forEach(function(p){ls=ls.concat(doc.splitTextToSize(p,it.w));});
          var tot=ls.length*lh;
          var y0=o.valign==='middle'?it.y+(it.h-tot)/2:(o.valign==='bottom'?it.y+it.h-tot:it.y);
          var xa=o.align==='center'?it.x+it.w/2:(o.align==='right'?it.x+it.w:it.x);
          ls.forEach(function(l,li){doc.text(l,xa,y0+li*lh+(lh-fs/72)/2,{baseline:'top',align:o.align||'left'});});
        }
        doc.setCharSpace(0);
      }
      else if(it.t==='table'){
        var o3=it.o,cols=it.cab.length,colW=o3.colW||it.cab.map(function(){return it.w/cols;}),al=o3.align||[];
        var cs={};colW.forEach(function(w,ci){cs[ci]={cellWidth:w,halign:al[ci]||'left'};});
        var hf=_hexRGB(o3.headFill||'001528');
        doc.autoTable({startY:it.y,margin:{left:it.x,right:SL.W-it.x-it.w,top:SL.M,bottom:SL.M},tableWidth:it.w,pageBreak:'avoid',rowPageBreak:'avoid',
          head:[it.cab.map(_celTxt)],body:it.corpo.map(function(r){return r.map(_celTxt);}),
          styles:{font:'helvetica',fontSize:o3.fs||9,cellPadding:{top:0.02,bottom:0.02,left:0.06,right:0.06},minCellHeight:o3.rowH||0.3,valign:'middle',lineColor:[221,221,221],lineWidth:0.007,textColor:[51,51,51],overflow:'linebreak'},
          headStyles:{fillColor:hf,textColor:[255,255,255],fontStyle:'bold'},
          alternateRowStyles:{fillColor:[245,245,245]},columnStyles:cs,
          didParseCell:function(dd){
            dd.cell.styles.halign=al[dd.column.index]||'left';
            if(dd.section!=='body')return;
            var c=it.corpo[dd.row.index]&&it.corpo[dd.row.index][dd.column.index];
            if(c&&typeof c==='object'){if(c.cor)dd.cell.styles.textColor=_hexRGB(c.cor);if(c.bold)dd.cell.styles.fontStyle='bold';if(c.fill)dd.cell.styles.fillColor=_hexRGB(c.fill);}
          }});
      }
    });
  });
  return doc;
}
/* Baixa PPTX e PDF juntos (mesmo conteúdo) */
function salvarDeck(deck,nomeBase){
  var pres=deckParaPPTX(deck);
  var p=pres.writeFile({fileName:nomeBase+'.pptx'});
  var fazPDF=function(){try{deckParaPDF(deck).save(nomeBase+'.pdf');}catch(e){console.log('Erro ao gerar PDF da apresentação:',e);alert('Não foi possível gerar o PDF da apresentação.');}};
  if(p&&typeof p.then==='function')p.then(function(){setTimeout(fazPDF,400);},function(){setTimeout(fazPDF,400);});else setTimeout(fazPDF,400);
}
function _lim(s){return String(s||'').replace(/[^a-zA-Z0-9]/g,'_');}

/* ---- Blocos visuais comuns ---- */
var PAL={navy:'001528',green:'5DC500',label:'8899AA',body:'556677',red:'D32F2F',amb:'F57C00',greenDk:'2E7D32'};
function slideCapa(deck,titulo1,titulo2,subtitulo,linha3,logo){
  var s=deck.add(PAL.navy);
  s.rect(0,0,SL.W,0.05,PAL.green);
  s.text([{text:titulo1+' ',color:'FFFFFF'},{text:titulo2,color:PAL.green}],0.8,0.9,8.4,0.7,{fs:34,bold:true});
  s.text(subtitulo,0.8,1.75,8.4,0.4,{fs:16,color:PAL.green});
  if(linha3)s.text(linha3,0.8,2.3,8.4,0.8,{fs:13,color:PAL.label});
  if(logo)s.img(logo,3.2,4.1,3.6,0.9);
  return s;
}
/* Slide com faixa navy no topo (título + logo) e área útil branca abaixo */
function slideFaixa(deck,titulo,subtitulo,logo,compacto){
  if(compacto)return slideFaixaCompacta(deck,titulo,subtitulo,logo);
  var s=deck.add('FFFFFF');
  s.rect(0,0,SL.W,1.0,PAL.navy);
  s.rect(0,1.0,SL.W,0.04,PAL.green);
  var fsT=18;while(fsT>11&&linhasTexto(titulo,6.9,fsT,true).length>1)fsT--;
  s.text(titulo,SL.M,SL.M,6.9,0.32,{fs:fsT,bold:true,color:'FFFFFF',valign:'middle'});
  if(subtitulo)s.text(subtitulo,SL.M,0.8,6.9,0.18,{fs:9,color:'B0C4DE',valign:'middle'});
  if(logo)s.img(logo,7.7,0.5,1.8,0.45);
  return s;
}
/* r151: faixa de topo menor (Comparativo) — conteúdo começa em 0,92 pol */
var TOPO_C=0.98;
function slideFaixaCompacta(deck,titulo,subtitulo,logo){
  var s=deck.add('FFFFFF');
  s.rect(0,0,SL.W,0.86,PAL.navy);
  s.rect(0,0.86,SL.W,0.03,PAL.green);
  var fsT=13;while(fsT>10&&linhasTexto(titulo,7.2,fsT,true).length>1)fsT--;
  s.text(titulo,SL.M,SL.M,7.2,0.2,{fs:fsT,bold:true,color:'FFFFFF',valign:'top'});
  if(subtitulo)s.text(subtitulo,SL.M,0.72,7.2,0.12,{fs:7,color:'B0C4DE',valign:'top'});
  if(logo)s.img(logo,8.3,0.5,1.2,0.3);
  return s;
}
/* Slide dividido (conteúdo à esquerda, painel navy à direita) — padrão do Resumo */
function slideDividido(deck,logo,cliente,unidade){
  var s=deck.add('FFFFFF');
  s.rect(6.8,0,3.2,SL.H,PAL.navy);
  if(logo)s.img(logo,7.1,4.52,2.4,0.6);
  s.text((cliente||'')+(unidade?'\n'+unidade:''),7.1,0.8,2.4,0.6,{fs:12,bold:true,color:'FFFFFF'});
  return s;
}
/* Linha de cartões (um por unidade) — até 6 por linha */
function cartoesUnidades(s,itens,y0,hMax){
  var n=itens.length;if(!n)return y0;
  var porLinha=Math.min(n,6),linhas=Math.ceil(n/porLinha),gap=0.12;
  var w=(SL.R-SL.M-gap*(porLinha-1))/porLinha,h=Math.min(0.9,(hMax-gap*(linhas-1))/linhas);
  itens.forEach(function(c,i){
    var col=i%porLinha,lin=Math.floor(i/porLinha),x=SL.M+col*(w+gap),y=y0+lin*(h+gap);
    s.rect(x,y,w,h,'F5F5F5');s.rect(x,y,0.05,h,c.cor||PAL.navy);
    s.text(_cortar(c.titulo,30),x+0.12,y+0.06,w-0.18,0.18,{fs:8,bold:true,color:PAL.body});
    s.text(c.valor,x+0.12,y+0.25,w-0.18,h*0.36,{fs:h<0.8?12:15,bold:true,color:PAL.navy,valign:'middle'});
    if(c.sub)s.text(c.sub,x+0.12,y+h-0.24,w-0.18,0.18,{fs:8,color:PAL.body});
  });
  return y0+linhas*h+(linhas-1)*gap;
}
function _px(wIn,hIn){return{w:Math.round(wIn*100),h:Math.round(hIn*100)};}
function imgGrafico(s,cfg,x,y,w,h){var px=_px(w,h);s.img(_chartPNG(cfg,px.w,px.h),x,y,w,h);}

/* ---- Formatação de valor do ranking do comparativo (tela, Excel e apresentação) ---- */
function fmtRank(r,v){
  if(v.valor===null||v.valor===undefined)return '—';
  if(r.fmt==='pct')return PCT(v.valor);
  if(r.fmt==='brl')return BRLi(v.valor);
  if(r.fmt==='brl_pct')return BRLi(v.valor)+' ('+PCT(v.pct||0)+')';
  if(r.fmt==='num')return NUM(v.valor);
  return String(v.valor);
}

/* ---- Lista de gráficos (um por página) usada no Resumo ---- */
function graficosResumo(results){
  var g=[];
  if(results.critica){var c=results.critica;g.push({titulo:'Crítica — Perdas × sobras'+(c.hasCategorias?' por categoria':'')+' (com saldo)',cfg:function(fs){return cfgCritica(c,fs);}});}
  if(results.ruptura){var r=results.ruptura;g.push({titulo:'Ruptura Loja x Depósito — SKUs e valor do estoque por curva ABC',cfg:function(fs){return cfgRuptura(r,fs);}});
    if(r.valorDeposito!==undefined)g.push({titulo:'Valor do estoque — depósito (retaguarda) × área de vendas',cfg:function(fs){return cfgRupturaEstoque(r,fs);}});}
  if(results.dias){var d=results.dias;g.push({titulo:'Dias de estoque — valor do estoque por potencial de ruptura',cfg:function(fs){return cfgDias(d,fs);}});}
  if(results.abc){var a=results.abc;g.push({titulo:'Investimento ABC — Pareto do valor do estoque por curva',cfg:function(fs){return cfgABCPareto(a,fs);}});}
  if(results.perda){var p=results.perda,cr=results.critica;g.push({titulo:'Projeção de perda — valor do estoque e perda projetada por curva',cfg:function(fs){return cfgPerda(p,cr,fs);}});}
  return g;
}

/* ===== r149: RESUMO EXECUTIVO — apresentação (PPTX + PDF com o mesmo conteúdo) ===== */
function montarDeckResumo(results,recs,info,unidade,logo){
  var deck=novoDeck();
  var sub=(info.cliente||'')+(unidade?' — '+unidade:'')+'  ·  '+(info.dataInventario||'');
  slideCapa(deck,'Resumo','Executivo','Análise de Inventário',sub,logo);

  /* Indicadores-chave */
  var s2=slideDividido(deck,logo,info.cliente,unidade);
  s2.text('01 / INDICADORES-CHAVE',0.8,SL.M,5.5,0.22,{fs:9,bold:true,color:PAL.green,cs:3});
  s2.text([{text:'Diagnóstico do ',color:PAL.navy},{text:'Inventário',color:PAL.green}],0.8,0.85,5.6,0.55,{fs:28,bold:true});
  s2.text((info.dataInventario||'')+'\n'+(info.diasVenda||90)+' dias de venda',7.1,1.5,2.4,0.5,{fs:10,color:PAL.label,ls:1.3});
  var dims=[];
  if(results.critica)dims.push({label:'ACURACIDADE',valor:PCT(results.critica.acuracidade),sub:'Meta: 95%',accent:results.critica.acuracidade<95});
  if(results.ruptura)dims.push({label:'RUPTURA',valor:PCT(results.ruptura.taxaRuptura),sub:NUM(results.ruptura.totalRupturas)+' SKUs em falta',accent:true});
  if(results.dias)dims.push({label:'COBERTURA',valor:Engine.round2(results.dias.coberturaGeral)+' dias',sub:'Sem giro: '+NUM(results.dias.semGiro)+' SKUs',accent:results.dias.coberturaGeral>45});
  if(results.abc)dims.push({label:'INVESTIMENTO',valor:'R$ '+Engine.formatNum(results.abc.totalInvest),sub:'Giro: '+(results.abc.totalInvest?Engine.round2(results.abc.totalFat/results.abc.totalInvest):0)+'x',accent:false});
  if(results.perda)dims.push({label:'PERDA MENSAL',valor:'R$ '+Engine.formatNum(results.perda.perdaMensal),sub:NUM(results.perda.totalSKUs)+' SKUs',accent:true});
  dims.forEach(function(d,i){
    var ky=1.65+i*0.68;
    s2.rect(0.8,ky,0.04,0.5,d.accent?PAL.green:PAL.navy);
    s2.text(d.valor,1.05,ky,2.0,0.32,{fs:20,bold:true,color:PAL.navy});
    s2.text(d.label,3.1,ky,2.8,0.18,{fs:8,bold:true,color:PAL.green,cs:2});
    s2.text(d.sub,3.1,ky+0.2,3.4,0.2,{fs:10,color:PAL.body});
  });

  /* Um gráfico por página, com os valores */
  var subG=(info.cliente||'')+(unidade?' — '+unidade:'')+'  ·  '+(info.dataInventario||'');
  graficosResumo(results).forEach(function(g,i){
    var sg=slideFaixa(deck,g.titulo,'02.'+(i+1)+' / GRÁFICOS  ·  '+subG,logo);
    imgGrafico(sg,g.cfg(12),SL.M,1.2,SL.R-SL.M,SL.B-1.2);
  });

  /* Recomendações — texto longo continua no slide seguinte */
  if(recs&&recs.length){
    var w=5.3,fs=10,lh=fs/72*1.2*1.25,y0=1.65,pag=null,y=0,n=0;
    function novaPag(){
      pag=slideDividido(deck,logo,info.cliente,unidade);n++;
      pag.text('03 / RECOMENDAÇÕES'+(n>1?' (continuação)':''),0.8,SL.M,5.5,0.22,{fs:9,bold:true,color:PAL.green,cs:3});
      pag.text([{text:'Ações ',color:PAL.navy},{text:'Recomendadas',color:PAL.green}],0.8,0.85,5.6,0.55,{fs:28,bold:true});
      y=y0;
    }
    novaPag();
    recs.forEach(function(rec){
      var cor=rec.prioridade===1?'CC3333':(rec.prioridade===2?'E8872B':'228B22');
      var ls=linhasTexto(rec.texto,w,fs,false);
      var h=Math.max(0.3,ls.length*lh);
      if(y+h>SL.B&&y>y0)novaPag();
      if(y+h>SL.B){ /* recomendação maior que a página: quebra em partes */
        var cabem=Math.max(1,Math.floor((SL.B-y)/lh));
        while(ls.length){
          var parte=ls.splice(0,cabem);
          pag.ellipse(0.8,y+0.02,0.25,0.25,cor);pag.text(String(rec.prioridade||''),0.8,y+0.02,0.25,0.25,{fs:10,bold:true,color:'FFFFFF',align:'center',valign:'middle'});
          pag.text(parte.join('\n'),1.15,y,w,parte.length*lh,{fs:fs,color:PAL.body,ls:1.25});
          if(ls.length){novaPag();cabem=Math.max(1,Math.floor((SL.B-y)/lh));}else{y+=parte.length*lh+0.14;}
        }
        return;
      }
      pag.ellipse(0.8,y+0.02,0.25,0.25,cor);
      pag.text(String(rec.prioridade||''),0.8,y+0.02,0.25,0.25,{fs:10,bold:true,color:'FFFFFF',align:'center',valign:'middle'});
      pag.text(ls.join('\n'),1.15,y,w,h,{fs:fs,color:PAL.body,ls:1.25});
      y+=h+0.14;
    });
  }
  return deck;
}
function generateResumoPPTX(results,recs,info,unidade,logo,asPDF){
  if(typeof PptxGenJS==='undefined'){alert('Biblioteca PptxGenJS não carregada.');return;}
  var deck=montarDeckResumo(results,recs,info,unidade,logo);
  var nome='resumo_executivo_'+_lim(info.cliente)+(unidade?'_'+_lim(unidade):'');
  if(asPDF){deckParaPDF(deck).save(nome+'.pdf');}
  else{deckParaPPTX(deck).writeFile({fileName:nome+'.pptx'});}
}

/* ===== r149: COMPARATIVO — apresentação (PPTX + PDF juntos) ===== */
function montarDeckComparativo(comp,info,logo){
  var deck=novoDeck();
  var nomes=comp.unidades.map(function(u){return u.unidade;});
  var sub=(info.cliente||'')+'  ·  Inventário: '+(info.dataInventario||'—');
  slideCapa(deck,'Comparativo','entre Unidades','Análise de Inventário — '+comp.unidades.length+' unidades',(info.cliente||'')+'  ·  '+(info.dataInventario||'')+'\n'+_cortar(nomes.join(' × '),160),logo,true);

  /* Ranking por métrica — até 5 unidades por slide */
  var grupo=5;
  for(var g0=0;g0<comp.unidades.length;g0+=grupo){
    var us=comp.unidades.slice(g0,g0+grupo);
    var tot=Math.ceil(comp.unidades.length/grupo),pagN=Math.floor(g0/grupo)+1;
    var s=slideFaixa(deck,'Ranking por métrica'+(tot>1?' ('+pagN+'/'+tot+')':''),sub+'  ·  verde = melhor · vermelho = pior',logo,true);
    var wM=2.7,wU=(SL.R-SL.M-wM)/us.length;
    var cab=['Métrica'].concat(us.map(function(u){return _cortar(u.unidade,24);}));
    var corpo=comp.rankings.map(function(r){
      var linha=[{t:r.label,bold:true,cor:PAL.navy}];
      us.forEach(function(u){
        var v=null;r.valores.forEach(function(x){if(x.unidade===u.unidade)v=x;});
        if(!v){linha.push('—');return;}
        var cor='333333',bold=false;
        if(r.melhor&&v.unidade===r.melhor){cor='2E7D32';bold=true;}else if(r.pior&&v.unidade===r.pior){cor='D32F2F';bold=true;}
        linha.push({t:fmtRank(r,v),cor:cor,bold:bold});
      });
      return linha;
    });
    var nLin=corpo.length+1,disp=SL.B-TOPO_C,rowH=Math.min(0.34,disp/nLin);
    s.table(cab,corpo,SL.M,TOPO_C,SL.R-SL.M,{colW:[wM].concat(us.map(function(){return wU;})),fs:us.length>4?7.5:8.5,rowH:rowH,align:['left'].concat(us.map(function(){return 'right';}))});
  }

  /* Página de gráfico 1 — Perdas, sobras e quebra */
  var s1=slideFaixa(deck,'Valor de perdas, sobras e quebra',sub,logo,true);
  imgGrafico(s1,cfgCompQuebra(comp,12),SL.M,TOPO_C,SL.R-SL.M,SL.B-TOPO_C);

  /* Página 2 — cartões Ruptura Depósito x Loja + gráfico depósito × loja */
  var s2=slideFaixa(deck,'Ruptura Depósito x Loja',sub,logo,true);
  /* r150: taxa de ruptura (mesmo cálculo da aba Ruptura) */
  var yC=cartoesUnidades(s2,comp.unidades.map(function(u){var tem=u.taxaRuptura!==null&&u.taxaRuptura!==undefined;return{titulo:u.unidade,valor:tem?PCT(u.taxaRuptura):'—',sub:tem?NUM(u.totalRupturas)+' de '+NUM(u.totalComDeposito)+' SKUs com depósito':'',cor:PAL.red};}),TOPO_C,1.1);
  imgGrafico(s2,cfgCompDepLoja(comp,12),SL.M,yC+0.12,SL.R-SL.M,SL.B-(yC+0.12));

  /* Página 3 — cartões Cobertura em dias + gráfico de estoque por curva */
  var s3=slideFaixa(deck,'Cobertura em dias e valor do estoque por curva',sub,logo,true);
  var yC3=cartoesUnidades(s3,comp.unidades.map(function(u){return{titulo:u.unidade,valor:u.coberturaGeral!==null&&u.coberturaGeral!==undefined?NUM(u.coberturaGeral)+' dias':'—',sub:'Cobertura em dias',cor:PAL.navy};}),TOPO_C,1.1);
  imgGrafico(s3,cfgCompCurvas(comp,12),SL.M,yC3+0.12,SL.R-SL.M,SL.B-(yC3+0.12));

  /* r150: Página 4 — potencial de venda perdida no mês por faixa de cobertura */
  var s4=slideFaixa(deck,'Potencial de venda perdida no mês por faixa de cobertura',sub+'  ·  faturamento mensal dos itens de cada faixa',logo,true);
  imgGrafico(s4,cfgCompPotencial(comp,12),SL.M,TOPO_C,SL.R-SL.M,SL.B-TOPO_C);
  return deck;
}
function generateComparativoApresentacao(comp,info,logo){
  if(typeof PptxGenJS==='undefined'){alert('Biblioteca PptxGenJS não carregada.');return;}
  info=info||{};
  var deck=montarDeckComparativo(comp,info,logo);
  salvarDeck(deck,'comparativo_'+_lim(info.cliente)+'_'+String(info.dataInventario||'').replace(/\//g,'-'));
}

return{generateExcel:generateExcel,generateExcelBlob:generateExcelBlob,buildExcelWorkbook:buildExcelWorkbook,generatePDF:generatePDF,generateComparativoPDF:generateComparativoPDF,generateComparativoExcel:generateComparativoExcel,generateResumoPDF:generateResumoPDF,generateResumoPPTX:generateResumoPPTX,
  /* r149 */
  generateComparativoApresentacao:generateComparativoApresentacao,fmtRank:fmtRank,BRLk:BRLk,cfgCritica:cfgCritica,valoresCurvasABC:valoresCurvasABC,
  _montarDeckResumo:montarDeckResumo,_montarDeckComparativo:montarDeckComparativo,_deckParaPPTX:deckParaPPTX,_deckParaPDF:deckParaPDF};
})();
