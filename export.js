/* export.js — v3.10 — Excel/PDF/HTML com linha de totais em todas as tabelas, análises por dimensão no HTML, gráficos por categoria no PDF (Crítica/ABC/Perda), Resumo Executivo com capa clara + cards coloridos + análises, sem listar produtos */
var Export=(function(){
"use strict";
var C={navy:'051323',green:'00B74A',red:'D32F2F',amb:'F57C00',blue:'1565C0',white:'FFFFFF',light:'F5F5F5',lightG:'F0F0F0',border:'D0D0D0',text:'333333',muted:'888888'};
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
/* r99: estilo + linha de totais no final de cada tabela (padrão pedido pelo cliente pra Excel/PDF/HTML) */
function sTOT(a){return{font:{bold:true,color:{rgb:C.navy},sz:10,name:'Arial'},fill:{fgColor:{rgb:'E8F5E9'}},alignment:{horizontal:a||'left',vertical:'center'},border:{top:{style:'medium',color:{rgb:C.green}},bottom:{style:'medium',color:{rgb:C.green}}}};}
function addDTt(ws,r,hd,dr,ca,totalRow){
  var rEnd=addDT(ws,r,hd,dr,ca);
  if(totalRow){
    var tr=rEnd-1;
    for(var c=0;c<totalRow.length;c++){var al=(ca&&ca[c])?ca[c]:'left';sC(ws,tr,c,totalRow[c]===null||totalRow[c]===undefined?'':totalRow[c],sTOT(al));}
    return rEnd+1;
  }
  return rEnd;
}
/* Soma um campo numérico de uma lista de itens, arredondando (evita ruído de ponto flutuante) */
function sumF(items,f){return R2(items.reduce(function(s,i){return s+(Number(i[f])||0);},0));}
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

/* ===== GRÁFICOS (Chart.js → PNG, para uso em PDF) ===== */
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
    var chart=new Chart(ctx,config);
    var url=chart.toBase64Image('image/png',1.0);
    chart.destroy();
    return url;
  }catch(e){console.log('Erro ao gerar gráfico:',e);return null;}
}
function chartCritica(c){
  if(c.hasCategorias&&c.categorias.length){
    var cats=c.categorias;
    return _chartPNG({type:'bar',data:{labels:cats.map(function(x){return x.nome;}),datasets:[
      {label:'Faltas (R$)',data:cats.map(function(x){return Math.abs(x.faltaVal);}),backgroundColor:'#D32F2F',borderRadius:3},
      {label:'Sobras (R$)',data:cats.map(function(x){return x.sobraVal;}),backgroundColor:'#F57C00',borderRadius:3}
    ]},options:{indexAxis:'y',plugins:{legend:{position:'top',labels:{font:{size:13}}},datalabels:{anchor:'end',align:'end',color:function(c){return c.dataset.backgroundColor;},font:{size:11,weight:'bold'},formatter:function(v){return _abbr(v);}}},scales:{x:{ticks:{font:{size:12},callback:function(v){return _abbr(v);}}},y:{ticks:{font:{size:12}}}},layout:{padding:{right:60}}}},1000,Math.min(530,Math.max(305,cats.length*80)));
  }
  return _chartPNG({type:'bar',data:{labels:['Faltas (R$)','Sobras (R$)'],datasets:[{data:[Math.abs(c.totalFaltas),c.totalSobras],backgroundColor:['#D32F2F','#F57C00'],borderRadius:4}]},options:{plugins:{legend:{display:false},datalabels:{anchor:'end',align:'top',color:'#333',font:{size:12,weight:'bold'},formatter:function(v){return _abbr(v);}}},scales:{y:{beginAtZero:true,ticks:{font:{size:12}}},x:{ticks:{font:{size:13}}}}}},900,280);
}
function chartRuptura(r){
  return _chartPNG({type:'bar',data:{labels:['Curva A','Curva B','Curva C'],datasets:[{label:'SKUs em ruptura',data:[r.rupturaA,r.rupturaB,r.rupturaC],backgroundColor:['#D32F2F','#F57C00','#888888'],borderRadius:4}]},options:{plugins:{legend:{display:false},datalabels:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:12}}},x:{ticks:{font:{size:13}}}}}},1000,330);
}
function chartDias(d){
  var fo=[{k:'Ruptura',v:d.ruptura,c:'#D32F2F'},{k:'Alto risco',v:d.altoRisco,c:'#F57C00'},{k:'Médio risco',v:d.medioRisco,c:'#FBC02D'},{k:'Cobertura ideal',v:d.coberturaIdeal,c:'#00B74A'},{k:'Excesso',v:d.excessos,c:'#1565C0'},{k:'Sem giro',v:d.semGiro,c:'#888888'}];
  return _chartPNG({type:'bar',data:{labels:fo.map(function(x){return x.k;}),datasets:[{data:fo.map(function(x){return x.v;}),backgroundColor:fo.map(function(x){return x.c;}),borderRadius:4}]},options:{plugins:{legend:{display:false},datalabels:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:12}}},x:{ticks:{font:{size:11}}}}}},1000,330);
}
function chartABC(a){
  return _chartPNG({type:'bar',data:{labels:['Curva A','Curva B','Curva C'],datasets:[{label:'Valor em estoque (R$)',data:[a.fatA.invest,a.fatB.invest,a.fatC.invest],backgroundColor:'#002B50',borderRadius:4},{label:'Faturamento (R$)',data:[a.fatA.fat,a.fatB.fat,a.fatC.fat],backgroundColor:'#61CF00',borderRadius:4}]},options:{plugins:{legend:{position:'top',labels:{font:{size:13}}},datalabels:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:12}}},x:{ticks:{font:{size:13}}}}}},1000,330);
}
/* r99: novo modelo — 3 barras por categoria (vendas/estoque/lucro), valores abreviados */
function chartABCCategoria(a){
  if(!a||!a.hasCategorias||!a.categorias||!a.categorias.length)return null;
  var cats=a.categorias;
  return _chartPNG({type:'bar',data:{labels:cats.map(function(x){return x.nome;}),datasets:[
    {label:'Vendas (%)',data:cats.map(function(x){return x.pctFat;}),backgroundColor:'#1565C0',borderRadius:3},
    {label:'Estoque (%)',data:cats.map(function(x){return x.pctInvest;}),backgroundColor:'#F57C00',borderRadius:3},
    {label:'Lucro (%)',data:cats.map(function(x){return x.pctLucro;}),backgroundColor:'#00B74A',borderRadius:3}
  ]},options:{indexAxis:'y',plugins:{legend:{position:'top',labels:{font:{size:11}}},datalabels:{anchor:'end',align:'end',color:function(c){return c.dataset.backgroundColor;},font:{size:9,weight:'bold'},formatter:function(v){return PCT(v);}}},scales:{x:{ticks:{font:{size:10},callback:function(v){return v+'%';}}},y:{ticks:{font:{size:10}}}},layout:{padding:{right:45}}}},1000,Math.min(560,Math.max(320,cats.length*90)));
}
function chartPerda(p){
  return _chartPNG({type:'bar',data:{labels:['Curva A','Curva B','Curva C'],datasets:[{label:'Perda mensal (R$)',data:[p.classA.perda*30,p.classB.perda*30,p.classC.perda*30],backgroundColor:['#D32F2F','#F57C00','#888888'],borderRadius:4}]},options:{plugins:{legend:{display:false},datalabels:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:12}}},x:{ticks:{font:{size:13}}}}}},1000,330);
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
  if(data.critica){var c=data.critica;R=addST(ws,R,'CRÍTICA DO INVENTÁRIO');R=addKR(ws,R,['ACURACIDADE','VALOR ESTOQUE','VALOR ESTOQUE CONTADO','VALOR DAS FALTAS','VALOR DAS SOBRAS','SALDO LÍQUIDO','PERDA DE ESTOQUE (%)'],[PCT(c.acuracidade),BRLi(c.valorEstoque),BRLi(c.valorEstoqueContado),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido),PCT(c.perdaEstoquePct)],[C.green,C.text,C.text,C.red,C.amb,C.red,c.perdaEstoquePct<0?C.red:C.green]);if(c.hasCategorias)R=addDTt(ws,R,['Categoria','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{0:'left',1:'right',2:'right',3:'right',4:'right'},['TOTAL',PCT(c.acuracidade),BRLi(sumF(c.categorias,'faltaVal')),BRLi(sumF(c.categorias,'sobraVal')),BRLi(sumF(c.categorias,'saldo'))]);}
  if(data.ruptura){var r=data.ruptura;R=addST(ws,R,'RUPTURA LOJA X DEPÓSITO');R=addKR(ws,R,['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CURVA A (FAT.)','RUPTURA CURVA A (LUCRO)'],[PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],[C.red,C.text,C.red,C.red]);}
  if(data.dias){var d=data.dias,fv=fxV(d.items);R=addST(ws,R,'DIAS DE ESTOQUE');R=addKR(ws,R,['COBERTURA GERAL','CURVA A','CURVA B','CURVA C','SEM GIRO'],[d.coberturaGeral+' dias',d.coberturaA+' dias',d.coberturaB+' dias',d.coberturaC+' dias',NUM(d.semGiro)],[C.text,C.text,C.text,C.text,C.red]);var fo=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];R=addDTt(ws,R,['Faixa','SKUs','% SKUs','Valor Estoque (R$)','% do Valor'],fo.map(function(f){var cn=d.items.filter(function(i){return i.faixa===f;}).length;var vl=fv.f[f]||0;return[f,cn,PCT(d.total?cn/d.total*100:0),BRLi(vl),PCT(fv.t?vl/fv.t*100:0)];}),{0:'left',1:'right',2:'right',3:'right',4:'right'},['TOTAL',d.total,PCT(100),BRLi(fv.t),PCT(100)]);if(d.hasCategorias)R=addDTt(ws,R,['Categoria','Cobertura média','Val. Ruptura+Alto risco (R$)','Val. Sem giro (R$)','Val. Excesso (R$)','Val. Estoque (R$)'],d.categorias.map(function(x){return[x.nome,x.mediaCobertura+' dias',BRLi(x.valorCriticos),BRLi(x.valorSemGiro),BRLi(x.valorExcessos),BRLi(x.valorEstoque)];}),{0:'left',1:'right',2:'right',3:'right',4:'right',5:'right'},['TOTAL',d.coberturaGeral+' dias',BRLi(sumF(d.categorias,'valorCriticos')),BRLi(sumF(d.categorias,'valorSemGiro')),BRLi(sumF(d.categorias,'valorExcessos')),BRLi(sumF(d.categorias,'valorEstoque'))]);}
  if(data.abc){var a=data.abc;R=addST(ws,R,'INVESTIMENTO ABC');R=addKR(ws,R,['VALOR TOTAL EM ESTOQUE','FATURAMENTO 90D','LUCRO 90D','SKUS'],[BRLi(a.totalInvest),BRLi(a.totalFat),BRLi(a.totalLucro),NUM(a.items.length)],[C.text,C.green,C.green,C.text]);R=addDTt(ws,R,['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],{0:'center',1:'right',2:'right',3:'right',4:'right'},['TOTAL',BRLi(a.totalInvest),PCT(100),BRLi(a.totalFat),PCT(100)]);}
  if(data.perda){var pe=data.perda;R=addST(ws,R,'PROJEÇÃO DE PERDA');R=addKR(ws,R,['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS'],[BRLi(pe.totalPerdaFat),BRLi(pe.totalPerdaLucro),BRLi(pe.perdaMensal),NUM(pe.totalSKUs)],[C.red,C.red,C.red,C.text]);R=addDTt(ws,R,['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',pe.classA.count,BRLi(pe.classA.perda),BRLi(pe.classA.lucro),PCT(pe.classA.pct),BRLi(pe.classA.perda*30)],['B',pe.classB.count,BRLi(pe.classB.perda),BRLi(pe.classB.lucro),PCT(pe.classB.pct),BRLi(pe.classB.perda*30)],['C',pe.classC.count,BRLi(pe.classC.perda),BRLi(pe.classC.lucro),PCT(pe.classC.pct),BRLi(pe.classC.perda*30)]],{0:'center',1:'right',2:'right',3:'right',4:'right',5:'right'},['TOTAL',pe.totalSKUs,BRLi(pe.totalPerdaFat),BRLi(pe.totalPerdaLucro),PCT(100),BRLi(pe.perdaMensal)]);}
  ws['!cols']=[{wch:28},{wch:18},{wch:16},{wch:18},{wch:16},{wch:18},{wch:16},{wch:16}];ws['!rows']=[{hpt:28},{hpt:20}];
  XLSX.utils.book_append_sheet(wb,ws,'Dashboard');

  /* CRITICA RESUMO + TOP20 */
  if(sel.criticaResumo&&data.critica){var wsC={},rw=0,c=data.critica;rw=addBH(wsC,rw,info,pd,6);rw=addST(wsC,rw,'RESUMO DA CRÍTICA');var sL=['ACURACIDADE','SKUs analisados','SKUs sem divergência','SKUs com falta','SKUs com sobra','Valor estoque (sistema)','Valor estoque contado','Valor das faltas','Valor das sobras','Saldo líquido','Perda de estoque (%)'],sV=[PCT(c.acuracidade),c.totalSKUs,c.okCount,c.faltaCount,c.sobraCount,BRLi(c.valorEstoque),BRLi(c.valorEstoqueContado),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido),PCT(c.perdaEstoquePct)];for(var i=0;i<sL.length;i++){sC(wsC,rw+i,0,sL[i],sB('left',true));sC(wsC,rw+i,1,sV[i],sB('right'));}rw+=sL.length+1;
  if(c.hasCategorias){rw=addST(wsC,rw,'RESULTADO POR CATEGORIA');rw=addDTt(wsC,rw,['Categoria','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{0:'left',1:'right',2:'right',3:'right',4:'right'},['TOTAL',PCT(c.acuracidade),BRLi(sumF(c.categorias,'faltaVal')),BRLi(sumF(c.categorias,'sobraVal')),BRLi(sumF(c.categorias,'saldo'))]);}
  var ct=top20Cat(c.items);var tH=['SKU','Descrição','Qtd Sist','Qtd Contada','Dif. Qtd','Dif. R$'],tA={0:'left',1:'left',2:'right',3:'right',4:'right',5:'right'};
  ct.forEach(function(cat){if(cat.faltas.length){rw=addST(wsC,rw,'TOP '+cat.faltas.length+' FALTAS — '+cat.nome);rw=addDT(wsC,rw,tH,cat.faltas.map(function(i){return[i.sku,i.descricao,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor)];}),tA);}if(cat.sobras.length){rw=addST(wsC,rw,'TOP '+cat.sobras.length+' SOBRAS — '+cat.nome);rw=addDT(wsC,rw,tH,cat.sobras.map(function(i){return[i.sku,i.descricao,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor)];}),tA);}if(cat.zerados.length){rw=addST(wsC,rw,'TOP '+cat.zerados.length+' ZERADOS — '+cat.nome);rw=addDT(wsC,rw,['SKU','Descrição','Qtd Sistema','Valor Perdido'],cat.zerados.map(function(i){return[i.sku,i.descricao,i.qtdSistema,BRL(i.qtdSistema*i.custoUnit)];}),{0:'left',1:'left',2:'right',3:'right'});}});
  wsC['!cols']=[{wch:16},{wch:32},{wch:14},{wch:14},{wch:12},{wch:16}];wsC['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsC,'Crítica - Resumo');}

  if(sel.criticaDetalhe&&data.critica){var wsCD={},rw=0;rw=addBH(wsCD,rw,info,pd,8);rw=addDTt(wsCD,rw,['SKU','Descrição','Categoria','Qtd Sistema','Qtd Contada','Dif. Qtd','Dif. R$','Status'],data.critica.items.map(function(i){return[i.sku,i.descricao,i.categoria,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor),i.status];}),{0:'left',1:'left',2:'left',3:'right',4:'right',5:'right',6:'right',7:'center'},['TOTAL ('+data.critica.items.length+' itens)','','',sumF(data.critica.items,'qtdSistema'),sumF(data.critica.items,'qtdContada'),sumF(data.critica.items,'difQtd'),BRL(sumF(data.critica.items,'difValor')),'']);wsCD['!cols']=[{wch:14},{wch:32},{wch:18},{wch:12},{wch:12},{wch:10},{wch:14},{wch:10}];wsCD['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsCD,'Crítica - Detalhado');}

  if(sel.ruptura&&data.ruptura){var wsR={},rw=0;rw=addBH(wsR,rw,info,pd,9);rw=addDTt(wsR,rw,['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Depósito','Qtd Loja','Venda Méd/Dia','Fat. Méd/Dia'],data.ruptura.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abc_valorVendido90||'C',i.abc_lucro90||'C',i.deposito,i.loja,R2(i.vendaMediaDia),BRL(i.fatMediaDia||0)];}),{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right'},['TOTAL ('+data.ruptura.items.length+' itens)','','','','',sumF(data.ruptura.items,'deposito'),sumF(data.ruptura.items,'loja'),'',BRL(sumF(data.ruptura.items,'fatMediaDia'))]);wsR['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:14},{wch:10},{wch:14},{wch:14}];wsR['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsR,'Ruptura Loja x Depósito');}

  if(sel.dias&&data.dias){var wsD={},rw=0;rw=addBH(wsD,rw,info,pd,9);rw=addDTt(wsD,rw,['SKU','Descrição','Categoria','Qtd Estoque','Venda Méd/Dia','Dias Estoque','Cobertura','Valor Estoque','ABC Fat.'],data.dias.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.qtdEstoque,R2(i.vendaMediaDia),i.diasEstoque!==null?R2(i.diasEstoque):'—',i.faixa,BRL(i.valorEstoque),i.abcFat];}),{0:'left',1:'left',2:'left',3:'right',4:'right',5:'right',6:'left',7:'right',8:'center'},['TOTAL ('+data.dias.items.length+' itens)','','',sumF(data.dias.items,'qtdEstoque'),'','','',BRL(sumF(data.dias.items,'valorEstoque')),'']);wsD['!cols']=[{wch:14},{wch:32},{wch:18},{wch:12},{wch:14},{wch:12},{wch:18},{wch:16},{wch:10}];wsD['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsD,'Dias de Estoque');}

  if(sel.abc&&data.abc){var wsA={},rw=0;rw=addBH(wsA,rw,info,pd,9);var totFatCol=data.abc.items.reduce(function(s,i){return s+(i.qtdEstoque<=0?(i.perdaVenda30||0):i.fat90);},0);rw=addDTt(wsA,rw,['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Estoque','Custo Unit.','Valor Estoque','Fat. 90d / Perda proj. 30d'],data.abc.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abcFat,i.abcLucro,i.qtdEstoque,BRL(i.custoUnit),BRL(i.valorInvestido),i.qtdEstoque<=0?('-'+BRL(i.perdaVenda30)+' (30d)'):BRL(i.fat90)];}),{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right'},['TOTAL ('+data.abc.items.length+' itens)','','','','',sumF(data.abc.items,'qtdEstoque'),'',BRL(sumF(data.abc.items,'valorInvestido')),BRL(R2(totFatCol))]);wsA['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:12},{wch:14},{wch:16},{wch:22}];wsA['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsA,'Investimento ABC');}

  if(sel.perda&&data.perda){var wsP={},rw=0;rw=addBH(wsP,rw,info,pd,11);rw=addDTt(wsP,rw,['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Venda 90d','Venda Méd/Dia','Perda Fat./Dia','Perda Lucro/Dia','Perda Fat./Mês','Perda Lucro/Mês'],data.perda.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abcFat,i.abcLucro||'C',R2(i.qtdVendida||0),R2(i.vendaMediaDia),BRL(i.perdaFatDia),BRL(i.perdaLucroDia),BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];}),{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right',9:'right',10:'right'},['TOTAL ('+data.perda.items.length+' itens)','','','','',sumF(data.perda.items,'qtdVendida'),'',BRL(sumF(data.perda.items,'perdaFatDia')),BRL(sumF(data.perda.items,'perdaLucroDia')),BRL(sumF(data.perda.items,'perdaFatMes')),BRL(sumF(data.perda.items,'perdaLucroMes'))]);wsP['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:12},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14}];wsP['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsP,'Projeção de Perda');}

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
  function hdr(){doc.setFillColor(5,19,35);doc.rect(0,0,W,22,'F');if(logo){try{doc.addImage(logo,'PNG',M,7,32,8);}catch(e){}}doc.setFontSize(9);doc.setTextColor(255,255,255);doc.text((info.cliente||'')+' — '+(info.unidade||''),W-M,7,{align:'right'});doc.setFontSize(7);doc.setTextColor(200,220,255);doc.text('Inventário: '+(info.dataInventario||'—'),W-M,12,{align:'right'});doc.setTextColor(180,180,200);doc.text('Processado em '+pd,W-M,17,{align:'right'});y=28;}
  function ftr(pg){doc.setFontSize(7);doc.setTextColor(150,150,150);doc.text('Formula Code Tecnologia, Gestão e Automação',M,H-6);doc.text('Página '+pg,W-M,H-6,{align:'right'});doc.setDrawColor(200,200,200);doc.line(M,H-10,W-M,H-10);}
  function chk(n){if(y+n>H-18){doc.addPage();hdr();ftr(doc.getNumberOfPages());}}
  function ttl(t){chk(12);doc.setFontSize(14);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFontSize(8);doc.setTextColor(150,150,150);doc.setFont(undefined,'normal');doc.text('Relatório gerado automaticamente pelo sistema Formula Code',M,y);y+=8;}
  function sec(t){chk(10);doc.setFontSize(11);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFont(undefined,'normal');}
  function aT(h,b,o,foot){var cfg={startY:y,head:[h],body:b,margin:{left:M,right:M},headStyles:{fillColor:[5,19,35],fontSize:7,fontStyle:'bold',halign:'left'},bodyStyles:{fontSize:7,halign:'left'},alternateRowStyles:{fillColor:[245,245,245]},styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},columnStyles:o||{}};chk(20);if(foot){cfg.foot=[foot];cfg.footStyles={fillColor:[232,245,233],textColor:[5,19,35],fontStyle:'bold',fontSize:7,lineWidth:0.3,lineColor:[97,207,0]};}doc.autoTable(cfg);y=doc.lastAutoTable.finalY+6;}
  function kpi(lb,vl,cl){chk(18);var cw=(W-2*M)/lb.length;doc.setFillColor(245,245,245);doc.roundedRect(M,y-2,W-2*M,16,2,2,'F');for(var i=0;i<lb.length;i++){var x=M+i*cw+4;doc.setFontSize(7);doc.setTextColor(150,150,150);doc.setFont(undefined,'bold');doc.text(lb[i],x,y+3);doc.setFontSize(11);doc.setFont(undefined,'bold');var cc=cl[i]||[51,51,51];doc.setTextColor(cc[0],cc[1],cc[2]);doc.text(String(vl[i]),x,y+10);}doc.setFont(undefined,'normal');y+=20;}
  function bloco(txt){chk(16);doc.setFontSize(8);doc.setTextColor(80,80,80);doc.setFont(undefined,'normal');var lines=doc.splitTextToSize(txt,W-2*M);doc.text(lines,M,y);y+=lines.length*3.5+4;}
  function img(url,h){if(!url)return;chk(h+8);try{doc.addImage(url,'PNG',M,y,W-2*M,h);}catch(e){}y+=h+8;}

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
    sec('Gráfico — Faltas e sobras'+(c.hasCategorias?' por categoria':''));
    img(chartCritica(c),c.hasCategorias?Math.min(95,Math.max(55,c.categorias.length*15)):55);
    if(c.hasCategorias){sec('Resultado por categoria');aT(['Categoria','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}},['TOTAL',PCT(c.acuracidade),BRLi(sumF(c.categorias,'faltaVal')),BRLi(sumF(c.categorias,'sobraVal')),BRLi(sumF(c.categorias,'saldo'))]);}
    var ct=top20Cat(c.items);var tO={2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}};
    ct.forEach(function(cat){
      if(cat.faltas.length){sec('Top '+cat.faltas.length+' faltas — '+cat.nome);aT(['SKU','Descrição','Qtd Sist','Qtd Contada','Dif. Qtd','Dif. R$'],cat.faltas.map(function(i){return[i.sku,i.descricao,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor)];}),tO);}
      if(cat.sobras.length){sec('Top '+cat.sobras.length+' sobras — '+cat.nome);aT(['SKU','Descrição','Qtd Sist','Qtd Contada','Dif. Qtd','Dif. R$'],cat.sobras.map(function(i){return[i.sku,i.descricao,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor)];}),tO);}
      if(cat.zerados.length){sec('Top '+cat.zerados.length+' zerados — '+cat.nome);aT(['SKU','Descrição','Qtd Sistema','Valor Perdido'],cat.zerados.map(function(i){return[i.sku,i.descricao,i.qtdSistema,BRL(i.qtdSistema*i.custoUnit)];}),{2:{halign:'right'},3:{halign:'right'}});}
    });
  }
  else if(rt==='ruptura'){
    var r=data.ruptura;var temVendas=r.items.some(function(i){return i.vendaMediaDia>0;});
    ttl('Ruptura Loja x Depósito — Resumo executivo');
    var iaRup=window._iaResumos&&window._iaResumos.ruptura;
    sec('Análise');bloco(iaRup||Engine.gerarAnaliseRuptura(r,info));
    sec('Metodologia');bloco(metRuptura(info.diasVenda));
    sec('Indicadores gerais');kpi(['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CURVA A (FAT.)','RUPTURA CURVA A (LUCRO)'],[PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],[[211,47,47],[51,51,51],[211,47,47],[211,47,47]]);
    sec('Gráfico — Rupturas por curva ABC');img(chartRuptura(r),60);
    if(temVendas){
      sec('Rupturas curva A — Top 30');var topA=r.items.filter(function(i){return i.abc_valorVendido90==='A';}).slice(0,30);
      aT(['SKU','Descrição','Categoria','ABC Fat.','Qtd Dep.','Venda Méd/Dia','Fat. Méd/Dia'],topA.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abc_valorVendido90,i.deposito,R2(i.vendaMediaDia),BRL(i.fatMediaDia||0)];}),{4:{halign:'right'},5:{halign:'right'},6:{halign:'right'}});
    }else{
      /* Sem vendas: listar por categoria com qtd depósito */
      var catMap={};r.items.forEach(function(i){var c=i.categoria||'Sem categoria';if(!catMap[c])catMap[c]=[];catMap[c].push(i);});
      Object.keys(catMap).sort().forEach(function(cat){
        var itens=catMap[cat].sort(function(a,b){return b.deposito-a.deposito;});
        sec(cat+' — '+itens.length+' itens em ruptura');
        aT(['SKU','Descrição','Qtd Depósito'],itens.map(function(i){return[i.sku,i.descricao,i.deposito];}),{2:{halign:'right'}},['TOTAL',(''),sumF(itens,'deposito')]);
      });
    }
  }
  else if(rt==='dias'){
    var d=data.dias,fv=fxV(d.items);ttl('Dias de estoque — Resumo executivo');
    var iaDias=window._iaResumos&&window._iaResumos.dias_estoque;
    sec('Análise');bloco(iaDias||Engine.gerarAnaliseDias(d,info));sec('Metodologia');bloco(metDias(info.diasVenda));
    sec('Indicadores gerais');kpi(['COBERTURA GERAL','CURVA A','CURVA B','CURVA C'],[d.coberturaGeral+' dias',d.coberturaA+' dias',d.coberturaB+' dias',d.coberturaC+' dias'],[[51,51,51],[211,47,47],[245,124,0],[136,136,136]]);
    sec('Gráfico — Distribuição por faixa de cobertura');img(chartDias(d),60);
    sec('Distribuição por faixa');var fo=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];
    aT(['Faixa','SKUs','% SKUs','Valor Estoque (R$)','% do Valor'],fo.map(function(f){var cn=d.items.filter(function(i){return i.faixa===f;}).length;var vl=fv.f[f]||0;return[f,cn,PCT(d.total?cn/d.total*100:0),BRLi(vl),PCT(fv.t?vl/fv.t*100:0)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}},['TOTAL',d.total,PCT(100),BRLi(fv.t),PCT(100)]);
    if(d.hasCategorias){sec('Cobertura por categoria');aT(['Categoria','Cobertura média','Rupt+Alto risco (R$)','Sem giro (R$)','Excesso (R$)','Val. estoque (R$)'],d.categorias.map(function(x){return[x.nome,x.mediaCobertura+' dias',BRLi(x.valorCriticos),BRLi(x.valorSemGiro),BRLi(x.valorExcessos),BRLi(x.valorEstoque)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}},['TOTAL',d.coberturaGeral+' dias',BRLi(sumF(d.categorias,'valorCriticos')),BRLi(sumF(d.categorias,'valorSemGiro')),BRLi(sumF(d.categorias,'valorExcessos')),BRLi(sumF(d.categorias,'valorEstoque'))]);}
    var pH=['SKU','Descrição','Categoria','Dias est.','ABC Fat.','Val. estoque'],pO={3:{halign:'right'},5:{halign:'right'}};
    fo.forEach(function(fx){var it=d.items.filter(function(i){return i.faixa===fx;});if(it.length){sec(fx+' — '+it.length+' itens');aT(pH,it.slice(0,50).map(function(i){return[i.sku,i.descricao,i.categoria||'',i.diasEstoque!==null?R2(i.diasEstoque):'—',i.abcFat,BRL(i.valorEstoque)];}),pO);if(it.length>50){doc.setFontSize(7);doc.setTextColor(150,150,150);doc.text('... e mais '+(it.length-50)+' itens (ver Excel)',M,y);y+=4;}}});
  }
  else if(rt==='abc'){
    var a=data.abc;ttl('Investimento por curva ABC — Resumo executivo');
    var iaABC=window._iaResumos&&window._iaResumos.abc;
    sec('Análise');bloco(iaABC||Engine.gerarAnaliseABC(a,info));sec('Metodologia');bloco(metABC(info.diasVenda));
    sec('Indicadores gerais');kpi(['VALOR TOTAL EM ESTOQUE','FATURAMENTO 90D','LUCRO 90D','SKUS ANALISADOS'],[BRLi(a.totalInvest),BRLi(a.totalFat),BRLi(a.totalLucro),NUM(a.items.length)],[[51,51,51],[0,183,74],[0,183,74],[51,51,51]]);
    sec('Gráfico — Valor em estoque × Faturamento por curva');img(chartABC(a),60);
    sec('Curva ABC por faturamento');aT(['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}},['TOTAL',BRLi(a.totalInvest),PCT(100),BRLi(a.totalFat),PCT(100)]);
    sec('Curva ABC por lucro');aT(['Curva','Valor Estoque (R$)','% Estoque','Lucro (R$)','% Lucro'],[['A',BRLi(a.lucA.invest),PCT(a.lucA.pctInvest),BRLi(a.lucA.luc),PCT(a.lucA.pctLuc)],['B',BRLi(a.lucB.invest),PCT(a.lucB.pctInvest),BRLi(a.lucB.luc),PCT(a.lucB.pctLuc)],['C',BRLi(a.lucC.invest),PCT(a.lucC.pctInvest),BRLi(a.lucC.luc),PCT(a.lucC.pctLuc)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}},['TOTAL',BRLi(a.totalInvest),PCT(100),BRLi(a.totalLucro),PCT(100)]);
    if(a.hasCategorias){
      sec('Investimento em Estoque por categoria');
      img(chartABCCategoria(a),Math.min(95,Math.max(55,a.categorias.length*15)));
      aT(['Categoria','Val. estoque (R$)','Vendas (%)','Estoque (%)','Lucro (%)','Cobertura'],a.categorias.map(function(x){return[x.nome,BRLi(x.investimento),PCT(x.pctFat),PCT(x.pctInvest),PCT(x.pctLucro),x.coberturaDias+'d'];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}},['TOTAL',BRLi(sumF(a.categorias,'investimento')),PCT(100),PCT(100),PCT(100),'—']);
    }
  }
  else if(rt==='perda'){
    var p=data.perda;ttl('Projeção de venda perdida — Resumo executivo');
    var iaPerda=window._iaResumos&&window._iaResumos.perda;
    sec('Análise');bloco(iaPerda||Engine.gerarAnalisePerda(p,info));sec('Metodologia');bloco(metPerda(info.diasVenda));
    sec('Indicadores gerais');kpi(['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS'],[BRLi(p.totalPerdaFat),BRLi(p.totalPerdaLucro),BRLi(p.perdaMensal),NUM(p.totalSKUs)],[[211,47,47],[211,47,47],[211,47,47],[51,51,51]]);
    sec('Impacto por curva ABC');aT(['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',p.classA.count,BRLi(p.classA.perda),BRLi(p.classA.lucro),PCT(p.classA.pct),BRLi(p.classA.perda*30)],['B',p.classB.count,BRLi(p.classB.perda),BRLi(p.classB.lucro),PCT(p.classB.pct),BRLi(p.classB.perda*30)],['C',p.classC.count,BRLi(p.classC.perda),BRLi(p.classC.lucro),PCT(p.classC.pct),BRLi(p.classC.perda*30)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}},['TOTAL',p.totalSKUs,BRLi(p.totalPerdaFat),BRLi(p.totalPerdaLucro),PCT(100),BRLi(p.perdaMensal)]);
    if(p.hasCategorias){
      sec('Perda projetada por categoria');
      img(chartPerdaCategoria(p),Math.min(95,Math.max(55,p.categorias.length*15)));
      aT(['Categoria','Rupturas','Perda fat./dia (R$)','Perda mensal (R$)'],p.categorias.map(function(x){return[x.nome,x.totalRupturas,BRLi(x.perdaFatDia),BRLi(x.perdaMensal)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'}},['TOTAL',p.totalSKUs,BRLi(sumF(p.categorias,'perdaFatDia')),BRLi(sumF(p.categorias,'perdaMensal'))]);
    }
    sec('Gráfico — Perda mensal projetada por curva');img(chartPerda(p),60);
    var pH2=['SKU','Descrição','Categoria','Perda Fat./Mês','Perda Lucro/Mês'],pO2={3:{halign:'right'},4:{halign:'right'}};
    ['A','B','C'].forEach(function(cls){var it=p.items.filter(function(i){return i.abcFat===cls;}).sort(function(a,b){return b.perdaFatMes-a.perdaFatMes;});if(it.length){sec('Curva '+cls+' — '+it.length+' itens');aT(pH2,it.map(function(i){return[i.sku,i.descricao,i.categoria||'',BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];}),pO2,['TOTAL','','',BRL(sumF(it,'perdaFatMes')),BRL(sumF(it,'perdaLucroMes'))]);}});
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

  function hdr(){doc.setFillColor(5,19,35);doc.rect(0,0,W,22,'F');if(logo){try{doc.addImage(logo,'PNG',M,7,32,8);}catch(e){}}doc.setFontSize(9);doc.setTextColor(255,255,255);doc.text('COMPARATIVO — '+(info.cliente||''),W-M,7,{align:'right'});doc.setFontSize(7);doc.setTextColor(200,220,255);doc.text(comp.unidades.map(function(u){return u.unidade;}).join(' × '),W-M,12,{align:'right'});doc.setTextColor(180,180,200);doc.text('Inventário: '+(info.dataInventario||'—')+' | Gerado em '+pd,W-M,17,{align:'right'});y=28;}
  function ftr(pg){doc.setFontSize(7);doc.setTextColor(150,150,150);doc.text('Formula Code Tecnologia, Gestão e Automação',M,H-6);doc.text('Página '+pg,W-M,H-6,{align:'right'});doc.setDrawColor(200,200,200);doc.line(M,H-10,W-M,H-10);}
  function chk(n){if(y+n>H-18){doc.addPage();hdr();ftr(doc.getNumberOfPages());}}
  function ttl(t){chk(12);doc.setFontSize(14);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFontSize(8);doc.setTextColor(150,150,150);doc.setFont(undefined,'normal');doc.text('Relatório comparativo gerado pelo Sistema Formula Code',M,y);y+=8;}
  function sec(t){chk(10);doc.setFontSize(11);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFont(undefined,'normal');}
  function bloco(txt){if(!txt)return;chk(16);doc.setFontSize(8);doc.setTextColor(80,80,80);doc.setFont(undefined,'normal');var lines=doc.splitTextToSize(txt,W-2*M);doc.text(lines,M,y);y+=lines.length*3.5+4;}
  function aT(h,b,o){chk(20);doc.autoTable({startY:y,head:[h],body:b,margin:{left:M,right:M},headStyles:{fillColor:[5,19,35],fontSize:7,fontStyle:'bold',halign:'left'},bodyStyles:{fontSize:7,halign:'left'},alternateRowStyles:{fillColor:[245,245,245]},styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},columnStyles:o||{}});y=doc.lastAutoTable.finalY+6;}
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
      var fmtVal=v.valor;
      if(r.fmt==='pct')fmtVal=PCT(v.valor);
      else if(r.fmt==='brl')fmtVal=BRLi(v.valor);
      else if(r.fmt==='num')fmtVal=NUM(v.valor);
      row.push(String(fmtVal)+(r.melhor&&v.unidade===r.melhor?' ★':'')+(r.pior&&v.unidade===r.pior?' ▼':''));
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

  /* SKU overlap */
  if(comp.skuOverlap){
    R=addKR(ws,R,['SKUS ÚNICOS','EM COMUM','SOBREPOSIÇÃO'],[NUM(comp.skuOverlap.totalUnique),NUM(comp.skuOverlap.emComum),PCT(comp.skuOverlap.pctComum)],[C.text,C.green,C.green]);
  }

  /* Rankings */
  R=addST(ws,R,'RANKING POR MÉTRICA');
  var hdr=['Métrica'];comp.unidades.forEach(function(u){hdr.push(u.unidade);});
  var rows=comp.rankings.map(function(r){
    var row=[r.label];
    r.valores.forEach(function(v){
      var fmtVal=v.valor;
      if(r.fmt==='pct')fmtVal=PCT(v.valor);
      else if(r.fmt==='brl')fmtVal=BRLi(v.valor);
      else if(r.fmt==='num')fmtVal=NUM(v.valor);
      row.push(fmtVal);
    });
    return row;
  });
  var ca={0:'left'};for(var i=1;i<=nUnits;i++)ca[i]='right';
  R=addDT(ws,R,hdr,rows,ca);

  var colWidths=[{wch:28}];for(var j=0;j<nUnits;j++)colWidths.push({wch:20});
  ws['!cols']=colWidths;
  ws['!rows']=[{hpt:28},{hpt:20}];
  XLSX.utils.book_append_sheet(wb,ws,'Comparativo');

  var out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  var blob=new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  var url=URL.createObjectURL(blob);var a=document.createElement('a');a.href=url;
  a.download='comparativo_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.xlsx';
  a.click();URL.revokeObjectURL(url);
}

/* ===== r68: RESUMO EXECUTIVO — PDF Documento ===== */
function generateResumoPDF(results,recs,info,unidade,logo){
  var jsPDF=window.jspdf.jsPDF;
  var doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  var W=210,H=297,M=20,cw=W-2*M;
  var y=M;
  var navy=[5,19,35],green=[0,183,74];

  /* r99: capa com fundo branco + tarja navy (igual aos outros relatórios), pra imprimir sem gastar tinta de fundo */
  function hdrR(){doc.setFillColor.apply(doc,navy);doc.rect(0,0,W,22,'F');if(logo){try{doc.addImage(logo,'PNG',M,7,32,8);}catch(e){}}doc.setFontSize(9);doc.setTextColor(255,255,255);doc.text((info.cliente||'')+(unidade?' — '+unidade:''),W-M,7,{align:'right'});doc.setFontSize(7);doc.setTextColor(200,220,255);doc.text('Inventário: '+(info.dataInventario||'—'),W-M,12,{align:'right'});doc.setTextColor(180,180,200);doc.text('Resumo Executivo',W-M,17,{align:'right'});y=32;}
  function ftrR(pg){doc.setFontSize(7);doc.setTextColor(150,150,150);doc.text('Formula Code Tecnologia, Gestão e Automação',M,H-6);doc.text('Página '+pg,W-M,H-6,{align:'right'});doc.setDrawColor(200,200,200);doc.line(M,H-10,W-M,H-10);}
  hdrR();ftrR(1);
  doc.setFontSize(22);doc.setTextColor.apply(doc,navy);doc.text('Resumo Executivo',M,y+8);
  doc.setFontSize(11);doc.setTextColor.apply(doc,green);doc.text('Análise de Inventário',M,y+16);
  y+=26;doc.setDrawColor(220,220,220);doc.line(M,y,W-M,y);y+=10;

  doc.setFontSize(16);doc.setTextColor.apply(doc,navy);
  doc.text('Indicadores-Chave',M,y);y+=10;

  var red=[211,47,47],amb=[245,124,0];
  var tintRed=[253,236,236],tintAmb=[255,244,229],tintGreen=[232,247,222],tintNavy=[240,242,245];
  var dims=[];
  if(results.critica)dims.push({label:'Acuracidade',valor:PCT(results.critica.acuracidade),detalhe:NUM(results.critica.totalSKUs)+' SKUs · '+NUM(results.critica.faltaCount)+' faltas · '+NUM(results.critica.sobraCount)+' sobras',cor:results.critica.acuracidade<90?red:(results.critica.acuracidade<95?amb:green),tint:results.critica.acuracidade<90?tintRed:(results.critica.acuracidade<95?tintAmb:tintGreen)});
  if(results.ruptura)dims.push({label:'Ruptura',valor:PCT(results.ruptura.taxaRuptura),detalhe:NUM(results.ruptura.totalRupturas)+' SKUs em falta · Curva A: '+NUM(results.ruptura.rupturaA),cor:results.ruptura.taxaRuptura>10?red:(results.ruptura.taxaRuptura>5?amb:green),tint:results.ruptura.taxaRuptura>10?tintRed:(results.ruptura.taxaRuptura>5?tintAmb:tintGreen)});
  if(results.dias)dims.push({label:'Cobertura',valor:Engine.round2(results.dias.coberturaGeral)+' dias',detalhe:'Sem giro: '+NUM(results.dias.semGiro)+' · Excesso: '+NUM(results.dias.excessos),cor:(results.dias.coberturaGeral<15||results.dias.coberturaGeral>60)?red:green,tint:(results.dias.coberturaGeral<15||results.dias.coberturaGeral>60)?tintRed:tintGreen});
  if(results.abc)dims.push({label:'Investimento',valor:BRLi(results.abc.totalInvest),detalhe:'Faturamento: '+BRLi(results.abc.totalFat),cor:navy,tint:tintNavy});
  if(results.perda)dims.push({label:'Perda Mensal',valor:BRLi(results.perda.perdaMensal),detalhe:NUM(results.perda.totalSKUs)+' SKUs identificados',cor:results.perda.perdaMensal>50000?red:(results.perda.perdaMensal>10000?amb:green),tint:results.perda.perdaMensal>50000?tintRed:(results.perda.perdaMensal>10000?tintAmb:tintGreen)});

  /* Cards em grade de 2 colunas — fundo tintado + barra colorida (r99: "cards coloridos") */
  var cardGap=6,cardW=(cw-cardGap)/2,cardH=32;
  dims.forEach(function(d,i){
    var col=i%2,row=Math.floor(i/2);
    var cx=M+col*(cardW+cardGap),cy=y+row*(cardH+cardGap);
    doc.setFillColor.apply(doc,d.tint);doc.roundedRect(cx,cy,cardW,cardH,2,2,'F');
    doc.setFillColor.apply(doc,d.cor);doc.rect(cx,cy,1.6,cardH,'F');
    doc.setFontSize(8);doc.setTextColor.apply(doc,green);
    doc.text(d.label.toUpperCase(),cx+7,cy+8);
    doc.setFontSize(16);doc.setTextColor.apply(doc,navy);
    doc.text(d.valor,cx+7,cy+17);
    doc.setFontSize(8);doc.setTextColor(85,102,119);
    doc.text(doc.splitTextToSize(d.detalhe,cardW-12),cx+7,cy+24);
  });
  y+=Math.ceil(dims.length/2)*(cardH+cardGap)+8;

  function ckR(n){if(y+n>270){doc.addPage();hdrR();ftrR(doc.getNumberOfPages());}}
  function imgR(url,h){if(!url)return;ckR(h+10);try{doc.addImage(url,'PNG',M,y,cw,h);}catch(e){}y+=h+10;}
  function tituloR(t){ckR(14);doc.setFontSize(11);doc.setTextColor.apply(doc,navy);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFont(undefined,'normal');}
  function blocoR(txt){if(!txt)return;ckR(16);doc.setFontSize(8);doc.setTextColor(80,80,80);var lines=doc.splitTextToSize(txt,cw);doc.text(lines,M,y);y+=lines.length*3.5+6;}

  /* r99: Análises por dimensão — antes não existiam no Resumo, só no relatório por aba */
  if(y>230){doc.addPage();hdrR();ftrR(doc.getNumberOfPages());}
  doc.setFontSize(16);doc.setTextColor.apply(doc,navy);doc.text('Análises',M,y);y+=10;
  if(results.critica){tituloR('Crítica do Inventário (Acuracidade)');blocoR(Engine.gerarAnaliseCritica(results.critica,info));}
  if(results.ruptura){tituloR('Ruptura Loja x Depósito');blocoR(Engine.gerarAnaliseRuptura(results.ruptura,info));}
  if(results.dias){tituloR('Dias de Estoque (Cobertura)');blocoR(Engine.gerarAnaliseDias(results.dias,info));}
  if(results.abc){tituloR('Investimento em Estoque');blocoR(Engine.gerarAnaliseABC(results.abc,info));}
  if(results.perda){tituloR('Projeção de Perda');blocoR(Engine.gerarAnalisePerda(results.perda,info));}

  if(y>245){doc.addPage();hdrR();ftrR(doc.getNumberOfPages());}
  doc.setFontSize(16);doc.setTextColor.apply(doc,navy);
  doc.text('Gráficos',M,y);y+=10;

  /* r99: gráfico por categoria (novo modelo) quando disponível, com fallback pra curva ABC */
  if(results.critica){var cc=results.critica;tituloR('Crítica — Faltas e sobras'+(cc.hasCategorias?' por categoria':''));imgR(chartCritica(cc),cc.hasCategorias?Math.min(80,Math.max(50,cc.categorias.length*12)):50);}
  if(results.ruptura){tituloR('Ruptura por curva ABC');imgR(chartRuptura(results.ruptura),55);}
  if(results.dias){tituloR('Distribuição por faixa de cobertura');imgR(chartDias(results.dias),55);}
  if(results.abc){var aa=results.abc;tituloR('Investimento em Estoque'+(aa.hasCategorias?' por categoria':' por curva ABC'));imgR(aa.hasCategorias?chartABCCategoria(aa):chartABC(aa),aa.hasCategorias?Math.min(85,Math.max(50,aa.categorias.length*14)):55);}
  if(results.perda){var pp=results.perda;tituloR('Perda projetada'+(pp.hasCategorias?' por categoria':' por curva ABC'));imgR(pp.hasCategorias?chartPerdaCategoria(pp):chartPerda(pp),pp.hasCategorias?Math.min(85,Math.max(50,pp.categorias.length*14)):55);}

  /* Recomendações */
  if(recs.length>0){
    y+=5;
    if(y>240){doc.addPage();hdrR();ftrR(doc.getNumberOfPages());}
    doc.setFontSize(16);doc.setTextColor.apply(doc,navy);
    doc.text('Recomendações',M,y);y+=10;

    recs.forEach(function(rec){
      if(y>265){doc.addPage();hdrR();ftrR(doc.getNumberOfPages());}
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

/* ===== r68: RESUMO EXECUTIVO — PPTX (e PDF apresentação) ===== */
function generateResumoPPTX(results,recs,info,unidade,logo,asPDF){
  if(typeof PptxGenJS==='undefined'){alert('Biblioteca PptxGenJS não carregada.');return;}
  var pres=new PptxGenJS();
  pres.layout='LAYOUT_16x9';

  var NAVY='002B50',GREEN='61CF00',DARK='051323',BODY='556677',LABEL='8899AA';

  function addLogo(s){if(logo){try{s.addImage({data:logo,x:7.05,y:4.57,w:2.7,h:0.67});}catch(e){}}}
  function splitBg(s){
    s.background={fill:'FFFFFF'};
    s.addShape(pres.ShapeType.rect,{x:6.8,y:0,w:3.2,h:5.625,fill:{color:DARK}});
    addLogo(s);
  }

  /* Slide 1: Capa */
  var s1=pres.addSlide();
  s1.background={fill:DARK};
  s1.addShape(pres.ShapeType.rect,{x:0,y:0,w:10,h:0.04,fill:{color:GREEN}});
  s1.addText([
    {text:'Resumo ',options:{color:'FFFFFF',fontSize:36,fontFace:'Calibri',bold:true}},
    {text:'Executivo',options:{color:GREEN,fontSize:36,fontFace:'Calibri',bold:true}}
  ],{x:0.8,y:0.8,w:8.4,h:1.0,isTextBox:true,margin:0});
  s1.addText('Análise de Inventário',{x:0.8,y:1.9,w:6,h:0.4,fontSize:16,fontFace:'Calibri',color:GREEN,isTextBox:true,margin:0});
  var sub2=info.cliente||'';if(unidade)sub2+=' — '+unidade;sub2+='  ·  '+(info.dataInventario||'');
  s1.addText(sub2,{x:0.8,y:2.5,w:6,h:0.4,fontSize:13,fontFace:'Arial',color:LABEL,isTextBox:true,margin:0});
  if(logo){try{s1.addImage({data:logo,x:3.2,y:4.37,w:3.6,h:0.9});}catch(e){}}

  /* Slide 2: Indicadores */
  var s2=pres.addSlide();
  splitBg(s2);
  s2.addText('01 / INDICADORES-CHAVE',{x:0.8,y:0.45,w:5.5,h:0.25,fontSize:9,fontFace:'Arial',bold:true,color:GREEN,isTextBox:true,margin:0,charSpacing:3});
  s2.addText([
    {text:'Diagnóstico do ',options:{color:DARK}},
    {text:'Inventário',options:{color:GREEN}}
  ],{x:0.8,y:0.9,w:5.6,h:0.6,fontSize:28,fontFace:'Calibri',bold:true,isTextBox:true,margin:0});

  var dims=[];
  if(results.critica)dims.push({label:'ACURACIDADE',valor:results.critica.acuracidade+'%',sub:'Meta: 95%',accent:results.critica.acuracidade<95});
  if(results.ruptura)dims.push({label:'RUPTURA',valor:results.ruptura.taxaRuptura+'%',sub:results.ruptura.totalRupturas+' SKUs em falta',accent:true});
  if(results.dias)dims.push({label:'COBERTURA',valor:Engine.round2(results.dias.coberturaGeral)+' dias',sub:'Sem giro: '+results.dias.semGiro+' SKUs',accent:results.dias.coberturaGeral>45});
  if(results.abc)dims.push({label:'INVESTIMENTO',valor:'R$ '+Engine.formatNum(results.abc.totalInvest),sub:'Giro: '+Engine.round2(results.abc.totalFat/results.abc.totalInvest)+'x',accent:false});
  if(results.perda)dims.push({label:'PERDA MENSAL',valor:'R$ '+Engine.formatNum(results.perda.perdaMensal),sub:results.perda.totalSKUs+' SKUs',accent:true});

  dims.forEach(function(d,i){
    var ky=1.7+i*0.7;
    s2.addShape(pres.ShapeType.rect,{x:0.8,y:ky,w:0.04,h:0.5,fill:{color:d.accent?GREEN:NAVY}});
    s2.addText(d.valor,{x:1.05,y:ky,w:2.0,h:0.3,fontSize:22,fontFace:'Calibri',bold:true,color:DARK,isTextBox:true,margin:0});
    s2.addText(d.label,{x:3.1,y:ky,w:1.8,h:0.18,fontSize:8,fontFace:'Arial',bold:true,color:GREEN,isTextBox:true,margin:0,charSpacing:2});
    s2.addText(d.sub,{x:3.1,y:ky+0.18,w:2.5,h:0.18,fontSize:10,fontFace:'Arial',color:BODY,isTextBox:true,margin:0});
  });

  /* Right panel info */
  s2.addText((info.cliente||'')+'\n'+(unidade||''),{x:7.05,y:0.8,w:2.7,h:0.5,fontSize:12,fontFace:'Arial',bold:true,color:'FFFFFF',isTextBox:true,margin:0});
  s2.addText((info.dataInventario||'')+'\n'+(info.diasVenda||90)+' dias de venda',{x:7.05,y:1.5,w:2.7,h:0.5,fontSize:10,fontFace:'Arial',color:LABEL,isTextBox:true,margin:0,lineSpacingMultiple:1.4});

  /* Slide 2.5: Gráficos */
  var chartDims=[];
  if(results.critica){
    var cc=results.critica;
    if(cc.hasCategorias&&cc.categorias.length){
      chartDims.push({title:'Crítica — Faltas/sobras por categoria',data:[{name:'Faltas (R$)',labels:cc.categorias.map(function(x){return x.nome;}),values:cc.categorias.map(function(x){return Math.abs(x.faltaVal);})},{name:'Sobras (R$)',labels:cc.categorias.map(function(x){return x.nome;}),values:cc.categorias.map(function(x){return x.sobraVal;})}],colors:['D32F2F','F57C00'],legend:true});
    }else{
      chartDims.push({title:'Crítica — Faltas × Sobras (R$)',data:[{name:'R$',labels:['Faltas','Sobras'],values:[Math.abs(cc.totalFaltas),cc.totalSobras]}],colors:['D32F2F'],legend:false});
    }
  }
  if(results.ruptura){var rr=results.ruptura;chartDims.push({title:'Ruptura por curva ABC',data:[{name:'SKUs',labels:['Curva A','Curva B','Curva C'],values:[rr.rupturaA,rr.rupturaB,rr.rupturaC]}],colors:['D32F2F'],legend:false});}
  if(results.dias){var dd=results.dias;chartDims.push({title:'Cobertura por faixa',data:[{name:'SKUs',labels:['Ruptura','Alto risco','Médio risco','Ideal','Excesso','Sem giro'],values:[dd.ruptura,dd.altoRisco,dd.medioRisco,dd.coberturaIdeal,dd.excessos,dd.semGiro]}],colors:['D32F2F'],legend:false});}
  if(results.abc){var aa=results.abc;chartDims.push({title:'Investimento por curva ABC',data:[{name:'Valor estoque (R$)',labels:['Curva A','Curva B','Curva C'],values:[aa.fatA.invest,aa.fatB.invest,aa.fatC.invest]}],colors:['002B50'],legend:false});}
  if(results.perda){var pp=results.perda;chartDims.push({title:'Perda mensal por curva',data:[{name:'R$/mês',labels:['Curva A','Curva B','Curva C'],values:[pp.classA.perda*30,pp.classB.perda*30,pp.classC.perda*30]}],colors:['D32F2F'],legend:false});}

  if(chartDims.length){
    var sG=pres.addSlide();
    splitBg(sG);
    sG.addText('02 / GRÁFICOS',{x:0.8,y:0.45,w:5.5,h:0.25,fontSize:9,fontFace:'Arial',bold:true,color:GREEN,isTextBox:true,margin:0,charSpacing:3});
    sG.addText([
      {text:'Visão ',options:{color:DARK}},
      {text:'Gráfica',options:{color:GREEN}}
    ],{x:0.8,y:0.9,w:5.6,h:0.6,fontSize:28,fontFace:'Calibri',bold:true,isTextBox:true,margin:0});
    var cols=chartDims.length>1?2:1,rows=Math.ceil(chartDims.length/cols);
    var gx=0.8,gy=1.75,gw=5.7,gh=3.55,gapX=0.3,gapY=0.35;
    var cw2=(gw-gapX*(cols-1))/cols,ch2=(gh-gapY*(rows-1))/rows;
    chartDims.forEach(function(cd,i){
      var col=i%cols,row=Math.floor(i/cols);
      var cx=gx+col*(cw2+gapX),cy=gy+row*(ch2+gapY);
      sG.addText(cd.title,{x:cx,y:cy,w:cw2,h:0.2,fontSize:8,fontFace:'Arial',bold:true,color:BODY,isTextBox:true,margin:0});
      try{
        sG.addChart(pres.ChartType.bar,cd.data,{x:cx,y:cy+0.24,w:cw2,h:ch2-0.28,barDir:'col',showLegend:cd.legend,legendPos:'b',legendFontSize:7,showTitle:false,showValue:false,chartColors:cd.colors,catAxisLabelFontSize:7,valAxisLabelFontSize:7,catAxisLabelColor:BODY,valAxisLabelColor:BODY});
      }catch(e){}
    });
    sG.addText((info.cliente||'')+'\n'+(unidade||''),{x:7.05,y:0.8,w:2.7,h:0.5,fontSize:12,fontFace:'Arial',bold:true,color:'FFFFFF',isTextBox:true,margin:0});
  }

  /* Slide 3: Recomendações */
  if(recs.length>0){
    var s3=pres.addSlide();
    splitBg(s3);
    s3.addText('03 / RECOMENDAÇÕES',{x:0.8,y:0.45,w:5.5,h:0.25,fontSize:9,fontFace:'Arial',bold:true,color:GREEN,isTextBox:true,margin:0,charSpacing:3});
    s3.addText([
      {text:'Ações ',options:{color:DARK}},
      {text:'Recomendadas',options:{color:GREEN}}
    ],{x:0.8,y:0.9,w:5.6,h:0.6,fontSize:28,fontFace:'Calibri',bold:true,isTextBox:true,margin:0});

    var ry=1.7;
    var maxRecs=Math.min(recs.length,6); /* Max 6 per slide */
    for(var i=0;i<maxRecs;i++){
      var rec=recs[i];
      var color=rec.prioridade===1?'CC3333':(rec.prioridade===2?'E8872B':'228B22');
      var icon=rec.prioridade===1?'!':(rec.prioridade===2?'▲':'✓');

      s3.addShape(pres.ShapeType.ellipse,{x:0.8,y:ry,w:0.25,h:0.25,fill:{color:color}});
      s3.addText(icon,{x:0.8,y:ry,w:0.25,h:0.25,fontSize:10,fontFace:'Arial',bold:true,color:'FFFFFF',align:'center',valign:'middle',isTextBox:true,margin:0});

      /* Truncar texto para caber */
      var txt=rec.texto;
      if(txt.length>150)txt=txt.substring(0,147)+'...';
      s3.addText(txt,{x:1.15,y:ry,w:5.3,h:0.55,fontSize:10,fontFace:'Arial',color:BODY,isTextBox:true,margin:0,lineSpacingMultiple:1.3});
      ry+=0.62;
    }

    s3.addText((info.cliente||'')+'\n'+(unidade||''),{x:7.05,y:0.8,w:2.7,h:0.5,fontSize:12,fontFace:'Arial',bold:true,color:'FFFFFF',isTextBox:true,margin:0});
  }

  /* Generate */
  if(asPDF){
    /* Download as PDF — pptxgenjs doesn't natively export PDF,
       so we save as PPTX and note the limitation */
    pres.writeFile({fileName:'resumo_executivo_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'.pptx'});
    setTimeout(function(){alert('Para converter em PDF: abra o arquivo PPTX no PowerPoint ou Google Slides e exporte como PDF.');},500);
  }else{
    pres.writeFile({fileName:'resumo_executivo_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'.pptx'});
  }
}

/* ========================================================================
   v3.9 — EXPORTAÇÃO EM HTML (por aba + relatório completo + comparativo)
   Gera um arquivo .html autônomo (CSS embutido, sem dependências externas
   além da imagem do gráfico de Perda, já convertida em PNG embutido).
   ======================================================================== */
function _abbr(v){var neg=(v||0)<0;var a=Math.abs(v||0);var s;if(a>=1000000)s=(a/1000000).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'M';else if(a>=1000)s=Math.round(a/1000)+'k';else s=Math.round(a).toLocaleString('pt-BR');return(neg?'-':'')+'R$ '+s;}

function chartPerdaCategoria(p){
  if(!p||!p.hasCategorias||!p.categorias||!p.categorias.length)return null;
  var cats=p.categorias;
  return _chartPNG({type:'bar',data:{labels:cats.map(function(c){return c.nome;}),datasets:[{label:'Perda mensal',data:cats.map(function(c){return c.perdaMensal;}),backgroundColor:'#D32F2F',borderRadius:3}]},options:{indexAxis:'y',plugins:{legend:{display:false},datalabels:{anchor:'end',align:'end',color:'#D32F2F',font:{size:11,weight:'bold'},formatter:function(v){return _abbr(v);}}},scales:{x:{ticks:{font:{size:11},callback:function(v){return _abbr(v);}}},y:{ticks:{font:{size:11}}}},layout:{padding:{right:70}}}},1000,Math.min(530,Math.max(220,cats.length*60)));
}

var _HTML_CSS=":root{--navy:#051323;--green:#00B74A;--red:#D32F2F;--amb:#F57C00;--blue:#1565C0;--bg:#f5f6f8;--border:#e0e0e0;--muted:#888;--light:#f5f5f5}"
+"*{box-sizing:border-box;margin:0;padding:0}"
+"body{font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:var(--bg);color:#333;font-size:14px;line-height:1.5;padding-bottom:40px}"
+".wrap{max-width:1100px;margin:0 auto}"
+".hdr{background:var(--navy);padding:16px 24px;display:flex;align-items:center;gap:16px}"
+".hdr img{height:34px}"
+".hdr-r{margin-left:auto;text-align:right;color:#fff}"
+".hdr-t{font-size:14px;font-weight:600}"
+".hdr-s{font-size:11px;color:rgba(255,255,255,.6);margin-top:2px}"
+".body{background:#fff;padding:24px;margin:0 0 30px}"
+".section-title{font-size:14px;font-weight:600;color:var(--navy);margin:20px 0 12px}"
+".metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:18px}"
+".metric{background:var(--light);border-radius:8px;padding:12px 14px}"
+".metric-label{font-size:11px;color:var(--muted);text-transform:uppercase;margin-bottom:3px}"
+".metric-value{font-size:20px;font-weight:700}"
+".metric-detail{font-size:11px;color:var(--muted);margin-top:2px}"
+".cat-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:16px}"
+".cat-card{border-radius:8px;padding:12px 14px;border:1px solid var(--border)}"
+".cat-name{font-size:13px;font-weight:600;margin-bottom:6px}"
+".cat-row{display:flex;justify-content:space-between;font-size:12px;padding:2px 0}"
+".cat-label{color:var(--muted)}"
+".cat-val{font-weight:600}"
+".text-red{color:var(--red)}.text-green{color:var(--green)}.text-amber{color:var(--amb)}.text-blue{color:var(--blue)}.text-muted{color:var(--muted)}"
+".minibar-legend{display:flex;flex-wrap:wrap;gap:16px;margin-bottom:12px;font-size:12px;color:var(--muted)}"
+".minibar-dot{width:10px;height:10px;border-radius:2px;display:inline-block;margin-right:5px;vertical-align:middle}"
+".minibar-list{display:flex;flex-direction:column;margin-bottom:16px}"
+".minibar-row{display:grid;grid-template-columns:170px 1fr 120px;gap:14px;align-items:center;padding:10px 0;border-bottom:1px solid var(--border)}"
+".minibar-row.inv{grid-template-columns:190px 1fr 84px}"
+".minibar-row:last-child{border-bottom:none}"
+".minibar-cat{font-weight:600;font-size:13px;color:var(--navy)}"
+".minibar-catval{font-size:11px;color:var(--muted);margin-top:2px}"
+".minibar-bars{display:flex;flex-direction:column;gap:6px}"
+".minibar-line{display:flex;align-items:center;gap:8px}"
+".minibar-tag{font-size:10px;color:var(--muted);width:46px;flex-shrink:0}"
+".minibar-track{flex:1;background:var(--light);border-radius:4px;height:12px;overflow:hidden}"
+".minibar-fill{height:100%;border-radius:4px}"
+".minibar-fill.falta{background:var(--red)}.minibar-fill.sobra{background:var(--amb)}.minibar-fill.vendas{background:var(--blue)}.minibar-fill.estoque{background:var(--amb)}.minibar-fill.lucro{background:var(--green)}"
+".minibar-cov{background:var(--light);border-radius:8px;padding:8px 6px;text-align:center}"
+".minibar-cov-label{display:block;font-size:9px;text-transform:uppercase;color:var(--muted);margin-bottom:2px}"
+".minibar-cov-val{display:block;font-size:14px;font-weight:700;color:var(--navy)}"
+".analysis-box{background:var(--light);border-left:3px solid var(--green);border-radius:6px;padding:12px 16px;margin-bottom:16px}"
+".analysis-box.meta{border-left-color:var(--blue)}"
+".analysis-box p{font-size:12.5px;line-height:1.6;color:#333;margin:0 0 8px}"
+".analysis-box p:last-child{margin-bottom:0}"
+".totals-row td{font-weight:700;color:var(--navy);background:#E8F5E9;border-top:2px solid var(--green);border-bottom:2px solid var(--green)}"
+".rescard{border-radius:12px;padding:16px 18px;border-left:4px solid var(--navy);background:#fff;box-shadow:0 1px 4px rgba(0,0,0,.06)}"
+".rescard.ok{border-left-color:var(--green);background:#F1FAF3}"
+".rescard.alerta{border-left-color:var(--amb);background:#FFF8F0}"
+".rescard.critico{border-left-color:var(--red);background:#FDF2F2}"
+".rescard-label{font-size:10px;text-transform:uppercase;letter-spacing:1.5px;color:var(--green);font-weight:700;margin-bottom:6px}"
+".rescard-valor{font-size:24px;font-weight:700;color:var(--navy);line-height:1.1}"
+".rescard-detalhe{font-size:11px;color:var(--muted);margin-top:6px}"
+".minibar-val{font-size:11px;font-weight:600;min-width:66px;text-align:right;flex-shrink:0}"
+".minibar-saldo{border-radius:8px;padding:8px 10px;text-align:center}"
+".minibar-saldo.neg{background:#FDEAEA;border:1px solid var(--red)}.minibar-saldo.pos{background:#E8F8EF;border:1px solid var(--green)}.minibar-saldo.zero{background:var(--light);border:1px solid var(--border)}"
+".minibar-saldo-label{display:block;font-size:9px;text-transform:uppercase;color:var(--muted);margin-bottom:2px}"
+".minibar-saldo-val{display:block;font-size:14px;font-weight:700}"
+".minibar-saldo.neg .minibar-saldo-val{color:var(--red)}.minibar-saldo.pos .minibar-saldo-val{color:var(--green)}"
+".cat-table-wrap{overflow-x:auto;margin-bottom:16px;border:1px solid var(--border);border-radius:8px}"
+".cat-table{width:100%;border-collapse:collapse;font-size:12.5px}"
+".cat-table th{background:var(--navy);color:#fff;text-align:left;padding:8px 10px;font-size:11px;white-space:nowrap}"
+".cat-table th.num,.cat-table td.num{text-align:right}"
+".cat-table td{padding:7px 10px;border-bottom:1px solid var(--border)}"
+".cat-table tr:last-child td{border-bottom:none}"
+".cat-table tr:nth-child(even) td{background:var(--light)}"
+".table-wrap{border:1px solid var(--border);border-radius:8px;overflow-x:auto;margin-bottom:12px}"
+".data-table{width:100%;border-collapse:collapse;font-size:12.5px}"
+".data-table th{background:var(--navy);color:#fff;text-align:left;padding:8px 10px;font-size:10.5px;white-space:nowrap}"
+".data-table td{padding:6px 10px;border-bottom:1px solid var(--border)}"
+".data-table tr:nth-child(even) td{background:var(--light)}"
+".text-right{text-align:right}.text-center{text-align:center}"
+".badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10.5px;font-weight:600}"
+".loss-cards{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px}"
+".loss-card{border-radius:8px;padding:14px;text-align:center}"
+".loss-card.a{background:#FDEAEA;border:1px solid var(--red)}.loss-card.b{background:#FFF3E0;border:1px solid var(--amb)}.loss-card.c{background:var(--light);border:1px solid var(--border)}"
+".loss-title{font-size:11px;font-weight:600;margin-bottom:4px}"
+".loss-card.a .loss-title,.loss-card.a .loss-main{color:var(--red)}.loss-card.b .loss-title,.loss-card.b .loss-main{color:var(--amb)}.loss-card.c .loss-title{color:var(--muted)}"
+".loss-main{font-size:20px;font-weight:700;margin-bottom:2px}.loss-sub{font-size:11px;color:var(--muted)}"
+".chart-wrap img{max-width:100%;display:block;margin-bottom:16px}"
+".note{font-size:11.5px;color:var(--muted);margin-top:10px}"
+".ftr{max-width:1100px;margin:0 auto;padding:16px 24px;font-size:11px;color:var(--muted);text-align:center}"
+"h2{font-size:17px}"
+"@media(max-width:768px){.minibar-row,.minibar-row.inv{grid-template-columns:1fr;gap:6px}.loss-cards{grid-template-columns:1fr}}";

function _htmlPage(titulo,info,pd,bodyHtml,logo){
  var h='<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">';
  h+='<title>'+titulo+' — '+(info.cliente||'')+'</title><style>'+_HTML_CSS+'</style></head><body>';
  h+='<div class="hdr">';
  if(logo)h+='<img src="'+logo+'" alt="Formula Code">';
  h+='<div class="hdr-r"><div class="hdr-t">'+titulo+'</div><div class="hdr-s">'+(info.cliente||'')+(info.unidade?' — '+info.unidade:'')+' · Inventário: '+(info.dataInventario||'—')+' · Processado em '+pd+'</div></div>';
  h+='</div>';
  h+='<div class="wrap"><div class="body">'+bodyHtml+'</div></div>';
  h+='<div class="ftr">Formula Code Tecnologia, Gestão e Automação · Gerado em '+new Date().toLocaleString('pt-BR')+'</div>';
  h+='</body></html>';
  return h;
}
function _downloadHTML(filename,htmlStr){
  var blob=new Blob([htmlStr],{type:'text/html;charset=utf-8'});
  var url=URL.createObjectURL(blob);
  var a=document.createElement('a');a.href=url;a.download=filename;a.click();URL.revokeObjectURL(url);
}

function _catCardsHtml(catList,fields){
  var h='<div class="cat-grid">';
  catList.forEach(function(cat){
    h+='<div class="cat-card"><div class="cat-name">'+cat.nome+'</div>';
    fields.forEach(function(f){var val=f.fmt?f.fmt(cat[f.key]):cat[f.key];h+='<div class="cat-row"><span class="cat-label">'+f.label+'</span><span class="cat-val">'+val+'</span></div>';});
    h+='</div>';
  });
  h+='</div>';
  return h;
}
function _catTableHtml(catList,columns,totalRow){
  var h='<div class="cat-table-wrap"><table class="cat-table"><thead><tr><th>Categoria</th>';
  columns.forEach(function(c){h+='<th class="num">'+c.label+'</th>';});
  h+='</tr></thead><tbody>';
  catList.forEach(function(cat){
    h+='<tr><td>'+cat.nome+'</td>';
    columns.forEach(function(c){var v=c.fmt?c.fmt(cat[c.key]):cat[c.key];h+='<td class="num">'+v+'</td>';});
    h+='</tr>';
  });
  if(totalRow){h+='<tr class="totals-row"><td>TOTAL</td>';totalRow.forEach(function(v){h+='<td class="num">'+(v===null||v===undefined?'':v)+'</td>';});h+='</tr>';}
  h+='</tbody></table></div>';
  return h;
}
function _minibarFaltaSobraHtml(catList){
  var maxVal=0;
  catList.forEach(function(c){maxVal=Math.max(maxVal,Math.abs(c.faltaVal||0),c.sobraVal||0);});
  if(maxVal<=0)maxVal=1;
  var h='<div class="minibar-legend"><span><span class="minibar-dot" style="background:#D32F2F"></span>Faltas (R$)</span><span><span class="minibar-dot" style="background:#F57C00"></span>Sobras (R$)</span></div><div class="minibar-list">';
  catList.forEach(function(c){
    var wF=Math.round(Math.abs(c.faltaVal||0)/maxVal*100),wS=Math.round((c.sobraVal||0)/maxVal*100);
    var saldoCls=c.saldo<0?'neg':(c.saldo>0?'pos':'zero');
    h+='<div class="minibar-row"><div class="minibar-cat">'+c.nome+'</div><div class="minibar-bars">';
    h+='<div class="minibar-line"><span class="minibar-tag">Faltas</span><div class="minibar-track"><div class="minibar-fill falta" style="width:'+wF+'%"></div></div><span class="minibar-val text-red">'+_abbr(Math.abs(c.faltaVal||0))+'</span></div>';
    h+='<div class="minibar-line"><span class="minibar-tag">Sobras</span><div class="minibar-track"><div class="minibar-fill sobra" style="width:'+wS+'%"></div></div><span class="minibar-val text-amber">'+_abbr(c.sobraVal||0)+'</span></div>';
    h+='</div><div class="minibar-saldo '+saldoCls+'"><span class="minibar-saldo-label">Saldo</span><span class="minibar-saldo-val">'+BRLi(c.saldo)+'</span></div></div>';
  });
  h+='</div>';
  return h;
}
function _minibarVendaEstoqueHtml(catList){
  var maxPct=0;
  catList.forEach(function(c){maxPct=Math.max(maxPct,c.pctFat||0,c.pctInvest||0,c.pctLucro||0);});
  if(maxPct<=0)maxPct=1;
  var h='<div class="minibar-legend"><span><span class="minibar-dot" style="background:#1565C0"></span>Participação nas vendas</span><span><span class="minibar-dot" style="background:#F57C00"></span>Participação no valor do estoque</span><span><span class="minibar-dot" style="background:#00B74A"></span>Participação no lucro</span></div><div class="minibar-list">';
  catList.forEach(function(c){
    var wV=Math.round((c.pctFat||0)/maxPct*100),wE=Math.round((c.pctInvest||0)/maxPct*100),wL=Math.round((c.pctLucro||0)/maxPct*100);
    h+='<div class="minibar-row inv"><div class="minibar-cat">'+c.nome+'<div class="minibar-catval">'+BRLi(c.investimento)+' em estoque</div></div><div class="minibar-bars">';
    h+='<div class="minibar-line"><span class="minibar-tag">Vendas</span><div class="minibar-track"><div class="minibar-fill vendas" style="width:'+wV+'%"></div></div><span class="minibar-val">'+PCT(c.pctFat)+'</span></div>';
    h+='<div class="minibar-line"><span class="minibar-tag">Estoque</span><div class="minibar-track"><div class="minibar-fill estoque" style="width:'+wE+'%"></div></div><span class="minibar-val">'+PCT(c.pctInvest)+'</span></div>';
    h+='<div class="minibar-line"><span class="minibar-tag">Lucro</span><div class="minibar-track"><div class="minibar-fill lucro" style="width:'+wL+'%"></div></div><span class="minibar-val">'+PCT(c.pctLucro)+'</span></div>';
    h+='</div><div class="minibar-cov"><span class="minibar-cov-label">Cobertura</span><span class="minibar-cov-val">'+NUM(c.coberturaDias)+'d</span></div></div>';
  });
  h+='</div>';
  return h;
}
function _dataTableHtml(cols,rows,totalRow){
  var h='<div class="table-wrap"><table class="data-table"><thead><tr>';
  cols.forEach(function(c){h+='<th class="'+(c.align||'')+'">'+c.label+'</th>';});
  h+='</tr></thead><tbody>';
  rows.forEach(function(r){
    h+='<tr>';
    cols.forEach(function(c){
      var v=c.render?c.render(r):r[c.f],out;
      if(c.render)out=v;
      else if(c.brl)out=BRL(v);else if(c.n)out=(v===null||v===undefined)?'—':NUMx(v);else out=(v===null||v===undefined||v==='')?'—':v;
      h+='<td class="'+(c.align||'')+'">'+out+'</td>';
    });
    h+='</tr>';
  });
  if(totalRow){
    h+='<tr class="totals-row"><td>TOTAL ('+rows.length+' itens)</td>';
    for(var i=1;i<cols.length;i++){h+='<td class="'+(cols[i].align||'')+'">'+(totalRow[i-1]===null||totalRow[i-1]===undefined?'':totalRow[i-1])+'</td>';}
    h+='</tr>';
  }
  h+='</tbody></table></div>';
  return h;
}
var NUMx=function(v){return Number(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});};
function _analiseHtml(texto,metodologia){
  var h='<div class="section-title">Análise</div><div class="analysis-box">';
  (texto||'').split('\n\n').forEach(function(par){h+='<p>'+par+'</p>';});
  h+='</div>';
  if(metodologia){
    h+='<div class="section-title">Metodologia</div><div class="analysis-box meta">';
    metodologia.split('\n\n').forEach(function(par){h+='<p>'+par+'</p>';});
    h+='</div>';
  }
  return h;
}

function _bodyCritica(c,info){
  info=info||{};
  var h=_analiseHtml(Engine.gerarAnaliseCritica(c,info),metCritica());
  h+='<div class="metrics">';
  h+='<div class="metric"><div class="metric-label">Acuracidade</div><div class="metric-value text-green">'+PCT(c.acuracidade)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Valor estoque contado</div><div class="metric-value">'+BRLi(c.valorEstoqueContado)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Valor das faltas</div><div class="metric-value text-red">'+BRLi(c.totalFaltas)+'</div><div class="metric-detail">'+NUM(c.faltaCount)+' SKUs</div></div>';
  h+='<div class="metric"><div class="metric-label">Valor das sobras</div><div class="metric-value text-amber">'+BRLi(c.totalSobras)+'</div><div class="metric-detail">'+NUM(c.sobraCount)+' SKUs</div></div>';
  h+='<div class="metric"><div class="metric-label">Saldo líquido</div><div class="metric-value text-red">'+BRLi(c.saldoLiquido)+'</div></div>';
  h+='</div>';
  if(c.hasCategorias){
    h+='<div class="section-title">Resultado por categoria</div>';
    h+=_catCardsHtml(c.categorias,[{label:'Acuracidade',key:'acuracidade',fmt:PCT},{label:'Faltas',key:'faltaVal',fmt:BRLi},{label:'Sobras',key:'sobraVal',fmt:BRLi},{label:'Saldo',key:'saldo',fmt:BRLi}]);
    h+=_minibarFaltaSobraHtml(c.categorias);
  }
  h+='<div class="section-title">Itens ('+NUM(c.items.length)+')</div>';
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'Categoria',f:'categoria'},{label:'Qtd sist.',f:'qtdSistema',align:'text-right',n:true},{label:'Qtd cont.',f:'qtdContada',align:'text-right',n:true},{label:'Dif. qtd',f:'difQtd',align:'text-right',n:true},{label:'Dif. R$',f:'difValor',align:'text-right',brl:true},{label:'Status',f:'status',align:'text-center'}],c.items,
    ['','',NUMx(sumF(c.items,'qtdSistema')),NUMx(sumF(c.items,'qtdContada')),NUMx(sumF(c.items,'difQtd')),BRL(sumF(c.items,'difValor')),'']);
  return h;
}
function _bodyRuptura(r,info){
  info=info||{};
  var h=_analiseHtml(Engine.gerarAnaliseRuptura(r,info),metRuptura(info.diasVenda));
  h+='<div class="metrics">';
  h+='<div class="metric"><div class="metric-label">Taxa de ruptura geral</div><div class="metric-value text-red">'+PCT(r.taxaRuptura)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">SKUs em ruptura</div><div class="metric-value">'+NUM(r.totalRupturas)+'</div><div class="metric-detail">de '+NUM(r.totalComDeposito)+' com depósito</div></div>';
  h+='<div class="metric"><div class="metric-label">Ruptura curva A (fat.)</div><div class="metric-value text-red">'+PCT(r.taxaA)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Ruptura curva A (lucro)</div><div class="metric-value text-red">'+PCT(r.taxaALucro)+'</div></div>';
  h+='</div>';
  h+='<div class="loss-cards"><div class="loss-card a"><div class="loss-title">Curva A em ruptura</div><div class="loss-main">'+NUM(r.rupturaA)+' SKUs</div></div><div class="loss-card b"><div class="loss-title">Curva B em ruptura</div><div class="loss-main">'+NUM(r.rupturaB)+' SKUs</div></div><div class="loss-card c"><div class="loss-title">Curva C em ruptura</div><div class="loss-main">'+NUM(r.rupturaC)+' SKUs</div></div></div>';
  if(r.hasCategorias){
    h+='<div class="section-title">Ruptura Loja x Depósito por categoria</div>';
    h+=_catCardsHtml(r.categorias,[{label:'Rupturas',key:'totalRupturas',fmt:NUM},{label:'Taxa ruptura',key:'taxa',fmt:PCT},{label:'Rupturas curva A',key:'rupturaA',fmt:NUM},{label:'Perda fat./dia',key:'perdaDia',fmt:BRLi}]);
  }
  h+='<div class="section-title">Itens ('+NUM(r.items.length)+')</div>';
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'Categoria',f:'categoria'},{label:'ABC fat.',f:'abc_valorVendido90',align:'text-center'},{label:'Qtd depósito',f:'deposito',align:'text-right',n:true},{label:'Qtd loja',f:'loja',align:'text-right',n:true},{label:'Venda méd/dia',f:'vendaMediaDia',align:'text-right',n:true},{label:'Fat. méd/dia',f:'fatMediaDia',align:'text-right',brl:true}],r.items,
    ['','','','',NUMx(sumF(r.items,'deposito')),NUMx(sumF(r.items,'loja')),'',BRL(sumF(r.items,'fatMediaDia'))]);
  return h;
}
function _bodyDias(d,info){
  info=info||{};
  var h=_analiseHtml(Engine.gerarAnaliseDias(d,info),metDias(info.diasVenda));
  h+='<div class="metrics">';
  h+='<div class="metric"><div class="metric-label">Cobertura de estoque</div><div class="metric-value">'+d.coberturaGeral+' dias</div></div>';
  h+='<div class="metric"><div class="metric-label">Cobertura curva A</div><div class="metric-value">'+d.coberturaA+' dias</div></div>';
  h+='<div class="metric"><div class="metric-label">Cobertura curva B</div><div class="metric-value">'+d.coberturaB+' dias</div></div>';
  h+='<div class="metric"><div class="metric-label">Cobertura curva C</div><div class="metric-value">'+d.coberturaC+' dias</div></div>';
  h+='<div class="metric"><div class="metric-label">SKUs sem giro</div><div class="metric-value text-red">'+NUM(d.semGiro)+'</div></div>';
  h+='</div>';
  if(d.hasCategorias){
    h+='<div class="section-title">Cobertura por categoria</div>';
    h+=_catTableHtml(d.categorias,[{label:'Cobertura média',key:'mediaCobertura',fmt:function(v){return v+' dias';}},{label:'Val. estoque',key:'valorEstoque',fmt:BRLi},{label:'Ruptura + Alto risco',key:'valorCriticos',fmt:BRLi},{label:'Sem giro',key:'valorSemGiro',fmt:BRLi},{label:'Excesso (31+d)',key:'valorExcessos',fmt:BRLi}],
      [d.coberturaGeral+' dias',BRLi(sumF(d.categorias,'valorEstoque')),BRLi(sumF(d.categorias,'valorCriticos')),BRLi(sumF(d.categorias,'valorSemGiro')),BRLi(sumF(d.categorias,'valorExcessos'))]);
  }
  h+='<div class="section-title">Itens ('+NUM(d.items.length)+')</div>';
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'Categoria',f:'categoria'},{label:'Qtd estoque',f:'qtdEstoque',align:'text-right',n:true},{label:'Dias estoque',f:'diasEstoque',align:'text-right'},{label:'Cobertura',f:'faixa',align:'text-center'},{label:'Val. estoque',f:'valorEstoque',align:'text-right',brl:true},{label:'ABC fat.',f:'abcFat',align:'text-center'}],d.items,
    ['','',NUMx(sumF(d.items,'qtdEstoque')),'','',BRL(sumF(d.items,'valorEstoque')),'']);
  return h;
}
function _bodyABC(a,info){
  info=info||{};
  var h=_analiseHtml(Engine.gerarAnaliseABC(a,info),metABC(info.diasVenda));
  h+='<div class="metrics">';
  h+='<div class="metric"><div class="metric-label">Valor total em estoque</div><div class="metric-value">'+BRLi(a.totalInvest)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Faturamento 90 dias</div><div class="metric-value">'+BRLi(a.totalFat)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Lucro 90 dias</div><div class="metric-value">'+BRLi(a.totalLucro)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">SKUs analisados</div><div class="metric-value">'+NUM(a.items.length)+'</div></div>';
  h+='</div>';
  if(a.hasCategorias){
    h+='<div class="section-title">Investimento em Estoque</div>';
    h+=_minibarVendaEstoqueHtml(a.categorias);
  }
  h+='<div class="section-title">Itens ('+NUM(a.items.length)+')</div>';
  var totFatCol=a.items.reduce(function(s,i){return s+(i.qtdEstoque<=0?(i.perdaVenda30||0):i.fat90);},0);
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'ABC fat.',f:'abcFat',align:'text-center'},{label:'ABC lucro',f:'abcLucro',align:'text-center'},{label:'Qtd estoque',f:'qtdEstoque',align:'text-right',n:true},{label:'Valor estoque',f:'valorInvestido',align:'text-right',brl:true},{label:'Fat. 90d / Perda proj. 30d',align:'text-right',render:function(r){return r.qtdEstoque<=0?'<span class="text-red">'+BRL(r.perdaVenda30)+' (30d)</span>':BRL(r.fat90);}}],a.items,
    ['','','','',NUMx(sumF(a.items,'qtdEstoque')),BRL(sumF(a.items,'valorInvestido')),BRL(R2(totFatCol))]);
  return h;
}
function _bodyPerda(p,info){
  info=info||{};
  var h=_analiseHtml(Engine.gerarAnalisePerda(p,info),metPerda(info.diasVenda));
  h+='<div class="metrics">';
  h+='<div class="metric"><div class="metric-label">Venda perdida / dia</div><div class="metric-value text-red">'+BRLi(p.totalPerdaFat)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Lucro perdido / dia</div><div class="metric-value text-red">'+BRLi(p.totalPerdaLucro)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Perda mensal (fat.)</div><div class="metric-value text-red">'+BRLi(p.perdaMensal)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">SKUs em ruptura</div><div class="metric-value">'+NUM(p.totalSKUs)+'</div></div>';
  h+='</div>';
  h+='<div class="loss-cards"><div class="loss-card a"><div class="loss-title">Curva A — perda/dia</div><div class="loss-main">'+BRLi(p.classA.perda)+'</div></div><div class="loss-card b"><div class="loss-title">Curva B — perda/dia</div><div class="loss-main">'+BRLi(p.classB.perda)+'</div></div><div class="loss-card c"><div class="loss-title">Curva C — perda/dia</div><div class="loss-main">'+BRLi(p.classC.perda)+'</div></div></div>';
  if(p.hasCategorias){
    h+='<div class="section-title">Perda projetada por categoria</div>';
    h+=_catCardsHtml(p.categorias,[{label:'Rupturas',key:'totalRupturas',fmt:NUM},{label:'Perda fat./dia',key:'perdaFatDia',fmt:BRLi},{label:'Perda lucro/dia',key:'perdaLucroDia',fmt:BRLi},{label:'Perda mensal',key:'perdaMensal',fmt:BRLi},{label:'Rupturas A',key:'rupturaA',fmt:NUM}]);
    var img=chartPerdaCategoria(p);
    if(img)h+='<div class="chart-wrap"><img src="'+img+'" alt="Perda projetada por categoria"></div>';
  }
  h+='<div class="section-title">Itens ('+NUM(p.items.length)+')</div>';
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'Categoria',f:'categoria'},{label:'ABC fat.',f:'abcFat',align:'text-center'},{label:'Perda fat./dia',f:'perdaFatDia',align:'text-right',brl:true},{label:'Perda lucro/dia',f:'perdaLucroDia',align:'text-right',brl:true},{label:'Perda fat./mês',f:'perdaFatMes',align:'text-right',brl:true}],p.items,
    ['','','','',BRL(sumF(p.items,'perdaFatDia')),BRL(sumF(p.items,'perdaLucroDia')),BRL(sumF(p.items,'perdaFatMes'))]);
  return h;
}

var _HTML_TITLES={critica:'Crítica do Inventário (Acuracidade)',ruptura:'Ruptura Loja x Depósito',dias:'Dias de Estoque (Cobertura)',abc:'Investimento em Estoque',perda:'Projeção de Perda'};
var _HTML_BODYFN={critica:_bodyCritica,ruptura:_bodyRuptura,dias:_bodyDias,abc:_bodyABC,perda:_bodyPerda};

function generateHTML(type,data,pd,logo,info){
  info=info||{};
  var d=data[type];
  if(!d){alert('Este relatório ainda não foi gerado.');return;}
  var titulo=_HTML_TITLES[type]||type;
  var body=_HTML_BODYFN[type](d,info);
  var html=_htmlPage(titulo,info,pd,body,logo);
  _downloadHTML('auditoria_'+type+'_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.html',html);
}

function generateResumoHTML(results,recs,info,unidade,logo){
  info=info||{};
  var pd=new Date().toLocaleString('pt-BR');
  var red='text-red',amb='text-amber',green='text-green';
  var dims=[];
  if(results.critica){var c=results.critica;dims.push({label:'Acuracidade',valor:PCT(c.acuracidade),detalhe:NUM(c.totalSKUs)+' SKUs · '+NUM(c.faltaCount)+' faltas · '+NUM(c.sobraCount)+' sobras',cls:c.acuracidade<90?'critico':(c.acuracidade<95?'alerta':'ok')});}
  if(results.ruptura){var ru=results.ruptura;dims.push({label:'Ruptura',valor:PCT(ru.taxaRuptura),detalhe:NUM(ru.totalRupturas)+' SKUs em falta · Curva A: '+NUM(ru.rupturaA),cls:ru.taxaRuptura>10?'critico':(ru.taxaRuptura>5?'alerta':'ok')});}
  if(results.dias){var d=results.dias;dims.push({label:'Cobertura',valor:Engine.round2(d.coberturaGeral)+' dias',detalhe:'Sem giro: '+NUM(d.semGiro)+' · Excesso: '+NUM(d.excessos),cls:(d.coberturaGeral<15||d.coberturaGeral>60)?'critico':'ok'});}
  if(results.abc){var a=results.abc;dims.push({label:'Investimento',valor:BRLi(a.totalInvest),detalhe:'Faturamento: '+BRLi(a.totalFat),cls:''});}
  if(results.perda){var pe=results.perda;dims.push({label:'Perda Mensal',valor:BRLi(pe.perdaMensal),detalhe:NUM(pe.totalSKUs)+' SKUs identificados',cls:pe.perdaMensal>50000?'critico':(pe.perdaMensal>10000?'alerta':'ok')});}

  var h='<div class="section-title" style="font-size:20px;margin-top:0">Indicadores-Chave</div><div class="metrics" style="grid-template-columns:repeat(auto-fit,minmax(200px,1fr))">';
  dims.forEach(function(dm){h+='<div class="rescard '+dm.cls+'"><div class="rescard-label">'+dm.label+'</div><div class="rescard-valor">'+dm.valor+'</div><div class="rescard-detalhe">'+dm.detalhe+'</div></div>';});
  h+='</div>';

  h+='<div class="section-title">Análises</div>';
  if(results.critica)h+='<h2 style="color:var(--navy);margin:16px 0 4px">Crítica do Inventário (Acuracidade)</h2>'+_analiseHtml(Engine.gerarAnaliseCritica(results.critica,info));
  if(results.ruptura)h+='<h2 style="color:var(--navy);margin:16px 0 4px">Ruptura Loja x Depósito</h2>'+_analiseHtml(Engine.gerarAnaliseRuptura(results.ruptura,info));
  if(results.dias)h+='<h2 style="color:var(--navy);margin:16px 0 4px">Dias de Estoque (Cobertura)</h2>'+_analiseHtml(Engine.gerarAnaliseDias(results.dias,info));
  if(results.abc)h+='<h2 style="color:var(--navy);margin:16px 0 4px">Investimento em Estoque</h2>'+_analiseHtml(Engine.gerarAnaliseABC(results.abc,info));
  if(results.perda)h+='<h2 style="color:var(--navy);margin:16px 0 4px">Projeção de Perda</h2>'+_analiseHtml(Engine.gerarAnalisePerda(results.perda,info));

  h+='<div class="section-title">Gráficos</div>';
  if(results.critica){var cc=results.critica;var imgC=chartCritica(cc);if(imgC)h+='<div class="chart-wrap"><img src="'+imgC+'" alt="Crítica — Faltas e sobras'+(cc.hasCategorias?' por categoria':'')+'"></div>';}
  if(results.ruptura){var imgR=chartRuptura(results.ruptura);if(imgR)h+='<div class="chart-wrap"><img src="'+imgR+'" alt="Ruptura por curva ABC"></div>';}
  if(results.dias){var imgD=chartDias(results.dias);if(imgD)h+='<div class="chart-wrap"><img src="'+imgD+'" alt="Distribuição por faixa de cobertura"></div>';}
  if(results.abc){var aa=results.abc;var imgA=aa.hasCategorias?chartABCCategoria(aa):chartABC(aa);if(imgA)h+='<div class="chart-wrap"><img src="'+imgA+'" alt="Investimento em Estoque'+(aa.hasCategorias?' por categoria':' por curva ABC')+'"></div>';}
  if(results.perda){var pp=results.perda;var imgP=pp.hasCategorias?chartPerdaCategoria(pp):chartPerda(pp);if(imgP)h+='<div class="chart-wrap"><img src="'+imgP+'" alt="Perda projetada'+(pp.hasCategorias?' por categoria':' por curva ABC')+'"></div>';}

  if(recs&&recs.length){
    h+='<div class="section-title">Recomendações</div><div style="background:var(--light);border-radius:8px;padding:16px 20px;margin-bottom:20px">';
    recs.forEach(function(rec){h+='<div style="padding:8px 0;border-bottom:1px solid var(--border);font-size:13px">'+rec.texto+'</div>';});
    h+='</div>';
  }
  var html=_htmlPage('Resumo Executivo — Análise de Inventário',info,pd,h,logo);
  _downloadHTML('resumo_executivo_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.html',html);
}

function generateComparativoHTML(comp,units,info,logo){
  info=info||{};
  var pd=new Date().toLocaleString('pt-BR');
  var h='<div class="section-title" style="font-size:20px;margin-top:0">Comparativo entre unidades — '+comp.unidades.length+' unidades</div>';
  if(comp.skuOverlap){
    h+='<div class="metrics"><div class="metric"><div class="metric-label">SKUs únicos (total)</div><div class="metric-value">'+NUM(comp.skuOverlap.totalUnique)+'</div></div><div class="metric"><div class="metric-label">Em comum</div><div class="metric-value text-green">'+NUM(comp.skuOverlap.emComum)+'</div></div><div class="metric"><div class="metric-label">Sobreposição</div><div class="metric-value text-green">'+PCT(comp.skuOverlap.pctComum)+'</div></div></div>';
  }
  h+='<div class="section-title">Ranking por métrica</div>';
  h+='<div class="table-wrap"><table class="data-table"><thead><tr><th>Métrica</th>';
  comp.unidades.forEach(function(u){h+='<th class="text-right">'+u.unidade+'</th>';});
  h+='</tr></thead><tbody>';
  comp.rankings.forEach(function(r){
    h+='<tr><td>'+r.label+'</td>';
    r.valores.forEach(function(v){
      var fv=v.valor;
      if(r.fmt==='pct')fv=PCT(v.valor);else if(r.fmt==='brl')fv=BRLi(v.valor);else if(r.fmt==='num')fv=NUM(v.valor);
      var marca=(r.melhor&&v.unidade===r.melhor?' ★':'')+(r.pior&&v.unidade===r.pior?' ▼':'');
      h+='<td class="text-right">'+fv+marca+'</td>';
    });
    h+='</tr>';
  });
  h+='</tbody></table></div>';
  h+='<div class="note">★ = melhor desempenho na métrica · ▼ = pior desempenho</div>';
  var html=_htmlPage('Comparativo entre unidades',info,pd,h,logo);
  _downloadHTML('comparativo_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.html',html);
}

return{generateExcel:generateExcel,generateExcelBlob:generateExcelBlob,buildExcelWorkbook:buildExcelWorkbook,generatePDF:generatePDF,generateComparativoPDF:generateComparativoPDF,generateComparativoExcel:generateComparativoExcel,generateResumoPDF:generateResumoPDF,generateResumoPPTX:generateResumoPPTX,generateHTML:generateHTML,generateResumoHTML:generateResumoHTML,generateComparativoHTML:generateComparativoHTML};

})();
