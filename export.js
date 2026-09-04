/* export.js — v3.0 — Excel formatado + Dashboard + PDF com resumo executivo */
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
    ]},options:{indexAxis:'y',plugins:{legend:{position:'top',labels:{font:{size:13}}},datalabels:{display:false}},scales:{x:{ticks:{font:{size:12}}},y:{ticks:{font:{size:12}}}}}},1000,Math.min(530,Math.max(305,cats.length*80)));
  }
  return _chartPNG({type:'bar',data:{labels:['Faltas (R$)','Sobras (R$)'],datasets:[{data:[Math.abs(c.totalFaltas),c.totalSobras],backgroundColor:['#D32F2F','#F57C00'],borderRadius:4}]},options:{plugins:{legend:{display:false},datalabels:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:12}}},x:{ticks:{font:{size:13}}}}}},900,280);
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
function chartPerda(p){
  return _chartPNG({type:'bar',data:{labels:['Curva A','Curva B','Curva C'],datasets:[{label:'Perda mensal (R$)',data:[p.classA.perda*30,p.classB.perda*30,p.classC.perda*30],backgroundColor:['#D32F2F','#F57C00','#888888'],borderRadius:4}]},options:{plugins:{legend:{display:false},datalabels:{display:false}},scales:{y:{beginAtZero:true,ticks:{font:{size:12}}},x:{ticks:{font:{size:13}}}}}},1000,330);
}
/* ===== r97: gráficos por categoria (foco financeiro), usados no Resumo Executivo ===== */
function chartRupturaCat(r){
  if(!(r.hasCategorias&&r.categorias.length))return null;
  var cats=r.categorias;
  return _chartPNG({type:'bar',data:{labels:cats.map(function(x){return x.nome;}),datasets:[
    {label:'Perda fat./dia (R$)',data:cats.map(function(x){return x.perdaDia;}),backgroundColor:'#D32F2F',borderRadius:3},
    {label:'Perda lucro/dia (R$)',data:cats.map(function(x){return x.perdaLucroDia;}),backgroundColor:'#F57C00',borderRadius:3}
  ]},options:{indexAxis:'y',plugins:{legend:{position:'top',labels:{font:{size:13}}},datalabels:{display:false}},scales:{x:{ticks:{font:{size:12}}},y:{ticks:{font:{size:12}}}}}},1000,Math.min(530,Math.max(305,cats.length*80)));
}
function chartDiasCat(d){
  if(!(d.hasCategorias&&d.categorias.length))return null;
  var cats=d.categorias;
  return _chartPNG({type:'bar',data:{labels:cats.map(function(x){return x.nome;}),datasets:[
    {label:'Sem giro (R$)',data:cats.map(function(x){return x.valorSemGiro;}),backgroundColor:'#888888',borderRadius:3},
    {label:'Excesso (R$)',data:cats.map(function(x){return x.valorExcesso;}),backgroundColor:'#1565C0',borderRadius:3}
  ]},options:{indexAxis:'y',plugins:{legend:{position:'top',labels:{font:{size:13}}},datalabels:{display:false}},scales:{x:{ticks:{font:{size:12}}},y:{ticks:{font:{size:12}}}}}},1000,Math.min(530,Math.max(305,cats.length*80)));
}
function chartAbcCat(a){
  if(!(a.hasCategorias&&a.categorias.length))return null;
  var cats=a.categorias;
  return _chartPNG({type:'bar',data:{labels:cats.map(function(x){return x.nome;}),datasets:[
    {label:'Investimento (R$)',data:cats.map(function(x){return x.investimento;}),backgroundColor:'#002B50',borderRadius:3},
    {label:'Faturamento (R$)',data:cats.map(function(x){return x.faturamento;}),backgroundColor:'#61CF00',borderRadius:3}
  ]},options:{indexAxis:'y',plugins:{legend:{position:'top',labels:{font:{size:13}}},datalabels:{display:false}},scales:{x:{ticks:{font:{size:12}}},y:{ticks:{font:{size:12}}}}}},1000,Math.min(530,Math.max(305,cats.length*80)));
}
function chartPerdaCat(p){
  if(!(p.hasCategorias&&p.categorias.length))return null;
  var cats=p.categorias;
  return _chartPNG({type:'bar',data:{labels:cats.map(function(x){return x.nome;}),datasets:[
    {label:'Perda fat./dia (R$)',data:cats.map(function(x){return x.perdaFatDia;}),backgroundColor:'#D32F2F',borderRadius:3},
    {label:'Perda lucro/dia (R$)',data:cats.map(function(x){return x.perdaLucroDia;}),backgroundColor:'#F57C00',borderRadius:3}
  ]},options:{indexAxis:'y',plugins:{legend:{position:'top',labels:{font:{size:13}}},datalabels:{display:false}},scales:{x:{ticks:{font:{size:12}}},y:{ticks:{font:{size:12}}}}}},1000,Math.min(530,Math.max(305,cats.length*80)));
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
  if(data.critica){var c=data.critica;R=addST(ws,R,'CRÍTICA DO INVENTÁRIO');R=addKR(ws,R,['ACURACIDADE','VALOR ESTOQUE','VALOR ESTOQUE CONTADO','VALOR DAS FALTAS','VALOR DAS SOBRAS','SALDO LÍQUIDO','PERDA DE ESTOQUE (%)'],[PCT(c.acuracidade),BRLi(c.valorEstoque),BRLi(c.valorEstoqueContado),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido),PCT(c.perdaEstoquePct)],[C.green,C.text,C.text,C.red,C.amb,C.red,c.perdaEstoquePct<0?C.red:C.green]);if(c.hasCategorias)R=addDT(ws,R,['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,x.total,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{0:'left',1:'right',2:'right',3:'right',4:'right',5:'right'});}
  if(data.ruptura){var r=data.ruptura;R=addST(ws,R,'RUPTURA LOJA X DEPÓSITO');R=addKR(ws,R,['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CURVA A (FAT.)','RUPTURA CURVA A (LUCRO)'],[PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],[C.red,C.text,C.red,C.red]);}
  if(data.dias){var d=data.dias,fv=fxV(d.items);R=addST(ws,R,'DIAS DE ESTOQUE');R=addKR(ws,R,['COBERTURA GERAL','CURVA A','CURVA B','CURVA C','SEM GIRO'],[d.coberturaGeral+' dias',d.coberturaA+' dias',d.coberturaB+' dias',d.coberturaC+' dias',NUM(d.semGiro)],[C.text,C.text,C.text,C.text,C.red]);var fo=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];R=addDT(ws,R,['Faixa','SKUs','% SKUs','Valor Estoque (R$)','% do Valor'],fo.map(function(f){var cn=d.items.filter(function(i){return i.faixa===f;}).length;var vl=fv.f[f]||0;return[f,cn,PCT(d.total?cn/d.total*100:0),BRLi(vl),PCT(fv.t?vl/fv.t*100:0)];}),{0:'left',1:'right',2:'right',3:'right',4:'right'});}
  if(data.abc){var a=data.abc;R=addST(ws,R,'INVESTIMENTO ABC');R=addKR(ws,R,['VALOR TOTAL EM ESTOQUE','FATURAMENTO 90D','LUCRO 90D','SKUS'],[BRLi(a.totalInvest),BRLi(a.totalFat),BRLi(a.totalLucro),NUM(a.items.length)],[C.text,C.green,C.green,C.text]);R=addDT(ws,R,['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],{0:'center',1:'right',2:'right',3:'right',4:'right'});}
  if(data.perda){var pe=data.perda;R=addST(ws,R,'PROJEÇÃO DE PERDA');R=addKR(ws,R,['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS'],[BRLi(pe.totalPerdaFat),BRLi(pe.totalPerdaLucro),BRLi(pe.perdaMensal),NUM(pe.totalSKUs)],[C.red,C.red,C.red,C.text]);R=addDT(ws,R,['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',pe.classA.count,BRLi(pe.classA.perda),BRLi(pe.classA.lucro),PCT(pe.classA.pct),BRLi(pe.classA.perda*30)],['B',pe.classB.count,BRLi(pe.classB.perda),BRLi(pe.classB.lucro),PCT(pe.classB.pct),BRLi(pe.classB.perda*30)],['C',pe.classC.count,BRLi(pe.classC.perda),BRLi(pe.classC.lucro),PCT(pe.classC.pct),BRLi(pe.classC.perda*30)]],{0:'center',1:'right',2:'right',3:'right',4:'right',5:'right'});}
  ws['!cols']=[{wch:28},{wch:18},{wch:16},{wch:18},{wch:16},{wch:18},{wch:16},{wch:16}];ws['!rows']=[{hpt:28},{hpt:20}];
  XLSX.utils.book_append_sheet(wb,ws,'Dashboard');

  /* CRITICA RESUMO + TOP20 */
  if(sel.criticaResumo&&data.critica){var wsC={},rw=0,c=data.critica;rw=addBH(wsC,rw,info,pd,6);rw=addST(wsC,rw,'RESUMO DA CRÍTICA');var sL=['ACURACIDADE','SKUs analisados','SKUs sem divergência','SKUs com falta','SKUs com sobra','Valor estoque (sistema)','Valor estoque contado','Valor das faltas','Valor das sobras','Saldo líquido','Perda de estoque (%)'],sV=[PCT(c.acuracidade),c.totalSKUs,c.okCount,c.faltaCount,c.sobraCount,BRLi(c.valorEstoque),BRLi(c.valorEstoqueContado),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido),PCT(c.perdaEstoquePct)];for(var i=0;i<sL.length;i++){sC(wsC,rw+i,0,sL[i],sB('left',true));sC(wsC,rw+i,1,sV[i],sB('right'));}rw+=sL.length+1;
  if(c.hasCategorias){rw=addST(wsC,rw,'RESULTADO POR CATEGORIA');rw=addDT(wsC,rw,['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,x.total,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{0:'left',1:'right',2:'right',3:'right',4:'right',5:'right'});}
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
  function hdr(){doc.setFillColor(5,19,35);doc.rect(0,0,W,22,'F');if(logo){try{doc.addImage(logo,'PNG',M,7,32,8);}catch(e){}}doc.setFontSize(9);doc.setTextColor(255,255,255);doc.text((info.cliente||'')+' — '+(info.unidade||''),W-M,7,{align:'right'});doc.setFontSize(7);doc.setTextColor(200,220,255);doc.text('Inventário: '+(info.dataInventario||'—'),W-M,12,{align:'right'});doc.setTextColor(180,180,200);doc.text('Processado em '+pd,W-M,17,{align:'right'});y=28;}
  function ftr(pg){doc.setFontSize(7);doc.setTextColor(150,150,150);doc.text('Formula Code Tecnologia, Gestão e Automação',M,H-6);doc.text('Página '+pg,W-M,H-6,{align:'right'});doc.setDrawColor(200,200,200);doc.line(M,H-10,W-M,H-10);}
  function chk(n){if(y+n>H-18){doc.addPage();hdr();ftr(doc.getNumberOfPages());}}
  function ttl(t){chk(12);doc.setFontSize(14);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFontSize(8);doc.setTextColor(150,150,150);doc.setFont(undefined,'normal');doc.text('Relatório gerado automaticamente pelo sistema Formula Code',M,y);y+=8;}
  function sec(t){chk(10);doc.setFontSize(11);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFont(undefined,'normal');}
  function aT(h,b,o){chk(20);doc.autoTable({startY:y,head:[h],body:b,margin:{left:M,right:M},headStyles:{fillColor:[5,19,35],fontSize:7,fontStyle:'bold',halign:'left'},bodyStyles:{fontSize:7,halign:'left'},alternateRowStyles:{fillColor:[245,245,245]},styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},columnStyles:o||{}});y=doc.lastAutoTable.finalY+6;}
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
    if(c.hasCategorias){sec('Resultado por categoria');aT(['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,x.total,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});}
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
        aT(['SKU','Descrição','Qtd Depósito'],itens.map(function(i){return[i.sku,i.descricao,i.deposito];}),{2:{halign:'right'}});
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
    aT(['Faixa','SKUs','% SKUs','Valor Estoque (R$)','% do Valor'],fo.map(function(f){var cn=d.items.filter(function(i){return i.faixa===f;}).length;var vl=fv.f[f]||0;return[f,cn,PCT(d.total?cn/d.total*100:0),BRLi(vl),PCT(fv.t?vl/fv.t*100:0)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
    if(d.hasCategorias){sec('Cobertura por categoria');aT(['Categoria','SKUs','Cobertura média','Rupt+Alto risco','Sem giro','Excessos','Val. estoque'],d.categorias.map(function(x){return[x.nome,x.total,x.mediaCobertura+' dias',x.criticos,x.semGiro,x.excessos,BRLi(x.valorEstoque)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'},6:{halign:'right'}});}
    var pH=['SKU','Descrição','Categoria','Dias est.','ABC Fat.','Val. estoque'],pO={3:{halign:'right'},5:{halign:'right'}};
    fo.forEach(function(fx){var it=d.items.filter(function(i){return i.faixa===fx;});if(it.length){sec(fx+' — '+it.length+' itens');aT(pH,it.slice(0,50).map(function(i){return[i.sku,i.descricao,i.categoria||'',i.diasEstoque!==null?R2(i.diasEstoque):'—',i.abcFat,BRL(i.valorEstoque)];}),pO);if(it.length>50){doc.setFontSize(7);doc.setTextColor(150,150,150);doc.text('... e mais '+(it.length-50)+' itens (ver Excel)',M,y);y+=4;}}});
  }
  else if(rt==='abc'){
    var a=data.abc;ttl('Investimento por curva ABC — Resumo executivo');
    var iaABC=window._iaResumos&&window._iaResumos.abc;
    sec('Análise');bloco(iaABC||Engine.gerarAnaliseABC(a,info));sec('Metodologia');bloco(metABC(info.diasVenda));
    sec('Indicadores gerais');kpi(['VALOR TOTAL EM ESTOQUE','FATURAMENTO 90D','LUCRO 90D','SKUS ANALISADOS'],[BRLi(a.totalInvest),BRLi(a.totalFat),BRLi(a.totalLucro),NUM(a.items.length)],[[51,51,51],[0,183,74],[0,183,74],[51,51,51]]);
    sec('Gráfico — Valor em estoque × Faturamento por curva');img(chartABC(a),60);
    sec('Curva ABC por faturamento');aT(['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
    sec('Curva ABC por lucro');aT(['Curva','Valor Estoque (R$)','% Estoque','Lucro (R$)','% Lucro'],[['A',BRLi(a.lucA.invest),PCT(a.lucA.pctInvest),BRLi(a.lucA.luc),PCT(a.lucA.pctLuc)],['B',BRLi(a.lucB.invest),PCT(a.lucB.pctInvest),BRLi(a.lucB.luc),PCT(a.lucB.pctLuc)],['C',BRLi(a.lucC.invest),PCT(a.lucC.pctInvest),BRLi(a.lucC.luc),PCT(a.lucC.pctLuc)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
  }
  else if(rt==='perda'){
    var p=data.perda;ttl('Projeção de venda perdida — Resumo executivo');
    var iaPerda=window._iaResumos&&window._iaResumos.perda;
    sec('Análise');bloco(iaPerda||Engine.gerarAnalisePerda(p,info));sec('Metodologia');bloco(metPerda(info.diasVenda));
    sec('Indicadores gerais');kpi(['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS'],[BRLi(p.totalPerdaFat),BRLi(p.totalPerdaLucro),BRLi(p.perdaMensal),NUM(p.totalSKUs)],[[211,47,47],[211,47,47],[211,47,47],[51,51,51]]);
    sec('Impacto por curva ABC');aT(['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',p.classA.count,BRLi(p.classA.perda),BRLi(p.classA.lucro),PCT(p.classA.pct),BRLi(p.classA.perda*30)],['B',p.classB.count,BRLi(p.classB.perda),BRLi(p.classB.lucro),PCT(p.classB.pct),BRLi(p.classB.perda*30)],['C',p.classC.count,BRLi(p.classC.perda),BRLi(p.classC.lucro),PCT(p.classC.pct),BRLi(p.classC.perda*30)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});
    sec('Gráfico — Perda mensal projetada por curva');img(chartPerda(p),60);
    var pH2=['SKU','Descrição','Categoria','Perda Fat./Mês','Perda Lucro/Mês'],pO2={3:{halign:'right'},4:{halign:'right'}};
    ['A','B','C'].forEach(function(cls){var it=p.items.filter(function(i){return i.abcFat===cls;}).sort(function(a,b){return b.perdaFatMes-a.perdaFatMes;});if(it.length){sec('Curva '+cls+' — '+it.length+' itens');aT(pH2,it.map(function(i){return[i.sku,i.descricao,i.categoria||'',BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];}),pO2);}});
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

return{generateExcel:generateExcel,generateExcelBlob:generateExcelBlob,buildExcelWorkbook:buildExcelWorkbook,generatePDF:generatePDF,generateComparativoPDF:generateComparativoPDF,generateComparativoExcel:generateComparativoExcel,generateResumoPDF:generateResumoPDF,generateResumoPPTX:generateResumoPPTX};

/* ===== r68: RESUMO EXECUTIVO — PDF Documento ===== */
function generateResumoPDF(results,recs,info,unidade,logo){
  var jsPDF=window.jspdf.jsPDF;
  var doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  var W=210,H=297,M=20,cw=W-2*M;
  var y=M;
  var navy=[0,43,80],green=[97,207,0],dark=[5,19,35];

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
  if(results.critica)dims.push({label:'Acuracidade',valor:results.critica.acuracidade+'%',detalhe:'Estoque: '+BRLi(results.critica.valorEstoque)+' · Contado: '+BRLi(results.critica.valorEstoqueContado)+' · Divergência: '+BRLi(results.critica.saldoLiquido),cor:results.critica.acuracidade<90?red:(results.critica.acuracidade<95?amb:green)});
  if(results.ruptura)dims.push({label:'Ruptura',valor:results.ruptura.taxaRuptura+'%',detalhe:'R$/dia: '+BRLi(results.ruptura.perdaFatDia)+' · R$/mês: '+BRLi(results.ruptura.perdaMensal)+' · Estoque parado: '+BRLi(results.ruptura.valorEstoqueParado),cor:results.ruptura.taxaRuptura>10?red:(results.ruptura.taxaRuptura>5?amb:green)});
  if(results.dias)dims.push({label:'Cobertura',valor:Engine.round2(results.dias.coberturaGeral)+' dias',detalhe:'Sem giro: '+BRLi(results.dias.valorSemGiro)+' · Excesso: '+BRLi(results.dias.valorExcesso),cor:(results.dias.coberturaGeral<15||results.dias.coberturaGeral>60)?red:green});
  if(results.abc)dims.push({label:'Investimento',valor:'R$ '+Engine.formatNum(results.abc.totalInvest),detalhe:'Faturamento: R$ '+Engine.formatNum(results.abc.totalFat)+' · Giro: '+Engine.round2(results.abc.totalFat/results.abc.totalInvest)+'x',cor:navy});
  if(results.perda)dims.push({label:'Perda Mensal',valor:'R$ '+Engine.formatNum(results.perda.perdaMensal),detalhe:'Diária: '+BRLi(results.perda.totalPerdaFat)+' · Mensal: '+BRLi(results.perda.perdaMensal),cor:results.perda.perdaMensal>50000?red:(results.perda.perdaMensal>10000?amb:green)});

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

  /* Detalhamento por categoria (foco financeiro em todas as dimensões) */
  function ckR(n){if(y+n>270){doc.addPage();y=M;}}
  function imgR(url,h){if(!url)return;ckR(h+10);try{doc.addImage(url,'PNG',M,y,cw,h);}catch(e){}y+=h+10;}
  function tituloR(t){ckR(14);doc.setFontSize(11);doc.setTextColor.apply(doc,navy);doc.text(t,M,y);y+=6;}
  function tabR(h,b,o){ckR(20);doc.autoTable({startY:y,head:[h],body:b,margin:{left:M,right:M},headStyles:{fillColor:navy,fontSize:7,fontStyle:'bold',halign:'left'},bodyStyles:{fontSize:7,halign:'left'},alternateRowStyles:{fillColor:[245,245,245]},styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},columnStyles:o||{}});y=doc.lastAutoTable.finalY+6;}

  if(y>245){doc.addPage();y=M;}
  doc.setFontSize(18);doc.setTextColor.apply(doc,navy);
  doc.text('Detalhamento por categoria',M,y);y+=10;

  if(results.critica){
    var cc=results.critica;
    tituloR('Acuracidade'+(cc.hasCategorias?' por categoria':''));
    if(cc.hasCategorias){tabR(['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],cc.categorias.map(function(x){return[x.nome,x.total,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});}
    imgR(chartCritica(cc),cc.hasCategorias?Math.min(80,Math.max(50,cc.categorias.length*12)):50);
  }
  if(results.ruptura){
    var rr=results.ruptura;
    tituloR('Ruptura'+(rr.hasCategorias?' por categoria':' por curva ABC'));
    if(rr.hasCategorias){
      tabR(['Categoria','Rupturas','Taxa ruptura','Rupturas A','Perda fat./dia','Perda lucro/dia','Perda mensal','Estoque parado'],rr.categorias.map(function(x){return[x.nome,x.totalRupturas,PCT(x.taxa),x.rupturaA,BRLi(x.perdaDia),BRLi(x.perdaLucroDia),BRLi(x.perdaMensal),BRLi(x.valorEstoqueParado)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'},6:{halign:'right'},7:{halign:'right'}});
      imgR(chartRupturaCat(rr),Math.min(80,Math.max(50,rr.categorias.length*12)));
    }else{
      imgR(chartRuptura(rr),55);
    }
  }
  if(results.dias){
    var dd=results.dias;
    tituloR('Cobertura'+(dd.hasCategorias?' por categoria':' por faixa'));
    if(dd.hasCategorias){
      tabR(['Categoria','SKUs','Cobertura média','Sem giro (R$)','Excesso (R$)','Val. estoque'],dd.categorias.map(function(x){return[x.nome,x.total,x.mediaCobertura+' dias',BRLi(x.valorSemGiro),BRLi(x.valorExcesso),BRLi(x.valorEstoque)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});
      imgR(chartDiasCat(dd),Math.min(80,Math.max(50,dd.categorias.length*12)));
    }else{
      imgR(chartDias(dd),55);
    }
  }
  if(results.abc){
    var aa=results.abc;
    tituloR('Investimento'+(aa.hasCategorias?' por categoria':' por curva ABC'));
    if(aa.hasCategorias){
      tabR(['Categoria','Investimento (R$)','Faturamento (R$)','Lucro (R$)','% Investimento'],aa.categorias.map(function(x){return[x.nome,BRLi(x.investimento),BRLi(x.faturamento),BRLi(x.lucro),PCT(x.pctInvest)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
      imgR(chartAbcCat(aa),Math.min(80,Math.max(50,aa.categorias.length*12)));
    }else{
      imgR(chartABC(aa),55);
    }
  }
  if(results.perda){
    var pp=results.perda;
    tituloR('Perda projetada'+(pp.hasCategorias?' por categoria':' por curva ABC'));
    if(pp.hasCategorias){
      tabR(['Categoria','Rupturas','Perda fat./dia','Perda lucro/dia','Perda mensal'],pp.categorias.map(function(x){return[x.nome,x.totalRupturas,BRLi(x.perdaFatDia),BRLi(x.perdaLucroDia),BRLi(x.perdaMensal)];}),{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
      imgR(chartPerdaCat(pp),Math.min(80,Math.max(50,pp.categorias.length*12)));
    }else{
      imgR(chartPerda(pp),55);
    }
  }

  /* Recomendações */
  if(recs.length>0){
    y+=5;
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
  if(results.critica)dims.push({label:'ACURACIDADE',valor:results.critica.acuracidade+'%',sub:'Divergência: '+BRLi(results.critica.saldoLiquido),accent:results.critica.acuracidade<95});
  if(results.ruptura)dims.push({label:'RUPTURA',valor:results.ruptura.taxaRuptura+'%',sub:'Perda: '+BRLi(results.ruptura.perdaMensal)+'/mês',accent:true});
  if(results.dias)dims.push({label:'COBERTURA',valor:Engine.round2(results.dias.coberturaGeral)+' dias',sub:'Sem giro: '+BRLi(results.dias.valorSemGiro),accent:results.dias.coberturaGeral>45});
  if(results.abc)dims.push({label:'INVESTIMENTO',valor:'R$ '+Engine.formatNum(results.abc.totalInvest),sub:'Giro: '+Engine.round2(results.abc.totalFat/results.abc.totalInvest)+'x',accent:false});
  if(results.perda)dims.push({label:'PERDA MENSAL',valor:'R$ '+Engine.formatNum(results.perda.perdaMensal),sub:'Diária: '+BRLi(results.perda.totalPerdaFat)+' · Mensal: '+BRLi(results.perda.perdaMensal),accent:true});

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
  if(results.ruptura){
    var rr=results.ruptura;
    if(rr.hasCategorias&&rr.categorias.length){
      chartDims.push({title:'Ruptura — Perda por categoria (R$/dia)',data:[{name:'Perda fat./dia',labels:rr.categorias.map(function(x){return x.nome;}),values:rr.categorias.map(function(x){return x.perdaDia;})},{name:'Perda lucro/dia',labels:rr.categorias.map(function(x){return x.nome;}),values:rr.categorias.map(function(x){return x.perdaLucroDia;})}],colors:['D32F2F','F57C00'],legend:true});
    }else{
      var perdaA=rr.items.filter(function(i){return i.abc_valorVendido90==='A';}).reduce(function(s,i){return s+(i.fatMediaDia||0);},0);
      var perdaB=rr.items.filter(function(i){return i.abc_valorVendido90==='B';}).reduce(function(s,i){return s+(i.fatMediaDia||0);},0);
      var perdaC=rr.items.filter(function(i){return i.abc_valorVendido90==='C';}).reduce(function(s,i){return s+(i.fatMediaDia||0);},0);
      chartDims.push({title:'Ruptura — Perda por curva ABC (R$/dia)',data:[{name:'Perda fat./dia',labels:['Curva A','Curva B','Curva C'],values:[Engine.round2(perdaA),Engine.round2(perdaB),Engine.round2(perdaC)]}],colors:['D32F2F'],legend:false});
    }
  }
  if(results.dias){
    var dd=results.dias;
    if(dd.hasCategorias&&dd.categorias.length){
      chartDims.push({title:'Cobertura — Valor por categoria (R$)',data:[{name:'Sem giro (R$)',labels:dd.categorias.map(function(x){return x.nome;}),values:dd.categorias.map(function(x){return x.valorSemGiro;})},{name:'Excesso (R$)',labels:dd.categorias.map(function(x){return x.nome;}),values:dd.categorias.map(function(x){return x.valorExcesso;})}],colors:['888888','1565C0'],legend:true});
    }else{
      var foD=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];
      var foVals=foD.map(function(f){return Engine.round2(dd.items.filter(function(i){return i.faixa===f;}).reduce(function(s,i){return s+(i.valorEstoque||0);},0));});
      chartDims.push({title:'Cobertura — Valor por faixa (R$)',data:[{name:'Valor estoque (R$)',labels:foD,values:foVals}],colors:['1565C0'],legend:false});
    }
  }
  if(results.abc){
    var aa=results.abc;
    if(aa.hasCategorias&&aa.categorias.length){
      chartDims.push({title:'Investimento por categoria (R$)',data:[{name:'Investimento (R$)',labels:aa.categorias.map(function(x){return x.nome;}),values:aa.categorias.map(function(x){return x.investimento;})},{name:'Faturamento (R$)',labels:aa.categorias.map(function(x){return x.nome;}),values:aa.categorias.map(function(x){return x.faturamento;})}],colors:['002B50','61CF00'],legend:true});
    }else{
      chartDims.push({title:'Investimento por curva ABC (R$)',data:[{name:'Valor estoque (R$)',labels:['Curva A','Curva B','Curva C'],values:[aa.fatA.invest,aa.fatB.invest,aa.fatC.invest]}],colors:['002B50'],legend:false});
    }
  }
  if(results.perda){
    var pp=results.perda;
    if(pp.hasCategorias&&pp.categorias.length){
      chartDims.push({title:'Perda mensal por categoria (R$)',data:[{name:'Perda fat./dia',labels:pp.categorias.map(function(x){return x.nome;}),values:pp.categorias.map(function(x){return x.perdaFatDia;})},{name:'Perda lucro/dia',labels:pp.categorias.map(function(x){return x.nome;}),values:pp.categorias.map(function(x){return x.perdaLucroDia;})}],colors:['D32F2F','F57C00'],legend:true});
    }else{
      chartDims.push({title:'Perda mensal por curva ABC',data:[{name:'R$/mês',labels:['Curva A','Curva B','Curva C'],values:[pp.classA.perda*30,pp.classB.perda*30,pp.classC.perda*30]}],colors:['D32F2F'],legend:false});
    }
  }

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

})();
