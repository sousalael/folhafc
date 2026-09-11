/* export.js — r103 — mantém tudo do r102 + Comparativo (PDF/HTML): remove os cards de Sobreposição de SKUs, adiciona faixa de cards por unidade (Acuracidade/Valor Quebra/Cobertura de Estoque/SKUs em Ruptura/Venda perdida por dia) e usa o mesmo texto da Análise Comparativa exibido na tela */
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
/* ===== r102: MINI-BARRAS DESENHADAS DIRETO NO jsPDF (substituem os gráficos Chart.js/PNG) ===== */
/* Desenha uma mini-barra dupla (2 métricas lado a lado) por linha — usada em Crítica (Faltas/Sobras) e ABC (Estoque/Faturamento).
   doc: instância jsPDF; x,y: canto superior esquerdo; w: largura total disponível (respeita as margens do caller);
   rows: [{label:'...', a:valor, b:valor}]; retorna a altura (mm) efetivamente usada, para o caller avançar o cursor y. */
function _pdfMinibarDual(doc,x,y,w,rows,labelA,labelB,colorA,colorB,fmtA,fmtB){
  var labelW=Math.min(46,w*0.30), valW=26, barW=Math.max(20,w-labelW-valW-4);
  var barH=3.4, gapBars=1.4, gapRows=3.4;
  var maxVal=0; rows.forEach(function(r){maxVal=Math.max(maxVal,Math.abs(r.a||0),Math.abs(r.b||0));});
  if(maxVal<=0)maxVal=1;
  doc.setFontSize(7);doc.setTextColor(80,80,80);doc.text(labelA,x+labelW,y-1.5);
  doc.setTextColor(colorA[0],colorA[1],colorA[2]);
  var legendW=doc.getTextWidth(labelA)+6;
  doc.setTextColor(colorB[0],colorB[1],colorB[2]);doc.text(labelB,x+labelW+legendW,y-1.5);
  var yy=y+2;
  rows.forEach(function(r){
    doc.setFontSize(7.2);doc.setFont(undefined,'bold');doc.setTextColor(51,51,51);
    doc.text(String(r.label).substring(0,30),x,yy+barH+0.6);
    doc.setFont(undefined,'normal');
    var wA=Math.max(0.6,Math.abs(r.a||0)/maxVal*barW);
    doc.setFillColor(colorA[0],colorA[1],colorA[2]);doc.roundedRect(x+labelW,yy,wA,barH,0.6,0.6,'F');
    doc.setFontSize(6.6);doc.setTextColor(colorA[0],colorA[1],colorA[2]);
    doc.text(fmtA(r.a),x+labelW+barW+3,yy+barH-0.3);
    yy+=barH+gapBars;
    var wB=Math.max(0.6,Math.abs(r.b||0)/maxVal*barW);
    doc.setFillColor(colorB[0],colorB[1],colorB[2]);doc.roundedRect(x+labelW,yy,wB,barH,0.6,0.6,'F');
    doc.setFontSize(6.6);doc.setTextColor(colorB[0],colorB[1],colorB[2]);
    doc.text(fmtB(r.b),x+labelW+barW+3,yy+barH-0.3);
    yy+=barH+gapRows;
  });
  return (yy-y);
}
/* Mini-barra de série única — usada em Ruptura (SKUs por curva) e Perda (perda mensal por curva/categoria).
   rows: [{label, valor, color:[r,g,b]}]; fmt(valor) formata o rótulo da direita. */
function _pdfMinibarSingle(doc,x,y,w,rows,fmt){
  var labelW=Math.min(46,w*0.30), valW=28, barW=Math.max(20,w-labelW-valW-4);
  var barH=4.2, gapRows=3;
  var maxVal=0; rows.forEach(function(r){maxVal=Math.max(maxVal,Math.abs(r.valor||0));});
  if(maxVal<=0)maxVal=1;
  var yy=y;
  rows.forEach(function(r){
    doc.setFontSize(7.4);doc.setFont(undefined,'bold');doc.setTextColor(51,51,51);
    doc.text(String(r.label).substring(0,30),x,yy+barH-0.8);
    doc.setFont(undefined,'normal');
    var wv=Math.max(0.6,Math.abs(r.valor||0)/maxVal*barW);
    var c=r.color||[211,47,47];
    doc.setFillColor(c[0],c[1],c[2]);doc.roundedRect(x+labelW,yy,wv,barH,0.7,0.7,'F');
    doc.setFontSize(7);doc.setTextColor(c[0],c[1],c[2]);
    doc.text(fmt(r.valor),x+labelW+barW+3,yy+barH-1);
    yy+=barH+gapRows;
  });
  return (yy-y);
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
  if(data.critica){var c=data.critica;R=addST(ws,R,'CRÍTICA DO INVENTÁRIO');R=addKR(ws,R,['ACURACIDADE','VALOR ESTOQUE','VALOR ESTOQUE CONTADO','VALOR DAS FALTAS','VALOR DAS SOBRAS','VALOR QUEBRA','PERDA DE ESTOQUE (%)'],[PCT(c.acuracidade),BRLi(c.valorEstoque),BRLi(c.valorEstoqueContado),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido),PCT(c.perdaEstoquePct)],[C.green,C.text,C.text,C.red,C.amb,C.red,c.perdaEstoquePct<0?C.red:C.green]);if(c.hasCategorias)R=addDT(ws,R,['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,x.total,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{0:'left',1:'right',2:'right',3:'right',4:'right',5:'right'});}
  if(data.ruptura){var r=data.ruptura;R=addST(ws,R,'RUPTURA LOJA X DEPÓSITO');R=addKR(ws,R,['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CURVA A (FAT.)','RUPTURA CURVA A (LUCRO)'],[PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],[C.red,C.text,C.red,C.red]);}
  if(data.dias){var d=data.dias,fv=fxV(d.items);R=addST(ws,R,'DIAS DE ESTOQUE');R=addKR(ws,R,['COBERTURA GERAL','CURVA A','CURVA B','CURVA C','SEM GIRO'],[d.coberturaGeral+' dias',d.coberturaA+' dias',d.coberturaB+' dias',d.coberturaC+' dias',NUM(d.semGiro)],[C.text,C.text,C.text,C.text,C.red]);var fo=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];R=addDT(ws,R,['Faixa','SKUs','% SKUs','Valor Estoque (R$)','% do Valor'],fo.map(function(f){var cn=d.items.filter(function(i){return i.faixa===f;}).length;var vl=fv.f[f]||0;return[f,cn,PCT(d.total?cn/d.total*100:0),BRLi(vl),PCT(fv.t?vl/fv.t*100:0)];}),{0:'left',1:'right',2:'right',3:'right',4:'right'});}
  if(data.abc){var a=data.abc;R=addST(ws,R,'INVESTIMENTO ABC');R=addKR(ws,R,['VALOR TOTAL EM ESTOQUE','FATURAMENTO 90D','LUCRO 90D','SKUS'],[BRLi(a.totalInvest),BRLi(a.totalFat),BRLi(a.totalLucro),NUM(a.items.length)],[C.text,C.green,C.green,C.text]);R=addDT(ws,R,['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],{0:'center',1:'right',2:'right',3:'right',4:'right'});}
  if(data.perda){var pe=data.perda;R=addST(ws,R,'PROJEÇÃO DE PERDA');R=addKR(ws,R,['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS'],[BRLi(pe.totalPerdaFat),BRLi(pe.totalPerdaLucro),BRLi(pe.perdaMensal),NUM(pe.totalSKUs)],[C.red,C.red,C.red,C.text]);R=addDT(ws,R,['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',pe.classA.count,BRLi(pe.classA.perda),BRLi(pe.classA.lucro),PCT(pe.classA.pct),BRLi(pe.classA.perda*30)],['B',pe.classB.count,BRLi(pe.classB.perda),BRLi(pe.classB.lucro),PCT(pe.classB.pct),BRLi(pe.classB.perda*30)],['C',pe.classC.count,BRLi(pe.classC.perda),BRLi(pe.classC.lucro),PCT(pe.classC.pct),BRLi(pe.classC.perda*30)]],{0:'center',1:'right',2:'right',3:'right',4:'right',5:'right'});}
  ws['!cols']=[{wch:28},{wch:18},{wch:16},{wch:18},{wch:16},{wch:18},{wch:16},{wch:16}];ws['!rows']=[{hpt:28},{hpt:20}];
  XLSX.utils.book_append_sheet(wb,ws,'Dashboard');

  /* CRITICA RESUMO + TOP20 */
  if(sel.criticaResumo&&data.critica){var wsC={},rw=0,c=data.critica;rw=addBH(wsC,rw,info,pd,6);rw=addST(wsC,rw,'RESUMO DA CRÍTICA');var sL=['ACURACIDADE','SKUs analisados','SKUs sem divergência','SKUs com falta','SKUs com sobra','Valor estoque (sistema)','Valor estoque contado','Valor das faltas','Valor das sobras','Valor Quebra','Perda de estoque (%)'],sV=[PCT(c.acuracidade),c.totalSKUs,c.okCount,c.faltaCount,c.sobraCount,BRLi(c.valorEstoque),BRLi(c.valorEstoqueContado),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido),PCT(c.perdaEstoquePct)];for(var i=0;i<sL.length;i++){sC(wsC,rw+i,0,sL[i],sB('left',true));sC(wsC,rw+i,1,sV[i],sB('right'));}rw+=sL.length+1;
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
    kpi(['VALOR DAS FALTAS','VALOR DAS SOBRAS','VALOR QUEBRA'],[BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido)],[[211,47,47],[245,124,0],[211,47,47]]);
    sec('Gráfico — Faltas e sobras'+(c.hasCategorias?' por categoria':''));
    var rowsCrit=c.hasCategorias?c.categorias.map(function(x){return{label:x.nome,a:Math.abs(x.faltaVal),b:x.sobraVal};}):[{label:'Total',a:Math.abs(c.totalFaltas),b:c.totalSobras}];
    chk(rowsCrit.length*9+8);y+=4;y+=_pdfMinibarDual(doc,M,y,W-2*M,rowsCrit,'Faltas (R$)','Sobras (R$)',[211,47,47],[245,124,0],BRLi,BRLi)+6;
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
    sec('Gráfico — Rupturas por curva ABC');
    var rowsRup=[{label:'Curva A',valor:r.rupturaA,color:[211,47,47]},{label:'Curva B',valor:r.rupturaB,color:[245,124,0]},{label:'Curva C',valor:r.rupturaC,color:[136,136,136]}];
    chk(rowsRup.length*8+8);y+=4;y+=_pdfMinibarSingle(doc,M,y,W-2*M,rowsRup,NUM)+6;
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
    sec('Gráfico — Valor em estoque × Faturamento por curva');
    var rowsABC=[{label:'Curva A',a:a.fatA.invest,b:a.fatA.fat},{label:'Curva B',a:a.fatB.invest,b:a.fatB.fat},{label:'Curva C',a:a.fatC.invest,b:a.fatC.fat}];
    chk(rowsABC.length*9+8);y+=4;y+=_pdfMinibarDual(doc,M,y,W-2*M,rowsABC,'Valor em estoque (R$)','Faturamento (R$)',[0,43,80],[97,207,0],BRLi,BRLi)+6;
    sec('Curva ABC por faturamento');aT(['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
    sec('Curva ABC por lucro');aT(['Curva','Valor Estoque (R$)','% Estoque','Lucro (R$)','% Lucro'],[['A',BRLi(a.lucA.invest),PCT(a.lucA.pctInvest),BRLi(a.lucA.luc),PCT(a.lucA.pctLuc)],['B',BRLi(a.lucB.invest),PCT(a.lucB.pctInvest),BRLi(a.lucB.luc),PCT(a.lucB.pctLuc)],['C',BRLi(a.lucC.invest),PCT(a.lucC.pctInvest),BRLi(a.lucC.luc),PCT(a.lucC.pctLuc)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
  }
  else if(rt==='perda'){
    var p=data.perda;ttl('Projeção de venda perdida — Resumo executivo');
    var iaPerda=window._iaResumos&&window._iaResumos.perda;
    sec('Análise');bloco(iaPerda||Engine.gerarAnalisePerda(p,info));sec('Metodologia');bloco(metPerda(info.diasVenda));
    sec('Indicadores gerais');kpi(['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS'],[BRLi(p.totalPerdaFat),BRLi(p.totalPerdaLucro),BRLi(p.perdaMensal),NUM(p.totalSKUs)],[[211,47,47],[211,47,47],[211,47,47],[51,51,51]]);
    sec('Impacto por curva ABC');aT(['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',p.classA.count,BRLi(p.classA.perda),BRLi(p.classA.lucro),PCT(p.classA.pct),BRLi(p.classA.perda*30)],['B',p.classB.count,BRLi(p.classB.perda),BRLi(p.classB.lucro),PCT(p.classB.pct),BRLi(p.classB.perda*30)],['C',p.classC.count,BRLi(p.classC.perda),BRLi(p.classC.lucro),PCT(p.classC.pct),BRLi(p.classC.perda*30)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});
    sec('Gráfico — Perda mensal projetada por curva');
    var rowsPerda=[{label:'Curva A',valor:p.classA.perda*30,color:[211,47,47]},{label:'Curva B',valor:p.classB.perda*30,color:[245,124,0]},{label:'Curva C',valor:p.classC.perda*30,color:[136,136,136]}];
    chk(rowsPerda.length*8+8);y+=4;y+=_pdfMinibarSingle(doc,M,y,W-2*M,rowsPerda,BRLi)+6;
    var pH2=['SKU','Descrição','Categoria','Perda Fat./Mês','Perda Lucro/Mês'],pO2={3:{halign:'right'},4:{halign:'right'}};
    ['A','B','C'].forEach(function(cls){var it=p.items.filter(function(i){return i.abcFat===cls;}).sort(function(a,b){return b.perdaFatMes-a.perdaFatMes;});if(it.length){sec('Curva '+cls+' — '+it.length+' itens');aT(pH2,it.map(function(i){return[i.sku,i.descricao,i.categoria||'',BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];}),pO2);}});
  }
  chk(12);doc.setFontSize(7);doc.setTextColor(150,150,150);
  doc.text('Nota: relatório baseado em dados processados em '+pd+'. Valores projetados são estimativas.',M,y);
  doc.save('resumo_'+rt+'_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.pdf');
}
/* ========== PDF COMPARATIVO ========== */
function generateComparativoPDF(comp, units, info, analiseTxt, logo){
  info=info||{};
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
  /* r103: faixa de cards por unidade — mesmas 5 métricas do Comparativo na tela (Acuracidade, Valor Quebra, Cobertura de Estoque, SKUs em Ruptura, Venda perdida/dia) */
  function storeCards(){
    var n=comp.unidades.length;if(!n)return;
    var gap=4,cw2=(W-2*M-(n-1)*gap)/n;
    var rowsDef=[
      {label:'Acuracidade',get:function(u){return u.acuracidade!=null?PCT(u.acuracidade):'—';},cor:function(){return[0,183,74];}},
      {label:'Valor Quebra',get:function(u){return u.saldoLiquido!=null?BRLi(u.saldoLiquido):'—';},cor:function(u){return u.saldoLiquido<0?[211,47,47]:(u.saldoLiquido>0?[0,183,74]:[51,51,51]);}},
      {label:'Cobertura de Estoque',get:function(u){return u.coberturaGeral!=null?(u.coberturaGeral+' dias'):'—';},cor:function(){return[51,51,51];}},
      {label:'SKUs em Ruptura',get:function(u){return u.totalRupturas!=null?NUM(u.totalRupturas):'—';},cor:function(u){return u.totalRupturas>0?[211,47,47]:[136,136,136];}},
      {label:'Venda perdida / dia',get:function(u){return u.perdaFatDia!=null?BRLi(u.perdaFatDia):'—';},cor:function(){return[211,47,47];}}
    ];
    var lineH=6.2,headH=9,padTop=4,ch=headH+padTop+rowsDef.length*lineH+3;
    chk(ch+8);
    var cy=y;
    comp.unidades.forEach(function(u,i){
      var cx=M+i*(cw2+gap);
      doc.setFillColor(250,250,250);doc.setDrawColor(224,224,224);doc.roundedRect(cx,cy,cw2,ch,2,2,'FD');
      doc.setFillColor(0,183,74);doc.rect(cx,cy,cw2,1.2,'F');
      doc.setFontSize(8.5);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');
      doc.text(String(u.unidade),cx+3,cy+7,{maxWidth:cw2-6});
      doc.setFont(undefined,'normal');
      var ry=cy+headH+padTop;
      rowsDef.forEach(function(r){
        doc.setFontSize(6.3);doc.setTextColor(136,136,136);
        doc.text(r.label,cx+3,ry);
        ry+=3.2;
        doc.setFontSize(8);var cc=r.cor(u);doc.setTextColor(cc[0],cc[1],cc[2]);doc.setFont(undefined,'bold');
        doc.text(r.get(u),cx+3,ry);
        doc.setFont(undefined,'normal');
        ry+=lineH-3.2;
      });
    });
    y=cy+ch+8;
  }

  hdr();ftr(1);
  ttl('Comparativo entre unidades — '+comp.unidades.length+' unidades');

  /* r103: Análise Comparativa — mesmo texto exibido na tela (IA ou fallback local) */
  sec('Análise Comparativa');
  bloco(analiseTxt||'Análise comparativa entre '+comp.unidades.length+' unidades do cliente '+(info.cliente||'')+' referente ao inventário de '+(info.dataInventario||'')+'.');

  /* r103: cards por unidade */
  sec('Indicadores por unidade');
  storeCards();

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

  /* ===== r102: Análises (Ruptura + Dias) ===== */
  if(results.ruptura||results.dias){
    if(y>245){doc.addPage();y=M;}
    doc.setFontSize(18);doc.setTextColor.apply(doc,navy);
    doc.text('Análises',M,y);y+=10;
    if(results.ruptura){
      if(y>250){doc.addPage();y=M;}
      doc.setFontSize(11);doc.setTextColor.apply(doc,navy);doc.setFont(undefined,'bold');
      doc.text('Ruptura Loja x Depósito',M,y);y+=6;doc.setFont(undefined,'normal');
      doc.setFontSize(8);doc.setTextColor(80,80,80);
      var linesRup=doc.splitTextToSize(Engine.gerarAnaliseRuptura(results.ruptura,info),cw);
      if(y+linesRup.length*4.2+8>270){doc.addPage();y=M;}
      doc.text(linesRup,M,y);y+=linesRup.length*4.2+8;
    }
    if(results.dias){
      if(y>250){doc.addPage();y=M;}
      doc.setFontSize(11);doc.setTextColor.apply(doc,navy);doc.setFont(undefined,'bold');
      doc.text('Dias de Estoque',M,y);y+=6;doc.setFont(undefined,'normal');
      doc.setFontSize(8);doc.setTextColor(80,80,80);
      var linesDias=doc.splitTextToSize(Engine.gerarAnaliseDias(results.dias,info),cw);
      if(y+linesDias.length*4.2+8>270){doc.addPage();y=M;}
      doc.text(linesDias,M,y);y+=linesDias.length*4.2+8;
      var faixasR=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];
      var rowsFaixa=faixasR.map(function(fx){var itsFx=results.dias.items.filter(function(i){return i.faixa===fx;});var vlFx=itsFx.reduce(function(s,i){return s+(i.valorEstoque||0);},0);return[fx,itsFx.length,BRLi(vlFx)];});
      if(y+20>270){doc.addPage();y=M;}
      doc.autoTable({startY:y,head:[['Faixa','SKUs','Valor Estoque']],body:rowsFaixa,margin:{left:M,right:M},headStyles:{fillColor:[5,19,35],fontSize:7,fontStyle:'bold',halign:'left'},bodyStyles:{fontSize:7,halign:'left'},alternateRowStyles:{fillColor:[245,245,245]},styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},columnStyles:{1:{halign:'right'},2:{halign:'right'}}});
      y=doc.lastAutoTable.finalY+8;
    }
  }

  /* Gráficos por dimensão */
  function ckR(n){if(y+n>270){doc.addPage();y=M;}}
  function imgR(url,h){if(!url)return;ckR(h+10);try{doc.addImage(url,'PNG',M,y,cw,h);}catch(e){}y+=h+10;}
  function tituloR(t){ckR(14);doc.setFontSize(11);doc.setTextColor.apply(doc,navy);doc.text(t,M,y);y+=6;}

  if(y>245){doc.addPage();y=M;}
  doc.setFontSize(18);doc.setTextColor.apply(doc,navy);
  doc.text('Gráficos',M,y);y+=10;

  if(results.critica){
    var cc=results.critica;
    tituloR('Crítica — Faltas e sobras'+(cc.hasCategorias?' por categoria':''));
    var rowsCritR=cc.hasCategorias?cc.categorias.map(function(x){return{label:x.nome,a:Math.abs(x.faltaVal),b:x.sobraVal};}):[{label:'Total',a:Math.abs(cc.totalFaltas),b:cc.totalSobras}];
    ckR(rowsCritR.length*9+10);
    y+=_pdfMinibarDual(doc,M,y+4,cw,rowsCritR,'Faltas (R$)','Sobras (R$)',[211,47,47],[245,124,0],BRLi,BRLi)+10;
  }
  if(results.abc){
    var aa2=results.abc;
    tituloR('Investimento por curva ABC');
    var rowsABCR=[{label:'Curva A',a:aa2.fatA.invest,b:aa2.fatA.fat},{label:'Curva B',a:aa2.fatB.invest,b:aa2.fatB.fat},{label:'Curva C',a:aa2.fatC.invest,b:aa2.fatC.fat}];
    ckR(rowsABCR.length*9+10);
    y+=_pdfMinibarDual(doc,M,y+4,cw,rowsABCR,'Valor em estoque (R$)','Faturamento (R$)',[0,43,80],[97,207,0],BRLi,BRLi)+10;
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

  /* Slide 2.5: Gráficos (r102: mini-barras desenhadas como shapes, sem Chart.js/addChart nativo) */
  var chartDims=[];
  if(results.critica){
    var cc=results.critica;
    if(cc.hasCategorias&&cc.categorias.length){
      chartDims.push({title:'Crítica — Faltas/sobras por categoria',rows:cc.categorias.map(function(x){return{label:x.nome,a:Math.abs(x.faltaVal),b:x.sobraVal};}),labelA:'Faltas (R$)',labelB:'Sobras (R$)',colorA:'D32F2F',colorB:'F57C00',fmtA:function(v){return BRLi(v);},fmtB:function(v){return BRLi(v);}});
    }else{
      chartDims.push({title:'Crítica — Faltas × Sobras (R$)',rows:[{label:'Total',a:Math.abs(cc.totalFaltas),b:cc.totalSobras}],labelA:'Faltas (R$)',labelB:'Sobras (R$)',colorA:'D32F2F',colorB:'F57C00',fmtA:function(v){return BRLi(v);},fmtB:function(v){return BRLi(v);}});
    }
  }
  if(results.abc){var aa=results.abc;chartDims.push({title:'Investimento por curva ABC',rows:[{label:'Curva A',a:aa.fatA.invest,b:aa.fatA.fat},{label:'Curva B',a:aa.fatB.invest,b:aa.fatB.fat},{label:'Curva C',a:aa.fatC.invest,b:aa.fatC.fat}],labelA:'Valor estoque (R$)',labelB:'Faturamento (R$)',colorA:'002B50',colorB:'61CF00',fmtA:BRLi,fmtB:BRLi});}

  /* r102: desenha uma mini-barra dupla (shapes+texto) dentro dos limites estritos do retângulo [tx,tx+tw]x[ty,ty+th] — nunca extrapola a tile. */
  function _pptxMinibarTile(slide,cd,tx,ty,tw,th){
    var legendH=0.15,hasLegend=false,contentTy=ty,contentTh=th;
    if(th>(cd.rows.length*0.22+legendH+0.05)){
      hasLegend=true;contentTy=ty+legendH;contentTh=th-legendH;
      slide.addShape(pres.ShapeType.rect,{x:tx,y:ty,w:0.08,h:0.08,fill:{color:cd.colorA}});
      slide.addText(cd.labelA,{x:tx+0.11,y:ty-0.03,w:tw*0.45,h:legendH,fontSize:5,fontFace:'Arial',color:BODY,isTextBox:true,margin:0});
      slide.addShape(pres.ShapeType.rect,{x:tx+tw*0.5,y:ty,w:0.08,h:0.08,fill:{color:cd.colorB}});
      slide.addText(cd.labelB,{x:tx+tw*0.5+0.11,y:ty-0.03,w:tw*0.45,h:legendH,fontSize:5,fontFace:'Arial',color:BODY,isTextBox:true,margin:0});
    }
    var n=cd.rows.length;if(n<1)return;
    var rowH=Math.min(contentTh/n,0.42);
    var labelW=Math.min(tw*0.32,1.1);
    var maxBarW=Math.max(0.2,tw-labelW-0.65);
    var maxVal=0;cd.rows.forEach(function(r){maxVal=Math.max(maxVal,Math.abs(r.a||0),Math.abs(r.b||0));});
    if(maxVal<=0)maxVal=1;
    var cursorY=contentTy;
    for(var i=0;i<n;i++){
      if(cursorY+rowH>ty+th+0.001)break; /* nunca ultrapassa o limite inferior da tile */
      var r=cd.rows[i];
      var barH=Math.max(0.04,Math.min((rowH-0.04)/2,0.16));
      slide.addText(String(r.label).substring(0,18),{x:tx,y:cursorY,w:labelW,h:rowH,fontSize:6,fontFace:'Arial',color:BODY,isTextBox:true,margin:0,valign:'middle'});
      var wA=Math.max(0.02,Math.abs(r.a||0)/maxVal*maxBarW);
      slide.addShape(pres.ShapeType.rect,{x:tx+labelW,y:cursorY,w:wA,h:barH,fill:{color:cd.colorA}});
      slide.addText(cd.fmtA(r.a),{x:tx+labelW+maxBarW+0.03,y:cursorY,w:0.6,h:barH,fontSize:5.5,fontFace:'Arial',color:cd.colorA,isTextBox:true,margin:0});
      var y2=cursorY+barH+0.04;
      var wB=Math.max(0.02,Math.abs(r.b||0)/maxVal*maxBarW);
      slide.addShape(pres.ShapeType.rect,{x:tx+labelW,y:y2,w:wB,h:barH,fill:{color:cd.colorB}});
      slide.addText(cd.fmtB(r.b),{x:tx+labelW+maxBarW+0.03,y:y2,w:0.6,h:barH,fontSize:5.5,fontFace:'Arial',color:cd.colorB,isTextBox:true,margin:0});
      cursorY+=rowH;
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
        _pptxMinibarTile(sG,cd,cx,cy+0.24,cw2,ch2-0.28);
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

/* r102: substitui chartPerdaCategoria (Chart.js→PNG) por HTML de mini-barra nativa, reaproveitando as classes .minibar-* já definidas em _HTML_CSS. */
function _minibarPerdaHtml(catList){
  if(!catList||!catList.length)return'';
  var maxVal=0;
  catList.forEach(function(c){maxVal=Math.max(maxVal,c.perdaMensal||0);});
  if(maxVal<=0)maxVal=1;
  var h='<div class="minibar-legend"><span><span class="minibar-dot" style="background:#D32F2F"></span>Perda mensal (R$)</span></div><div class="minibar-list">';
  catList.forEach(function(c){
    var wP=Math.round((c.perdaMensal||0)/maxVal*100);
    h+='<div class="minibar-row"><div class="minibar-cat">'+c.nome+'</div><div class="minibar-bars">';
    h+='<div class="minibar-line"><span class="minibar-tag">Perda mensal</span><div class="minibar-track"><div class="minibar-fill falta" style="width:'+wP+'%"></div></div><span class="minibar-val text-red">'+_abbr(c.perdaMensal||0)+'</span></div>';
    h+='</div></div>';
  });
  h+='</div>';
  return h;
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
+".minibar-row.inv{grid-template-columns:190px 1fr}"
+".minibar-row:last-child{border-bottom:none}"
+".minibar-cat{font-weight:600;font-size:13px;color:var(--navy)}"
+".minibar-catval{font-size:11px;color:var(--muted);margin-top:2px}"
+".minibar-bars{display:flex;flex-direction:column;gap:6px}"
+".minibar-line{display:flex;align-items:center;gap:8px}"
+".minibar-tag{font-size:10px;color:var(--muted);width:46px;flex-shrink:0}"
+".minibar-track{flex:1;background:var(--light);border-radius:4px;height:12px;overflow:hidden}"
+".minibar-fill{height:100%;border-radius:4px}"
+".minibar-fill.falta{background:var(--red)}.minibar-fill.sobra{background:var(--amb)}.minibar-fill.vendas{background:var(--blue)}.minibar-fill.estoque{background:var(--amb)}"
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
+".analysis-box{background:var(--light);border-left:3px solid var(--green);border-radius:6px;padding:14px 18px;margin-bottom:20px}"
+".analysis-box p{font-size:13px;line-height:1.6;color:#333;margin:0 0 10px}.analysis-box p:last-child{margin-bottom:0}"
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
function _catTableHtml(catList,columns){
  var h='<div class="cat-table-wrap"><table class="cat-table"><thead><tr><th>Categoria</th>';
  columns.forEach(function(c){h+='<th class="num">'+c.label+'</th>';});
  h+='</tr></thead><tbody>';
  catList.forEach(function(cat){
    h+='<tr><td>'+cat.nome+'</td>';
    columns.forEach(function(c){var v=c.fmt?c.fmt(cat[c.key]):cat[c.key];h+='<td class="num">'+v+'</td>';});
    h+='</tr>';
  });
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
  catList.forEach(function(c){maxPct=Math.max(maxPct,c.pctFat||0,c.pctInvest||0);});
  if(maxPct<=0)maxPct=1;
  var h='<div class="minibar-legend"><span><span class="minibar-dot" style="background:#1565C0"></span>Participação nas vendas</span><span><span class="minibar-dot" style="background:#F57C00"></span>Participação no valor do estoque</span></div><div class="minibar-list">';
  catList.forEach(function(c){
    var wV=Math.round((c.pctFat||0)/maxPct*100),wE=Math.round((c.pctInvest||0)/maxPct*100);
    h+='<div class="minibar-row inv"><div class="minibar-cat">'+c.nome+'<div class="minibar-catval">'+BRLi(c.investimento)+' em estoque</div></div><div class="minibar-bars">';
    h+='<div class="minibar-line"><span class="minibar-tag">Vendas</span><div class="minibar-track"><div class="minibar-fill vendas" style="width:'+wV+'%"></div></div><span class="minibar-val">'+PCT(c.pctFat)+'</span></div>';
    h+='<div class="minibar-line"><span class="minibar-tag">Estoque</span><div class="minibar-track"><div class="minibar-fill estoque" style="width:'+wE+'%"></div></div><span class="minibar-val">'+PCT(c.pctInvest)+'</span></div>';
    h+='</div></div>';
  });
  h+='</div>';
  return h;
}
function _dataTableHtml(cols,rows){
  var h='<div class="table-wrap"><table class="data-table"><thead><tr>';
  cols.forEach(function(c){h+='<th class="'+(c.align||'')+'">'+c.label+'</th>';});
  h+='</tr></thead><tbody>';
  rows.forEach(function(r){
    h+='<tr>';
    cols.forEach(function(c){
      var v=r[c.f],out;
      if(c.brl)out=BRL(v);else if(c.n)out=(v===null||v===undefined)?'—':NUMx(v);else out=(v===null||v===undefined||v==='')?'—':v;
      h+='<td class="'+(c.align||'')+'">'+out+'</td>';
    });
    h+='</tr>';
  });
  h+='</tbody></table></div>';
  return h;
}
var NUMx=function(v){return Number(v).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});};

function _bodyCritica(c){
  var h='<div class="metrics">';
  h+='<div class="metric"><div class="metric-label">Acuracidade</div><div class="metric-value text-green">'+PCT(c.acuracidade)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Valor estoque contado</div><div class="metric-value">'+BRLi(c.valorEstoqueContado)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Valor das faltas</div><div class="metric-value text-red">'+BRLi(c.totalFaltas)+'</div><div class="metric-detail">'+NUM(c.faltaCount)+' SKUs</div></div>';
  h+='<div class="metric"><div class="metric-label">Valor das sobras</div><div class="metric-value text-amber">'+BRLi(c.totalSobras)+'</div><div class="metric-detail">'+NUM(c.sobraCount)+' SKUs</div></div>';
  h+='<div class="metric"><div class="metric-label">Valor Quebra</div><div class="metric-value text-red">'+BRLi(c.saldoLiquido)+'</div></div>';
  h+='</div>';
  if(c.hasCategorias){
    h+='<div class="section-title">Resultado por categoria</div>';
    h+=_catCardsHtml(c.categorias,[{label:'Acuracidade',key:'acuracidade',fmt:PCT},{label:'Faltas',key:'faltaVal',fmt:BRLi},{label:'Sobras',key:'sobraVal',fmt:BRLi},{label:'Saldo',key:'saldo',fmt:BRLi}]);
    h+=_minibarFaltaSobraHtml(c.categorias);
  }
  h+='<div class="section-title">Itens ('+NUM(c.items.length)+')</div>';
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'Categoria',f:'categoria'},{label:'Qtd sist.',f:'qtdSistema',align:'text-right',n:true},{label:'Qtd cont.',f:'qtdContada',align:'text-right',n:true},{label:'Dif. qtd',f:'difQtd',align:'text-right',n:true},{label:'Dif. R$',f:'difValor',align:'text-right',brl:true},{label:'Status',f:'status',align:'text-center'}],c.items);
  return h;
}
function _bodyRuptura(r){
  var h='<div class="metrics">';
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
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'Categoria',f:'categoria'},{label:'ABC fat.',f:'abc_valorVendido90',align:'text-center'},{label:'Qtd depósito',f:'deposito',align:'text-right',n:true},{label:'Qtd loja',f:'loja',align:'text-right',n:true},{label:'Venda méd/dia',f:'vendaMediaDia',align:'text-right',n:true},{label:'Fat. méd/dia',f:'fatMediaDia',align:'text-right',brl:true}],r.items);
  return h;
}
function _bodyDias(d){
  var h='<div class="metrics">';
  h+='<div class="metric"><div class="metric-label">Cobertura de estoque</div><div class="metric-value">'+d.coberturaGeral+' dias</div></div>';
  h+='<div class="metric"><div class="metric-label">Cobertura curva A</div><div class="metric-value">'+d.coberturaA+' dias</div></div>';
  h+='<div class="metric"><div class="metric-label">Cobertura curva B</div><div class="metric-value">'+d.coberturaB+' dias</div></div>';
  h+='<div class="metric"><div class="metric-label">Cobertura curva C</div><div class="metric-value">'+d.coberturaC+' dias</div></div>';
  h+='<div class="metric"><div class="metric-label">SKUs sem giro</div><div class="metric-value text-red">'+NUM(d.semGiro)+'</div></div>';
  h+='</div>';
  if(d.hasCategorias){
    h+='<div class="section-title">Cobertura por categoria</div>';
    h+=_catTableHtml(d.categorias,[{label:'Cobertura média',key:'mediaCobertura',fmt:function(v){return v+' dias';}},{label:'Val. estoque',key:'valorEstoque',fmt:BRLi},{label:'Ruptura + Alto risco',key:'criticos',fmt:NUM},{label:'Sem giro',key:'semGiro',fmt:NUM},{label:'Excesso (31+d)',key:'excessos',fmt:NUM}]);
  }
  h+='<div class="section-title">Itens ('+NUM(d.items.length)+')</div>';
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'Categoria',f:'categoria'},{label:'Qtd estoque',f:'qtdEstoque',align:'text-right',n:true},{label:'Dias estoque',f:'diasEstoque',align:'text-right'},{label:'Cobertura',f:'faixa',align:'text-center'},{label:'Val. estoque',f:'valorEstoque',align:'text-right',brl:true},{label:'ABC fat.',f:'abcFat',align:'text-center'}],d.items);
  return h;
}
function _bodyABC(a){
  var h='<div class="metrics">';
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
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'ABC fat.',f:'abcFat',align:'text-center'},{label:'ABC lucro',f:'abcLucro',align:'text-center'},{label:'Qtd estoque',f:'qtdEstoque',align:'text-right',n:true},{label:'Valor estoque',f:'valorInvestido',align:'text-right',brl:true},{label:'Fat. 90d',f:'fat90',align:'text-right',brl:true},{label:'Lucro 90d',f:'lucro90',align:'text-right',brl:true}],a.items);
  return h;
}
function _bodyPerda(p){
  var h='<div class="metrics">';
  h+='<div class="metric"><div class="metric-label">Venda perdida / dia</div><div class="metric-value text-red">'+BRLi(p.totalPerdaFat)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Lucro perdido / dia</div><div class="metric-value text-red">'+BRLi(p.totalPerdaLucro)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">Perda mensal (fat.)</div><div class="metric-value text-red">'+BRLi(p.perdaMensal)+'</div></div>';
  h+='<div class="metric"><div class="metric-label">SKUs em ruptura</div><div class="metric-value">'+NUM(p.totalSKUs)+'</div></div>';
  h+='</div>';
  h+='<div class="loss-cards"><div class="loss-card a"><div class="loss-title">Curva A — perda/dia</div><div class="loss-main">'+BRLi(p.classA.perda)+'</div></div><div class="loss-card b"><div class="loss-title">Curva B — perda/dia</div><div class="loss-main">'+BRLi(p.classB.perda)+'</div></div><div class="loss-card c"><div class="loss-title">Curva C — perda/dia</div><div class="loss-main">'+BRLi(p.classC.perda)+'</div></div></div>';
  if(p.hasCategorias){
    h+='<div class="section-title">Perda projetada por categoria</div>';
    h+=_catCardsHtml(p.categorias,[{label:'Rupturas',key:'totalRupturas',fmt:NUM},{label:'Perda fat./dia',key:'perdaFatDia',fmt:BRLi},{label:'Perda lucro/dia',key:'perdaLucroDia',fmt:BRLi},{label:'Perda mensal',key:'perdaMensal',fmt:BRLi},{label:'Rupturas A',key:'rupturaA',fmt:NUM}]);
    h+=_minibarPerdaHtml(p.categorias);
  }
  h+='<div class="section-title">Itens ('+NUM(p.items.length)+')</div>';
  h+=_dataTableHtml([{label:'SKU',f:'sku'},{label:'Descrição',f:'descricao'},{label:'Categoria',f:'categoria'},{label:'ABC fat.',f:'abcFat',align:'text-center'},{label:'Perda fat./dia',f:'perdaFatDia',align:'text-right',brl:true},{label:'Perda lucro/dia',f:'perdaLucroDia',align:'text-right',brl:true},{label:'Perda fat./mês',f:'perdaFatMes',align:'text-right',brl:true}],p.items);
  return h;
}

var _HTML_TITLES={critica:'Crítica do Inventário (Acuracidade)',ruptura:'Ruptura Loja x Depósito',dias:'Dias de Estoque (Cobertura)',abc:'Investimento em Estoque',perda:'Projeção de Perda'};
var _HTML_BODYFN={critica:_bodyCritica,ruptura:_bodyRuptura,dias:_bodyDias,abc:_bodyABC,perda:_bodyPerda};

function generateHTML(type,data,pd,logo,info){
  info=info||{};
  var d=data[type];
  if(!d){alert('Este relatório ainda não foi gerado.');return;}
  var titulo=_HTML_TITLES[type]||type;
  var body=_HTML_BODYFN[type](d);
  var html=_htmlPage(titulo,info,pd,body,logo);
  _downloadHTML('auditoria_'+type+'_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.html',html);
}

function generateResumoHTML(results,recs,info,unidade,logo){
  info=info||{};
  var pd=new Date().toLocaleString('pt-BR');
  var h='<div class="section-title" style="font-size:20px;margin-top:0">Resumo Executivo</div>';
  h+='<div class="metrics">';
  if(results.critica){var c=results.critica;h+='<div class="metric"><div class="metric-label">Acuracidade</div><div class="metric-value text-green">'+PCT(c.acuracidade)+'</div><div class="metric-detail">'+NUM(c.totalSKUs)+' SKUs · '+NUM(c.faltaCount)+' faltas · '+NUM(c.sobraCount)+' sobras</div></div>';}
  if(results.ruptura){var ru=results.ruptura;h+='<div class="metric"><div class="metric-label">Ruptura loja x depósito</div><div class="metric-value text-red">'+PCT(ru.taxaRuptura)+'</div><div class="metric-detail">'+NUM(ru.totalRupturas)+' SKUs em falta</div></div>';}
  if(results.dias){var d=results.dias;h+='<div class="metric"><div class="metric-label">Cobertura</div><div class="metric-value">'+Engine.round2(d.coberturaGeral)+' dias</div><div class="metric-detail">Sem giro: '+NUM(d.semGiro)+' · Excesso: '+NUM(d.excessos)+'</div></div>';}
  if(results.abc){var a=results.abc;h+='<div class="metric"><div class="metric-label">Investimento em estoque</div><div class="metric-value">'+BRLi(a.totalInvest)+'</div><div class="metric-detail">Faturamento: '+BRLi(a.totalFat)+'</div></div>';}
  if(results.perda){var pe=results.perda;h+='<div class="metric"><div class="metric-label">Projeção de perda mensal</div><div class="metric-value text-red">'+BRLi(pe.perdaMensal)+'</div><div class="metric-detail">'+NUM(pe.totalSKUs)+' SKUs identificados</div></div>';}
  h+='</div>';
  if(recs&&recs.length){
    h+='<div class="section-title">Recomendações</div><div style="background:var(--light);border-radius:8px;padding:16px 20px;margin-bottom:20px">';
    recs.forEach(function(rec){h+='<div style="padding:8px 0;border-bottom:1px solid var(--border);font-size:13px">'+rec.texto+'</div>';});
    h+='</div>';
  }
  if(results.critica)h+='<h2 style="color:var(--navy);margin:28px 0 4px">Crítica do Inventário (Acuracidade)</h2>'+_bodyCritica(results.critica);
  if(results.ruptura)h+='<h2 style="color:var(--navy);margin:28px 0 4px">Ruptura Loja x Depósito</h2>'+_bodyRuptura(results.ruptura);
  if(results.dias)h+='<h2 style="color:var(--navy);margin:28px 0 4px">Dias de Estoque (Cobertura)</h2>'+_bodyDias(results.dias);
  if(results.abc)h+='<h2 style="color:var(--navy);margin:28px 0 4px">Investimento em Estoque</h2>'+_bodyABC(results.abc);
  if(results.perda)h+='<h2 style="color:var(--navy);margin:28px 0 4px">Projeção de Perda</h2>'+_bodyPerda(results.perda);
  var html=_htmlPage('Relatório Completo — Análise de Inventários',info,pd,h,logo);
  _downloadHTML('relatorio_completo_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.html',html);
}

function generateComparativoHTML(comp,units,info,analiseTxt,logo){
  info=info||{};
  var pd=new Date().toLocaleString('pt-BR');
  var h='<div class="section-title" style="font-size:20px;margin-top:0">Comparativo entre unidades — '+comp.unidades.length+' unidades</div>';
  /* r103: Análise Comparativa — mesmo texto exibido na tela (IA ou fallback local); cards de Sobreposição de SKUs removidos */
  var txt=analiseTxt||'Análise comparativa entre '+comp.unidades.length+' unidades do cliente '+(info.cliente||'')+' referente ao inventário de '+(info.dataInventario||'')+'.';
  h+='<div class="section-title">Análise Comparativa</div>';
  h+='<div class="analysis-box">'+String(txt).split('\n\n').map(function(par){return '<p>'+par+'</p>';}).join('')+'</div>';
  h+='<div class="section-title">Indicadores por unidade</div>';
  h+=_catCardsHtml(comp.unidades.map(function(u){return{nome:u.unidade,acuracidade:u.acuracidade,valorQuebra:u.saldoLiquido,coberturaGeral:u.coberturaGeral,totalRupturas:u.totalRupturas,perdaFatDia:u.perdaFatDia};}),[
    {label:'Acuracidade',key:'acuracidade',fmt:function(v){return v!=null?PCT(v):'—';}},
    {label:'Valor Quebra',key:'valorQuebra',fmt:function(v){return v!=null?BRLi(v):'—';}},
    {label:'Cobertura de Estoque',key:'coberturaGeral',fmt:function(v){return v!=null?v+' dias':'—';}},
    {label:'SKUs em Ruptura',key:'totalRupturas',fmt:function(v){return v!=null?NUM(v):'—';}},
    {label:'Venda perdida / dia',key:'perdaFatDia',fmt:function(v){return v!=null?BRLi(v):'—';}}
  ]);
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
