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
  t+='\n\nUma acuracidade de '+PCT(c.acuracidade)+' indica que a cada 100 posições, cerca de '+desvio+' apresentam algum desvio — número que serve de termômetro para a confiabilidade do estoque no ERP.';
  if(w.length){
    var c1=w[0], nomes=c1.nome+' ('+PCT(c1.acuracidade)+')';
    if(w.length>1){nomes+=' e '+w[1].nome+' ('+PCT(w[1].acuracidade)+')';}
    t+=' As categorias com menor acuracidade foram '+nomes+', que concentram a maior fragilidade de controle e devem ser priorizadas em recontagens e na revisão dos processos de entrada e baixa de mercadoria.';
  }
  t+=' Faltas em produtos perecíveis costumam apontar para inversão de códigos no registro de vendas ou perdas não registradas (quebra, vencimento, furto, desidratação), enquanto sobras sugerem falhas de lançamento na entrada.';
  return t;
}
function metCritica(){return'Cada SKU é comparado entre o saldo registrado no sistema (ERP) e a contagem física realizada no inventário. A diferença (contado - sistema) determina a classificação: Falta (negativo), Sobra (positivo) ou Sem divergência (zero). O valor financeiro da divergência é calculado multiplicando a diferença pelo custo unitário do produto. A acuracidade representa o percentual de SKUs sem nenhuma divergência sobre o total analisado.';}
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
  if(data.critica){var c=data.critica;R=addST(ws,R,'CRÍTICA DO INVENTÁRIO');R=addKR(ws,R,['ACURACIDADE','VALOR DAS FALTAS','VALOR DAS SOBRAS','SALDO LÍQUIDO'],[PCT(c.acuracidade),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido)],[C.green,C.red,C.amb,C.red]);if(c.hasCategorias)R=addDT(ws,R,['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,x.total,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{0:'left',1:'right',2:'right',3:'right',4:'right',5:'right'});}
  if(data.ruptura){var r=data.ruptura;R=addST(ws,R,'RUPTURA');R=addKR(ws,R,['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CURVA A (FAT.)','RUPTURA CURVA A (LUCRO)'],[PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],[C.red,C.text,C.red,C.red]);}
  if(data.dias){var d=data.dias,fv=fxV(d.items);R=addST(ws,R,'DIAS DE ESTOQUE');R=addKR(ws,R,['COBERTURA GERAL','CURVA A','CURVA B','CURVA C','SEM GIRO'],[d.coberturaGeral+' dias',d.coberturaA+' dias',d.coberturaB+' dias',d.coberturaC+' dias',NUM(d.semGiro)],[C.text,C.text,C.text,C.text,C.red]);var fo=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];R=addDT(ws,R,['Faixa','SKUs','% SKUs','Valor Estoque (R$)','% do Valor'],fo.map(function(f){var cn=d.items.filter(function(i){return i.faixa===f;}).length;var vl=fv.f[f]||0;return[f,cn,PCT(d.total?cn/d.total*100:0),BRLi(vl),PCT(fv.t?vl/fv.t*100:0)];}),{0:'left',1:'right',2:'right',3:'right',4:'right'});}
  if(data.abc){var a=data.abc;R=addST(ws,R,'INVESTIMENTO ABC');R=addKR(ws,R,['VALOR TOTAL EM ESTOQUE','FATURAMENTO 90D','LUCRO 90D','SKUS'],[BRLi(a.totalInvest),BRLi(a.totalFat),BRLi(a.totalLucro),NUM(a.items.length)],[C.text,C.green,C.green,C.text]);R=addDT(ws,R,['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],{0:'center',1:'right',2:'right',3:'right',4:'right'});}
  if(data.perda){var pe=data.perda;R=addST(ws,R,'PROJEÇÃO DE PERDA');R=addKR(ws,R,['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS'],[BRLi(pe.totalPerdaFat),BRLi(pe.totalPerdaLucro),BRLi(pe.perdaMensal),NUM(pe.totalSKUs)],[C.red,C.red,C.red,C.text]);R=addDT(ws,R,['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',pe.classA.count,BRLi(pe.classA.perda),BRLi(pe.classA.lucro),PCT(pe.classA.pct),BRLi(pe.classA.perda*30)],['B',pe.classB.count,BRLi(pe.classB.perda),BRLi(pe.classB.lucro),PCT(pe.classB.pct),BRLi(pe.classB.perda*30)],['C',pe.classC.count,BRLi(pe.classC.perda),BRLi(pe.classC.lucro),PCT(pe.classC.pct),BRLi(pe.classC.perda*30)]],{0:'center',1:'right',2:'right',3:'right',4:'right',5:'right'});}
  ws['!cols']=[{wch:28},{wch:18},{wch:16},{wch:18},{wch:16},{wch:18},{wch:16},{wch:16}];ws['!rows']=[{hpt:28},{hpt:20}];
  XLSX.utils.book_append_sheet(wb,ws,'Dashboard');

  /* CRITICA RESUMO + TOP20 */
  if(sel.criticaResumo&&data.critica){var wsC={},rw=0,c=data.critica;rw=addBH(wsC,rw,info,pd,6);rw=addST(wsC,rw,'RESUMO DA CRÍTICA');var sL=['ACURACIDADE','SKUs analisados','SKUs sem divergência','SKUs com falta','SKUs com sobra','Valor das faltas','Valor das sobras','Saldo líquido'],sV=[PCT(c.acuracidade),c.totalSKUs,c.okCount,c.faltaCount,c.sobraCount,BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido)];for(var i=0;i<sL.length;i++){sC(wsC,rw+i,0,sL[i],sB('left',true));sC(wsC,rw+i,1,sV[i],sB('right'));}rw+=sL.length+1;
  if(c.hasCategorias){rw=addST(wsC,rw,'RESULTADO POR CATEGORIA');rw=addDT(wsC,rw,['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],c.categorias.map(function(x){return[x.nome,x.total,PCT(x.acuracidade),BRLi(x.faltaVal),BRLi(x.sobraVal),BRLi(x.saldo)];}),{0:'left',1:'right',2:'right',3:'right',4:'right',5:'right'});}
  var ct=top20Cat(c.items);var tH=['SKU','Descrição','Qtd Sist','Qtd Contada','Dif. Qtd','Dif. R$'],tA={0:'left',1:'left',2:'right',3:'right',4:'right',5:'right'};
  ct.forEach(function(cat){if(cat.faltas.length){rw=addST(wsC,rw,'TOP '+cat.faltas.length+' FALTAS — '+cat.nome);rw=addDT(wsC,rw,tH,cat.faltas.map(function(i){return[i.sku,i.descricao,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor)];}),tA);}if(cat.sobras.length){rw=addST(wsC,rw,'TOP '+cat.sobras.length+' SOBRAS — '+cat.nome);rw=addDT(wsC,rw,tH,cat.sobras.map(function(i){return[i.sku,i.descricao,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor)];}),tA);}if(cat.zerados.length){rw=addST(wsC,rw,'TOP '+cat.zerados.length+' ZERADOS — '+cat.nome);rw=addDT(wsC,rw,['SKU','Descrição','Qtd Sistema','Valor Perdido'],cat.zerados.map(function(i){return[i.sku,i.descricao,i.qtdSistema,BRL(i.qtdSistema*i.custoUnit)];}),{0:'left',1:'left',2:'right',3:'right'});}});
  wsC['!cols']=[{wch:16},{wch:32},{wch:14},{wch:14},{wch:12},{wch:16}];wsC['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsC,'Crítica - Resumo');}

  if(sel.criticaDetalhe&&data.critica){var wsCD={},rw=0;rw=addBH(wsCD,rw,info,pd,8);rw=addDT(wsCD,rw,['SKU','Descrição','Categoria','Qtd Sistema','Qtd Contada','Dif. Qtd','Dif. R$','Status'],data.critica.items.map(function(i){return[i.sku,i.descricao,i.categoria,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.difValor),i.status];}),{0:'left',1:'left',2:'left',3:'right',4:'right',5:'right',6:'right',7:'center'});wsCD['!cols']=[{wch:14},{wch:32},{wch:18},{wch:12},{wch:12},{wch:10},{wch:14},{wch:10}];wsCD['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsCD,'Crítica - Detalhado');}

  if(sel.ruptura&&data.ruptura){var wsR={},rw=0;rw=addBH(wsR,rw,info,pd,9);rw=addDT(wsR,rw,['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Depósito','Qtd Loja','Venda Méd/Dia','Fat. Méd/Dia'],data.ruptura.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abc_valorVendido90||'C',i.abc_lucro90||'C',i.deposito,i.loja,R2(i.vendaMediaDia),BRL(i.fatMediaDia||0)];}),{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right'});wsR['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:14},{wch:10},{wch:14},{wch:14}];wsR['!rows']=[{hpt:28},{hpt:20}];XLSX.utils.book_append_sheet(wb,wsR,'Ruptura');}

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
  if(apiUrl&&data[rt]){
    var metricas='';
    if(rt==='critica'){metricas='Total SKUs: '+data.critica.totalSKUs+'\nAcuracidade: '+PCT(data.critica.acuracidade)+'\nFaltas: '+NUM(data.critica.faltaCount)+' ('+BRLi(data.critica.totalFaltas)+')\nSobras: '+NUM(data.critica.sobraCount)+' ('+BRLi(data.critica.totalSobras)+')\nSaldo líquido: '+BRLi(data.critica.saldoLiquido);if(data.critica.hasCategorias)metricas+='\nCategorias:\n'+data.critica.categorias.map(function(c){return c.nome+': '+NUM(c.total)+' SKUs, acuracidade '+PCT(c.acuracidade)+', faltas '+BRLi(c.faltaVal)+', sobras '+BRLi(c.sobraVal);}).join('\n');}
    else if(rt==='ruptura'){metricas='Total rupturas: '+NUM(data.ruptura.totalRupturas)+'\nTaxa: '+PCT(data.ruptura.taxaRuptura)+'\nCurva A: '+NUM(data.ruptura.rupturaA)+'\nCurva B: '+NUM(data.ruptura.rupturaB);}
    else if(rt==='dias'){metricas='Cobertura geral: '+data.dias.coberturaGeral+' dias\nCobertura A: '+data.dias.coberturaA+' dias\nCobertura B: '+data.dias.coberturaB+' dias\nTotal SKUs: '+NUM(data.dias.total);}
    else if(rt==='abc'){metricas='Curva A Fat: '+BRLi(data.abc.fatA.invest)+' ('+PCT(data.abc.fatA.pctInvest)+' estoque, '+PCT(data.abc.fatA.pctFat)+' fat.)\nCurva B Fat: '+BRLi(data.abc.fatB.invest)+'\nCurva C Fat: '+BRLi(data.abc.fatC.invest);}
    else if(rt==='perda'){metricas='Perda fat./dia: '+BRLi(data.perda.totalPerdaFat)+'\nPerda lucro/dia: '+BRLi(data.perda.totalPerdaLucro)+'\nPerda mensal: '+BRLi(data.perda.perdaMensal)+'\nTotal SKUs: '+NUM(data.perda.totalSKUs);}
    try{
      var xhr=new XMLHttpRequest();
      xhr.open('POST',apiUrl,false);// síncrono
      xhr.send(JSON.stringify({action:'gerarResumoInventarioIA',cliente:info.cliente||'',unidade:info.unidade||'',data:info.dataInventario||'',diasVenda:info.diasVenda||90,metricas:metricas,secao:rt}));
      var resp=JSON.parse(xhr.responseText);
      if(resp.ok&&resp.resumos&&resp.resumos[rt]){
        window._iaResumos=window._iaResumos||{};
        window._iaResumos[rt]=resp.resumos[rt];
      }
    }catch(e){console.log('IA fallback:',e);}
  }
  _generatePDFInternal(rt,data,pd,logo,info);
}
function _generatePDFInternal(rt,data,pd,logo,info){
  info=info||{};var jsPDF=window.jspdf.jsPDF;var doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});var W=210,H=297,M=15,y=0;
  function hdr(){doc.setFillColor(5,19,35);doc.rect(0,0,W,22,'F');if(logo){try{doc.addImage(logo,'PNG',M,3,36,12);}catch(e){}}doc.setFontSize(9);doc.setTextColor(255,255,255);doc.text((info.cliente||'')+' — '+(info.unidade||''),W-M,7,{align:'right'});doc.setFontSize(7);doc.setTextColor(200,220,255);doc.text('Inventário: '+(info.dataInventario||'—'),W-M,12,{align:'right'});doc.setTextColor(180,180,200);doc.text('Processado em '+pd,W-M,17,{align:'right'});y=28;}
  function ftr(pg){doc.setFontSize(7);doc.setTextColor(150,150,150);doc.text('Formula Code Tecnologia, Gestão e Automação',M,H-6);doc.text('Página '+pg,W-M,H-6,{align:'right'});doc.setDrawColor(200,200,200);doc.line(M,H-10,W-M,H-10);}
  function chk(n){if(y+n>H-18){doc.addPage();hdr();ftr(doc.getNumberOfPages());}}
  function ttl(t){chk(12);doc.setFontSize(14);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFontSize(8);doc.setTextColor(150,150,150);doc.setFont(undefined,'normal');doc.text('Relatório gerado automaticamente pelo sistema Formula Code',M,y);y+=8;}
  function sec(t){chk(10);doc.setFontSize(11);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');doc.text(t,M,y);y+=6;doc.setFont(undefined,'normal');}
  function aT(h,b,o){chk(20);doc.autoTable({startY:y,head:[h],body:b,margin:{left:M,right:M},headStyles:{fillColor:[5,19,35],fontSize:7,fontStyle:'bold',halign:'left'},bodyStyles:{fontSize:7,halign:'left'},alternateRowStyles:{fillColor:[245,245,245]},styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},columnStyles:o||{}});y=doc.lastAutoTable.finalY+6;}
  function kpi(lb,vl,cl){chk(18);var cw=(W-2*M)/lb.length;doc.setFillColor(245,245,245);doc.roundedRect(M,y-2,W-2*M,16,2,2,'F');for(var i=0;i<lb.length;i++){var x=M+i*cw+4;doc.setFontSize(7);doc.setTextColor(150,150,150);doc.setFont(undefined,'bold');doc.text(lb[i],x,y+3);doc.setFontSize(11);doc.setFont(undefined,'bold');var cc=cl[i]||[51,51,51];doc.setTextColor(cc[0],cc[1],cc[2]);doc.text(String(vl[i]),x,y+10);}doc.setFont(undefined,'normal');y+=20;}
  function bloco(txt){chk(16);doc.setFontSize(8);doc.setTextColor(80,80,80);doc.setFont(undefined,'normal');var lines=doc.splitTextToSize(txt,W-2*M);doc.text(lines,M,y);y+=lines.length*3.5+4;}

  hdr();ftr(1);

  if(rt==='critica'){
    var c=data.critica;
    ttl('Crítica do inventário — Resumo executivo');
    var iaR=window._iaResumos&&window._iaResumos.critica;
    sec('Análise');bloco(iaR||sumCritica(c));
    sec('Metodologia');bloco(metCritica());
    sec('Indicadores gerais');
    kpi(['ACURACIDADE','VALOR DAS FALTAS','VALOR DAS SOBRAS','SALDO LÍQUIDO'],[PCT(c.acuracidade),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido)],[[0,183,74],[211,47,47],[245,124,0],[211,47,47]]);
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
    ttl('Ruptura — Resumo executivo');
    var iaRup=window._iaResumos&&window._iaResumos.ruptura;
    sec('Análise');bloco(iaRup||(temVendas?sumRuptura(r):'Foram identificados '+NUM(r.totalRupturas)+' itens em ruptura (presentes no depósito mas ausentes no salão de vendas). Sem dados de vendas disponíveis para análise de impacto financeiro.'));
    sec('Metodologia');bloco(metRuptura(info.diasVenda));
    sec('Indicadores gerais');kpi(['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CURVA A (FAT.)','RUPTURA CURVA A (LUCRO)'],[PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],[[211,47,47],[51,51,51],[211,47,47],[211,47,47]]);
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
    sec('Análise');bloco(iaDias||sumDias(d,info.diasVenda));sec('Metodologia');bloco(metDias(info.diasVenda));
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
    sec('Análise');bloco(iaABC||sumABC(a,info.diasVenda));sec('Metodologia');bloco(metABC(info.diasVenda));
    sec('Curva ABC por faturamento');aT(['Curva','Valor Estoque (R$)','% Estoque','Faturamento (R$)','% Faturamento'],[['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
    sec('Curva ABC por lucro');aT(['Curva','Valor Estoque (R$)','% Estoque','Lucro (R$)','% Lucro'],[['A',BRLi(a.lucA.invest),PCT(a.lucA.pctInvest),BRLi(a.lucA.luc),PCT(a.lucA.pctLuc)],['B',BRLi(a.lucB.invest),PCT(a.lucB.pctInvest),BRLi(a.lucB.luc),PCT(a.lucB.pctLuc)],['C',BRLi(a.lucC.invest),PCT(a.lucC.pctInvest),BRLi(a.lucC.luc),PCT(a.lucC.pctLuc)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
  }
  else if(rt==='perda'){
    var p=data.perda;ttl('Projeção de venda perdida — Resumo executivo');
    var iaPerda=window._iaResumos&&window._iaResumos.perda;
    sec('Análise');bloco(iaPerda||sumPerda(p,info.diasVenda));sec('Metodologia');bloco(metPerda(info.diasVenda));
    sec('Indicadores gerais');kpi(['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL','SKUS'],[BRLi(p.totalPerdaFat),BRLi(p.totalPerdaLucro),BRLi(p.perdaMensal),NUM(p.totalSKUs)],[[211,47,47],[211,47,47],[211,47,47],[51,51,51]]);
    sec('Impacto por curva ABC');aT(['Curva','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% Perda','Perda Mensal'],[['A',p.classA.count,BRLi(p.classA.perda),BRLi(p.classA.lucro),PCT(p.classA.pct),BRLi(p.classA.perda*30)],['B',p.classB.count,BRLi(p.classB.perda),BRLi(p.classB.lucro),PCT(p.classB.pct),BRLi(p.classB.perda*30)],['C',p.classC.count,BRLi(p.classC.perda),BRLi(p.classC.lucro),PCT(p.classC.pct),BRLi(p.classC.perda*30)]],{1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});
    var pH2=['SKU','Descrição','Categoria','Perda Fat./Mês','Perda Lucro/Mês'],pO2={3:{halign:'right'},4:{halign:'right'}};
    ['A','B','C'].forEach(function(cls){var it=p.items.filter(function(i){return i.abcFat===cls;}).sort(function(a,b){return b.perdaFatMes-a.perdaFatMes;});if(it.length){sec('Curva '+cls+' — '+it.length+' itens');aT(pH2,it.map(function(i){return[i.sku,i.descricao,i.categoria||'',BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];}),pO2);}});
  }
  chk(12);doc.setFontSize(7);doc.setTextColor(150,150,150);
  doc.text('Nota: relatório baseado em dados processados em '+pd+'. Valores projetados são estimativas.',M,y);
  doc.save('resumo_'+rt+'_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.pdf');
}
/* ========== PDF COMPARATIVO ========== */
function generateComparativoPDF(comp, units, info, iaTextos){
  iaTextos=iaTextos||{};info=info||{};
  var jsPDF=window.jspdf.jsPDF;var doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
  var W=210,H=297,M=15,y=0;
  var pd=new Date().toLocaleString('pt-BR');

  function hdr(){doc.setFillColor(5,19,35);doc.rect(0,0,W,22,'F');try{var st=window.App?window.App.getState():null;if(st){var logoSrc=document.getElementById('logo').src;if(logoSrc)doc.addImage(logoSrc,'PNG',M,3,36,12);}}catch(e){}doc.setFontSize(9);doc.setTextColor(255,255,255);doc.text('COMPARATIVO — '+(info.cliente||''),W-M,7,{align:'right'});doc.setFontSize(7);doc.setTextColor(200,220,255);doc.text(comp.unidades.map(function(u){return u.unidade;}).join(' × '),W-M,12,{align:'right'});doc.setTextColor(180,180,200);doc.text('Inventário: '+(info.dataInventario||'—')+' | Gerado em '+pd,W-M,17,{align:'right'});y=28;}
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
  if(iaTextos.analise_ruptura){sec('Análise — Ruptura');bloco(iaTextos.analise_ruptura);}
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

return{generateExcel:generateExcel,generateExcelBlob:generateExcelBlob,buildExcelWorkbook:buildExcelWorkbook,generatePDF:generatePDF,generateComparativoPDF:generateComparativoPDF,generateComparativoExcel:generateComparativoExcel};
})();
