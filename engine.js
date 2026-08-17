/* engine.js — v1.0 — Motor de cálculos com custo derivado global e cobertura corrigida */
var Engine = (function(){
  "use strict";
  function round2(n){ return Math.round((n||0)*100)/100; }
  function roundInt(n){ return Math.round(n||0); }
  function safeDiv(a,b){ return b ? a/b : null; }

  function groupByCategoria(items, aggFn){
    var cats = {};
    items.forEach(function(it){
      var cat = it.categoria || 'Sem categoria';
      if(!cats[cat]) cats[cat] = {nome:cat, items:[]};
      cats[cat].items.push(it);
    });
    return Object.keys(cats).map(function(k){ return aggFn(cats[k]); }).sort(function(a,b){ return (b.destaque||0) - (a.destaque||0); });
  }
  function hasRealCategorias(catList){
    return catList.length > 1 || (catList.length===1 && catList[0].nome !== 'Sem categoria');
  }

  /* ===== Mapa global de custo unitário derivado das vendas (CMV/Qtde) ===== */
  function buildCustoMap(vendas90){
    var map = {};
    if(!vendas90 || !vendas90.length) return map;
    vendas90.forEach(function(r){
      var sku = String(r.sku||'').trim();
      var qtd = Number(r.qtdVendida)||0;
      var cmv = Number(r.custoVendido)||0;
      if(sku && qtd > 0 && cmv > 0){
        map[sku] = round2(cmv / qtd);
      }
    });
    return map;
  }

  /* Resolve custo: do item > do mapa de vendas > 0 */
  function resolveCusto(item, custoMap){
    if(item.custoUnit && item.custoUnit > 0) return item.custoUnit;
    return custoMap[item.sku] || 0;
  }

  /* ========== 1. CRITICA ========== */
  function calcCritica(estoque, contagem, cadastro, custoMap){
    custoMap = custoMap || {};
    var skuMap = {};
    estoque.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      skuMap[sku] = {sku:sku, descricao:row.descricao||'', categoria:row.categoria||'', qtdSistema:Number(row.qtdSistema)||0, custoUnit:Number(row.custoUnit)||0};
    });
    contagem.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      if(!skuMap[sku]) skuMap[sku] = {sku:sku, descricao:row.descricao||'', categoria:row.categoria||'', qtdSistema:0, custoUnit:Number(row.custoUnit)||0};
      var item = skuMap[sku];
      item.qtdContada = (item.qtdContada||0) + (Number(row.qtdContada)||0);
      if(row.descricao && !item.descricao) item.descricao = row.descricao;
      if(row.categoria && !item.categoria) item.categoria = row.categoria;
      if(row.custoUnit && !item.custoUnit) item.custoUnit = Number(row.custoUnit)||0;
    });
    if(cadastro && cadastro.length){
      var cadMap = {};
      cadastro.forEach(function(r){ cadMap[String(r.sku||'').trim()] = r; });
      Object.keys(skuMap).forEach(function(sku){
        var c = cadMap[sku];
        if(c){
          if(c.descricao && !skuMap[sku].descricao) skuMap[sku].descricao = c.descricao;
          if(c.categoria && !skuMap[sku].categoria) skuMap[sku].categoria = c.categoria;
          if(c.custoUnit && !skuMap[sku].custoUnit) skuMap[sku].custoUnit = Number(c.custoUnit)||0;
        }
      });
    }
    var items = [];
    Object.keys(skuMap).forEach(function(sku){
      var it = skuMap[sku];
      it.qtdContada = it.qtdContada || 0;
      it.custoUnit = resolveCusto(it, custoMap);
      it.difQtd = it.qtdContada - it.qtdSistema;
      it.difValor = round2(it.difQtd * it.custoUnit);
      it.status = it.difQtd < 0 ? 'Falta' : (it.difQtd > 0 ? 'Sobra' : 'OK');
      items.push(it);
    });
    // Sort: maior falta financeira primeiro
    items.sort(function(a,b){ return a.difValor - b.difValor; });
    var totalSKUs = items.length;
    var okCount = items.filter(function(i){return i.status==='OK'}).length;
    var faltaItems = items.filter(function(i){return i.status==='Falta'});
    var sobraItems = items.filter(function(i){return i.status==='Sobra'});
    var totalFaltas = faltaItems.reduce(function(s,i){return s+i.difValor},0);
    var totalSobras = sobraItems.reduce(function(s,i){return s+i.difValor},0);
    var catList = groupByCategoria(items, function(g){
      var ok=g.items.filter(function(i){return i.status==='OK'}).length;
      var fv=g.items.filter(function(i){return i.status==='Falta'}).reduce(function(s,i){return s+i.difValor},0);
      var sv=g.items.filter(function(i){return i.status==='Sobra'}).reduce(function(s,i){return s+i.difValor},0);
      return {nome:g.nome, total:g.items.length, ok:ok, faltaVal:round2(fv), sobraVal:round2(sv), acuracidade:g.items.length?round2(ok/g.items.length*100):0, saldo:round2(fv+sv), destaque:Math.abs(fv)};
    });
    return {items:items, totalSKUs:totalSKUs, okCount:okCount, acuracidade:totalSKUs?round2(okCount/totalSKUs*100):0, faltaCount:faltaItems.length, sobraCount:sobraItems.length, totalFaltas:round2(totalFaltas), totalSobras:round2(totalSobras), saldoLiquido:round2(totalFaltas+totalSobras), categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  /* ========== ABC helper ========== */
  function calcABC(items, valueField){
    var sorted = items.slice().filter(function(i){return (i[valueField]||0)>0;});
    sorted.sort(function(a,b){return (b[valueField]||0)-(a[valueField]||0);});
    var total = sorted.reduce(function(s,i){return s+(i[valueField]||0);},0);
    var cum = 0;
    sorted.forEach(function(it){
      cum += (it[valueField]||0);
      it['abc_'+valueField] = (total?cum/total*100:0)<=80?'A':((total?cum/total*100:0)<=95?'B':'C');
    });
    items.forEach(function(it){if(!it['abc_'+valueField]) it['abc_'+valueField]='C';});
    return items;
  }

  function isDeposito(local){
    var d=['deposito','depósito','dep','dep.','retaguarda','cd','estoque','armazem','armazém','back','reserva'];
    for(var i=0;i<d.length;i++) if(local.indexOf(d[i])>=0) return true;
    return false;
  }

  /* ========== 2. RUPTURA ========== */
  function calcRuptura(contagem, vendas90, cadastro, diasVenda){
    diasVenda = diasVenda || 90;
    var catMap={},descMap={};
    if(cadastro&&cadastro.length){cadastro.forEach(function(r){var s=String(r.sku||'').trim();if(s&&r.categoria)catMap[s]=r.categoria;if(s&&r.descricao)descMap[s]=r.descricao;});}
    if(vendas90&&vendas90.length){vendas90.forEach(function(r){var s=String(r.sku||'').trim();if(s&&r.categoria&&!catMap[s])catMap[s]=r.categoria;});}
    var skuLocais = {};
    contagem.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      if(!skuLocais[sku]) skuLocais[sku] = {sku:sku, descricao:row.descricao||descMap[sku]||'', categoria:row.categoria||catMap[sku]||'', deposito:0, loja:0, custoUnit:Number(row.custoUnit)||0};
      var local = String(row.local||'').toLowerCase().trim();
      var qty = Number(row.qtdContada)||0;
      if(isDeposito(local)) skuLocais[sku].deposito += qty;
      else skuLocais[sku].loja += qty;
      if(row.descricao) skuLocais[sku].descricao = row.descricao;
      if(row.categoria) skuLocais[sku].categoria = row.categoria;
    });
    var vendasMap = {};
    if(vendas90 && vendas90.length){
      vendas90.forEach(function(r){
        var sku = String(r.sku||'').trim();
        vendasMap[sku] = {qtdVendida:Number(r.qtdVendida)||0, valorVendido:Number(r.valorVendido)||0, custoVendido:Number(r.custoVendido)||0, lucro:Number(r.lucro)||0, categoria:r.categoria||''};
      });
    }
    var comDeposito=[], rupturas=[];
    Object.keys(skuLocais).forEach(function(sku){
      var it = skuLocais[sku];
      if(it.deposito > 0){
        var v = vendasMap[sku]||{};
        it.vendaMediaDia = safeDiv(v.qtdVendida,diasVenda);
        it.fatMediaDia = safeDiv(v.valorVendido,diasVenda);
        it.lucroMediaDia = safeDiv(v.lucro||(v.valorVendido-v.custoVendido),diasVenda);
        it.valorVendido90 = v.valorVendido||0;
        it.lucro90 = v.lucro||((v.valorVendido||0)-(v.custoVendido||0));
        if(!it.categoria) it.categoria = catMap[sku] || v.categoria || '';
        // Custo do CMV
        var custoU = it.custoUnit || (v.qtdVendida>0 ? round2(v.custoVendido/v.qtdVendida) : 0);
        it.custoUnit = custoU;
        it.valorEstoque = round2((it.deposito + it.loja) * custoU);
        comDeposito.push(it);
        if(it.loja===0) rupturas.push(it);
      }
    });
    calcABC(comDeposito,'valorVendido90');
    calcABC(comDeposito,'lucro90');
    // Sort rupturas: maior valor de estoque primeiro
    rupturas.sort(function(a,b){ return (b.valorEstoque||0)-(a.valorEstoque||0); });
    var ruptA=rupturas.filter(function(i){return i.abc_valorVendido90==='A'});
    var comDepA=comDeposito.filter(function(i){return i.abc_valorVendido90==='A'});
    var catList = groupByCategoria(rupturas, function(g){
      var dep=comDeposito.filter(function(i){return(i.categoria||'Sem categoria')===g.nome});
      var rA=g.items.filter(function(i){return i.abc_valorVendido90==='A'}).length;
      var taxa=dep.length?round2(g.items.length/dep.length*100):0;
      var perdaDia=g.items.reduce(function(s,i){return s+(i.fatMediaDia||0);},0);
      return {nome:g.nome, totalRupturas:g.items.length, totalDeposito:dep.length, taxa:taxa, rupturaA:rA, perdaDia:round2(perdaDia), destaque:g.items.length};
    });
    return {items:rupturas, allDeposito:comDeposito, totalComDeposito:comDeposito.length, totalRupturas:rupturas.length,
      taxaRuptura:comDeposito.length?round2(rupturas.length/comDeposito.length*100):0,
      rupturaA:ruptA.length, rupturaB:rupturas.filter(function(i){return i.abc_valorVendido90==='B'}).length,
      rupturaC:rupturas.filter(function(i){return i.abc_valorVendido90==='C'}).length,
      taxaA:comDepA.length?round2(ruptA.length/comDepA.length*100):0,
      taxaALucro:(function(){var cA=comDeposito.filter(function(i){return i.abc_lucro90==='A'});var rA2=rupturas.filter(function(i){return i.abc_lucro90==='A'});return cA.length?round2(rA2.length/cA.length*100):0;})(),
      comDepA:comDepA.length, categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  /* ========== 3. DIAS DE ESTOQUE ========== */
  function calcDiasEstoque(critica, vendas90, custoMap, diasVenda, cadastro){
    diasVenda = diasVenda || 90;
    custoMap = custoMap || {};
    var vendasMap = {};
    if(vendas90 && vendas90.length){
      vendas90.forEach(function(r){
        var sku=String(r.sku||'').trim();
        if(!vendasMap[sku]) vendasMap[sku]={qtdVendida:0, valorVendido:0, custoVendido:0, descricao:'', categoria:''};
        vendasMap[sku].qtdVendida += (Number(r.qtdVendida)||0);
        vendasMap[sku].valorVendido += (Number(r.valorVendido)||0);
        vendasMap[sku].custoVendido += (Number(r.custoVendido)||0);
        if(r.descricao && !vendasMap[sku].descricao) vendasMap[sku].descricao = r.descricao;
        if(r.categoria && !vendasMap[sku].categoria) vendasMap[sku].categoria = r.categoria;
      });
    }
    /* Mapa de categorias do cadastro */
    var catMap = {};
    if(cadastro && cadastro.length){
      cadastro.forEach(function(r){
        var sku=String(r.sku||'').trim();
        if(sku && r.categoria) catMap[sku] = r.categoria;
      });
    }
    /* Universo base: itens da crítica/contagem COM MOVIMENTAÇÃO */
    var skuSet = {};
    var items = critica.items.filter(function(it){
      /* Só inclui se teve movimentação real */
      if((it.qtdContada||0)>0 || (it.qtdSistema||0)>0) return true;
      var v=vendasMap[it.sku];
      if(v && v.qtdVendida>0) return true;
      return false;
    }).map(function(it){
      skuSet[it.sku] = true;
      var v = vendasMap[it.sku]||{};
      var vendaMediaDia = safeDiv(v.qtdVendida,diasVenda);
      var dias = vendaMediaDia ? roundInt(it.qtdContada/vendaMediaDia) : null;
      var custo = resolveCusto(it, custoMap);
      var faixa = dias===null?'Sem giro':(dias<=2?'Ruptura':(dias<=5?'Alto risco':(dias<=15?'Médio risco':(dias<=30?'Cobertura ideal':'Excesso de cobertura'))));
      return {sku:it.sku, descricao:it.descricao, categoria:it.categoria||catMap[it.sku]||'', qtdEstoque:it.qtdContada, vendaMediaDia:vendaMediaDia, diasEstoque:dias, faixa:faixa, valorEstoque:round2(it.qtdContada*custo), custoUnit:custo, valorVendido90:v.valorVendido||0};
    });
    /* Acrescentar SKUs vendidos que não estão na crítica/contagem */
    Object.keys(vendasMap).forEach(function(sku){
      if(skuSet[sku]) return;
      var v = vendasMap[sku];
      if(v.qtdVendida <= 0) return;
      var vendaMediaDia = safeDiv(v.qtdVendida,diasVenda);
      var custo = custoMap[sku] || (v.qtdVendida ? round2(v.custoVendido/v.qtdVendida) : 0);
      items.push({sku:sku, descricao:v.descricao||'', categoria:v.categoria||catMap[sku]||'', qtdEstoque:0, vendaMediaDia:vendaMediaDia, diasEstoque:0, faixa:'Ruptura', valorEstoque:0, custoUnit:custo, valorVendido90:v.valorVendido||0});
    });
    items = calcABC(items,'valorVendido90');
    items.forEach(function(it){it.abcFat=it.abc_valorVendido90||'C';});
    items.sort(function(a,b){ return (b.diasEstoque||0)-(a.diasEstoque||0); });
    var semGiro=items.filter(function(i){return i.faixa==='Sem giro'}).length;
    var ruptura=items.filter(function(i){return i.faixa==='Ruptura'}).length;
    var altoRisco=items.filter(function(i){return i.faixa==='Alto risco'}).length;
    var medioRisco=items.filter(function(i){return i.faixa==='Médio risco'}).length;
    var coberturaIdeal=items.filter(function(i){return i.faixa==='Cobertura ideal'}).length;
    var excessos=items.filter(function(i){return i.faixa==='Excesso de cobertura'}).length;
    function calcCobertura(filteredItems){
      var sE=0, sV=0;
      filteredItems.forEach(function(i){
        if(i.vendaMediaDia && i.vendaMediaDia>0){
          sE += i.qtdEstoque;
          sV += i.vendaMediaDia;
        }
      });
      return sV ? roundInt(sE/sV) : 0;
    }
    var coberturaGeral = calcCobertura(items);
    var coberturaA = calcCobertura(items.filter(function(i){return i.abcFat==='A';}));
    var coberturaB = calcCobertura(items.filter(function(i){return i.abcFat==='B';}));
    var coberturaC = calcCobertura(items.filter(function(i){return i.abcFat==='C';}));
    var valExcesso=items.filter(function(i){return i.faixa==='Excesso de cobertura'}).reduce(function(s,i){return s+i.valorEstoque},0);
    var catList = groupByCategoria(items, function(g){
      var cob = calcCobertura(g.items);
      var sg=g.items.filter(function(i){return i.faixa==='Sem giro'}).length;
      var cr=g.items.filter(function(i){return i.faixa==='Ruptura'||i.faixa==='Alto risco'}).length;
      var ex=g.items.filter(function(i){return i.faixa==='Excesso de cobertura'}).length;
      var valEst=g.items.reduce(function(s,i){return s+i.valorEstoque},0);
      return {nome:g.nome, total:g.items.length, mediaCobertura:cob, semGiro:sg, criticos:cr, excessos:ex, valorEstoque:round2(valEst), destaque:cr+sg};
    });
    return {items:items, coberturaGeral:coberturaGeral, coberturaA:coberturaA, coberturaB:coberturaB, coberturaC:coberturaC, semGiro:semGiro, ruptura:ruptura, altoRisco:altoRisco, medioRisco:medioRisco, coberturaIdeal:coberturaIdeal, excessos:excessos, valorExcesso:round2(valExcesso), total:items.length, categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  /* ========== 4. INVESTIMENTO ABC ========== */
  function calcInvestimentoABC(critica, vendas90, custoMap, diasVenda){
    diasVenda = diasVenda || 90;
    custoMap = custoMap || {};
    var vendasMap = {};
    vendas90.forEach(function(r){
      var sku=String(r.sku||'').trim();
      vendasMap[sku]={valorVendido:Number(r.valorVendido)||0, custoVendido:Number(r.custoVendido)||0, lucro:Number(r.lucro)||0, qtdVendida:Number(r.qtdVendida)||0};
    });
    var items = critica.items.map(function(it){
      var v=vendasMap[it.sku]||{};
      var lucro=v.lucro||((v.valorVendido||0)-(v.custoVendido||0));
      if(!lucro&&v.valorVendido&&it.custoUnit) lucro=v.valorVendido-(v.qtdVendida*it.custoUnit);
      var custo = resolveCusto(it, custoMap);
      return {sku:it.sku, descricao:it.descricao, categoria:it.categoria, qtdEstoque:it.qtdContada, custoUnit:custo, valorInvestido:round2(it.qtdContada*custo), fat90:v.valorVendido||0, lucro90:round2(lucro), qtdVendida90:v.qtdVendida||0};
    });
    items=calcABC(items,'fat90');
    items=calcABC(items,'lucro90');
    items.forEach(function(it){it.abcFat=it.abc_fat90||'C';it.abcLucro=it.abc_lucro90||'C';});
    // Sort: maior valor investido primeiro
    items.sort(function(a,b){ return (b.valorInvestido||0)-(a.valorInvestido||0); });
    function agg(cls,f){return items.filter(function(i){return i['abc_'+f]===cls}).reduce(function(s,i){return s+i.valorInvestido},0);}
    function aggF(cls,f,sf){return items.filter(function(i){return i['abc_'+f]===cls}).reduce(function(s,i){return s+i[sf]},0);}
    var totalInvest=items.reduce(function(s,i){return s+i.valorInvestido},0);
    var totalFat=items.reduce(function(s,i){return s+i.fat90},0);
    var totalLucro=items.reduce(function(s,i){return s+i.lucro90},0);
    var catList = groupByCategoria(items, function(g){
      var inv=g.items.reduce(function(s,i){return s+i.valorInvestido},0);
      var fat=g.items.reduce(function(s,i){return s+i.fat90},0);
      var luc=g.items.reduce(function(s,i){return s+i.lucro90},0);
      return {nome:g.nome, total:g.items.length, investimento:round2(inv), faturamento:round2(fat), lucro:round2(luc), pctInvest:totalInvest?round2(inv/totalInvest*100):0, destaque:inv};
    });
    return {items:items, totalInvest:round2(totalInvest), totalFat:round2(totalFat), totalLucro:round2(totalLucro),
      fatA:{invest:round2(agg('A','fat90')),fat:round2(aggF('A','fat90','fat90')),pctInvest:totalInvest?round2(agg('A','fat90')/totalInvest*100):0,pctFat:totalFat?round2(aggF('A','fat90','fat90')/totalFat*100):0},
      fatB:{invest:round2(agg('B','fat90')),fat:round2(aggF('B','fat90','fat90')),pctInvest:totalInvest?round2(agg('B','fat90')/totalInvest*100):0,pctFat:totalFat?round2(aggF('B','fat90','fat90')/totalFat*100):0},
      fatC:{invest:round2(agg('C','fat90')),fat:round2(aggF('C','fat90','fat90')),pctInvest:totalInvest?round2(agg('C','fat90')/totalInvest*100):0,pctFat:totalFat?round2(aggF('C','fat90','fat90')/totalFat*100):0},
      lucA:{invest:round2(agg('A','lucro90')),luc:round2(aggF('A','lucro90','lucro90')),pctInvest:totalInvest?round2(agg('A','lucro90')/totalInvest*100):0,pctLuc:totalLucro?round2(aggF('A','lucro90','lucro90')/totalLucro*100):0},
      lucB:{invest:round2(agg('B','lucro90')),luc:round2(aggF('B','lucro90','lucro90')),pctInvest:totalInvest?round2(agg('B','lucro90')/totalInvest*100):0,pctLuc:totalLucro?round2(aggF('B','lucro90','lucro90')/totalLucro*100):0},
      lucC:{invest:round2(agg('C','lucro90')),luc:round2(aggF('C','lucro90','lucro90')),pctInvest:totalInvest?round2(agg('C','lucro90')/totalInvest*100):0,pctLuc:totalLucro?round2(aggF('C','lucro90','lucro90')/totalLucro*100):0},
      categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  /* ========== 5. PROJECAO DE PERDA ========== */
  function calcProjecaoPerda(vendas90, contagem, cadastro, diasVenda){
    diasVenda = diasVenda || 90;
    /* Mapa de contagem: SKU → qtd total contada */
    var contagemMap = {};
    contagem.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      contagemMap[sku] = (contagemMap[sku]||0) + (Number(row.qtdContada)||0);
    });
    /* Mapa de categorias do cadastro */
    var catMap = {};
    if(cadastro && cadastro.length){
      cadastro.forEach(function(r){
        var sku = String(r.sku||'').trim();
        if(sku && r.categoria) catMap[sku] = r.categoria;
      });
    }
    /* Agrupar vendas por SKU */
    var vendasMap = {};
    vendas90.forEach(function(r){
      var sku = String(r.sku||'').trim();
      if(!sku) return;
      if(!vendasMap[sku]) vendasMap[sku] = {sku:sku, descricao:r.descricao||'', categoria:r.categoria||catMap[sku]||'',
        qtdVendida:0, valorVendido:0, custoVendido:0, lucro:0};
      vendasMap[sku].qtdVendida += (Number(r.qtdVendida)||0);
      vendasMap[sku].valorVendido += (Number(r.valorVendido)||0);
      var cv = Number(r.custoVendido)||0;
      var lc = Number(r.lucro)||0;
      vendasMap[sku].custoVendido += cv;
      vendasMap[sku].lucro += lc;
      if(r.descricao && !vendasMap[sku].descricao) vendasMap[sku].descricao = r.descricao;
      if(r.categoria && !vendasMap[sku].categoria) vendasMap[sku].categoria = r.categoria;
    });
    /* Filtrar: SKUs vendidos que NÃO constam na contagem (ou qtd contada = 0) */
    var items = [];
    Object.keys(vendasMap).forEach(function(sku){
      var v = vendasMap[sku];
      var qtdContada = contagemMap[sku] || 0;
      if(qtdContada > 0) return; /* tem estoque físico → não é perda */
      if(v.qtdVendida <= 0) return; /* sem venda → não projeta */
      var vendaMediaDia = round2(v.qtdVendida / diasVenda);
      var fatMediaDia = round2(v.valorVendido / diasVenda);
      var lucroMediaDia = round2(v.lucro / diasVenda);
      items.push({sku:sku, descricao:v.descricao, categoria:v.categoria || catMap[sku] || '',
        abcFat:'C', abcLucro:'C',
        qtdVendida:v.qtdVendida, vendaMediaDia:vendaMediaDia,
        fatMediaDia:fatMediaDia, lucroMediaDia:lucroMediaDia,
        perdaFatDia:fatMediaDia, perdaLucroDia:lucroMediaDia,
        perdaFatMes:round2(fatMediaDia*30), perdaLucroMes:round2(lucroMediaDia*30)});
    });
    /* Classificar ABC por faturamento e lucro */
    items = calcABC(items, 'perdaFatDia');
    items.forEach(function(it){ it.abcFat = it.abc_perdaFatDia || 'C'; });
    items = calcABC(items, 'perdaLucroDia');
    items.forEach(function(it){ it.abcLucro = it.abc_perdaLucroDia || 'C'; });
    /* Ordenar por maior perda */
    items.sort(function(a,b){ return (b.perdaFatDia||0)-(a.perdaFatDia||0); });
    var totalPerdaFat=round2(items.reduce(function(s,i){return s+i.perdaFatDia},0));
    var totalPerdaLucro=round2(items.reduce(function(s,i){return s+i.perdaLucroDia},0));
    function byAbc(cls,f){return round2(items.filter(function(i){return i.abcFat===cls}).reduce(function(s,i){return s+i[f]},0));}
    function cntAbc(cls){return items.filter(function(i){return i.abcFat===cls}).length;}
    var catList = groupByCategoria(items, function(g){
      var pf=g.items.reduce(function(s,i){return s+i.perdaFatDia},0);
      var pl=g.items.reduce(function(s,i){return s+i.perdaLucroDia},0);
      var rA=g.items.filter(function(i){return i.abcFat==='A'}).length;
      return {nome:g.nome, totalRupturas:g.items.length, perdaFatDia:round2(pf), perdaLucroDia:round2(pl), perdaMensal:round2(pf*30), rupturaA:rA, destaque:pf};
    });
    return {items:items, totalPerdaFat:totalPerdaFat, totalPerdaLucro:totalPerdaLucro, perdaMensal:round2(totalPerdaFat*30), totalSKUs:items.length,
      classA:{perda:byAbc('A','perdaFatDia'),lucro:byAbc('A','perdaLucroDia'),count:cntAbc('A'),pct:totalPerdaFat?round2(byAbc('A','perdaFatDia')/totalPerdaFat*100):0},
      classB:{perda:byAbc('B','perdaFatDia'),lucro:byAbc('B','perdaLucroDia'),count:cntAbc('B'),pct:totalPerdaFat?round2(byAbc('B','perdaFatDia')/totalPerdaFat*100):0},
      classC:{perda:byAbc('C','perdaFatDia'),lucro:byAbc('C','perdaLucroDia'),count:cntAbc('C'),pct:totalPerdaFat?round2(byAbc('C','perdaFatDia')/totalPerdaFat*100):0},
      categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  function buildItemsFromContagem(contagem, cadastro){
    var skuMap = {};
    contagem.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      if(!skuMap[sku]) skuMap[sku] = {sku:sku, descricao:row.descricao||'', categoria:row.categoria||'', qtdContada:0, custoUnit:Number(row.custoUnit)||0, qtdSistema:0, difQtd:0, difValor:0, status:'—'};
      skuMap[sku].qtdContada += (Number(row.qtdContada)||0);
      if(row.descricao && !skuMap[sku].descricao) skuMap[sku].descricao = row.descricao;
      if(row.categoria && !skuMap[sku].categoria) skuMap[sku].categoria = row.categoria;
    });
    if(cadastro && cadastro.length){
      var cadMap = {};
      cadastro.forEach(function(r){ cadMap[String(r.sku||'').trim()] = r; });
      Object.keys(skuMap).forEach(function(sku){
        var c = cadMap[sku];
        if(c){
          if(c.descricao && !skuMap[sku].descricao) skuMap[sku].descricao = c.descricao;
          if(c.categoria && !skuMap[sku].categoria) skuMap[sku].categoria = c.categoria;
        }
      });
    }
    return Object.keys(skuMap).map(function(k){ return skuMap[k]; });
  }

  return {calcCritica:calcCritica, calcRuptura:calcRuptura, calcDiasEstoque:calcDiasEstoque, calcInvestimentoABC:calcInvestimentoABC, calcProjecaoPerda:calcProjecaoPerda, calcABC:calcABC, buildItemsFromContagem:buildItemsFromContagem, buildCustoMap:buildCustoMap, resolveCusto:resolveCusto, round2:round2, roundInt:roundInt};
})();
