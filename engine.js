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
  function calcCritica(estoque, contagem, cadastro, custoMap, vendas90){
    custoMap = custoMap || {};
    vendas90 = vendas90 || [];
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
    /* Mapa de vendas: SKU -> qtd vendida total, usado só para qualificar o universo da acuracidade */
    var vendaQtdMap = {};
    vendas90.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      vendaQtdMap[sku] = (vendaQtdMap[sku]||0) + (Number(row.qtdVendida)||0);
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
      /* Universo da acuracidade: só entra se estoque de sistema for positivo/negativo,
         a contagem física for maior que zero, ou houver venda registrada maior que zero.
         Qualquer produto fora dessa qualificação é removido de toda a análise
         (inclusive Dias de Estoque e Investimento ABC, que partem destes mesmos items). */
      var qualifica = it.qtdSistema !== 0 || it.qtdContada > 0 || (vendaQtdMap[sku]||0) > 0;
      if(!qualifica) return;
      it.custoUnit = resolveCusto(it, custoMap);
      it.difQtd = round2(it.qtdContada - it.qtdSistema);
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
    /* Acuracidade = valor do estoque contado / valor do estoque antes da contagem (base: valor, não contagem de SKUs) */
    function acuraciaPorValor(valEstoque, valContado){ return valEstoque ? round2(valContado/valEstoque*100) : 0; }
    var catList = groupByCategoria(items, function(g){
      var ok=g.items.filter(function(i){return i.status==='OK'}).length;
      var fv=g.items.filter(function(i){return i.status==='Falta'}).reduce(function(s,i){return s+i.difValor},0);
      var sv=g.items.filter(function(i){return i.status==='Sobra'}).reduce(function(s,i){return s+i.difValor},0);
      var vE=g.items.reduce(function(s,i){return s+(i.qtdSistema*i.custoUnit);},0);
      var vC=g.items.reduce(function(s,i){return s+(i.qtdContada*i.custoUnit);},0);
      return {nome:g.nome, total:g.items.length, ok:ok, faltaVal:round2(fv), sobraVal:round2(sv), valorEstoque:round2(vE), valorEstoqueContado:round2(vC), acuracidade:acuraciaPorValor(vE,vC), saldo:round2(fv+sv), destaque:Math.abs(fv)};
    });
    /* KPI Perda de Estoque % = (valor estoque - valor estoque contado) / valor estoque * -1 */
    var valorEstoque = items.reduce(function(s,i){return s+(i.qtdSistema*i.custoUnit);},0);
    var valorEstoqueContado = items.reduce(function(s,i){return s+(i.qtdContada*i.custoUnit);},0);
    var perdaEstoquePct = valorEstoque ? round2((valorEstoque - valorEstoqueContado) / valorEstoque * -1 * 100) : 0;
    return {items:items, totalSKUs:totalSKUs, okCount:okCount, acuracidade:acuraciaPorValor(valorEstoque,valorEstoqueContado), faltaCount:faltaItems.length, sobraCount:sobraItems.length, totalFaltas:round2(totalFaltas), totalSobras:round2(totalSobras), saldoLiquido:round2(totalFaltas+totalSobras), valorEstoque:round2(valorEstoque), valorEstoqueContado:round2(valorEstoqueContado), perdaEstoquePct:perdaEstoquePct, categorias:catList, hasCategorias:hasRealCategorias(catList)};
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

  /* ========== 6. COMPARATIVO ENTRE UNIDADES ========== */
  /* Recebe array de objetos {unidade:'nome', results:{critica,ruptura,dias,abc,perda}, info:{...}}
     Retorna objeto com métricas lado a lado, rankings e overlap de SKUs. */
  function calcComparativo(unidades){
    if(!unidades || unidades.length < 2) return null;
    var comp = {unidades:[], rankings:[], skuOverlap:null};

    var rows = [];
    unidades.forEach(function(u){
      var r = u.results || {};
      rows.push({
        unidade: u.unidade || u.info.unidade || '?',
        data: u.info.dataInventario || '',
        /* Crítica */
        acuracidade: r.critica ? r.critica.acuracidade : null,
        totalSKUs: r.critica ? r.critica.totalSKUs : null,
        faltaCount: r.critica ? r.critica.faltaCount : null,
        sobraCount: r.critica ? r.critica.sobraCount : null,
        totalFaltas: r.critica ? r.critica.totalFaltas : null,
        totalSobras: r.critica ? r.critica.totalSobras : null,
        saldoLiquido: r.critica ? r.critica.saldoLiquido : null,
        /* Ruptura */
        taxaRuptura: r.ruptura ? r.ruptura.taxaRuptura : null,
        totalRupturas: r.ruptura ? r.ruptura.totalRupturas : null,
        rupturaA: r.ruptura ? r.ruptura.rupturaA : null,
        /* Dias de estoque */
        coberturaGeral: r.dias ? r.dias.coberturaGeral : null,
        coberturaA: r.dias ? r.dias.coberturaA : null,
        semGiro: r.dias ? r.dias.semGiro : null,
        valorExcesso: r.dias ? r.dias.valorExcesso : null,
        /* ABC */
        totalInvest: r.abc ? r.abc.totalInvest : null,
        totalFat: r.abc ? r.abc.totalFat : null,
        pctInvestA: r.abc ? r.abc.fatA.pctInvest : null,
        pctFatA: r.abc ? r.abc.fatA.pctFat : null,
        /* Perda */
        perdaFatDia: r.perda ? r.perda.totalPerdaFat : null,
        perdaMensal: r.perda ? r.perda.perdaMensal : null,
        perdaSKUs: r.perda ? r.perda.totalSKUs : null,
        /* Categorias por dimensão */
        categoriasCritica: r.critica ? r.critica.categorias : [],
        categoriasRuptura: r.ruptura ? r.ruptura.categorias : [],
        categoriasDias: r.dias ? r.dias.categorias : [],
        categoriasPerda: r.perda ? r.perda.categorias : []
      });
    });
    comp.unidades = rows;

    /* Rankings por métrica */
    var metricasDefs = [
      {key:'acuracidade',   label:'Acuracidade (%)',          melhor:'max',      fmt:'pct'},
      {key:'totalFaltas',   label:'Valor Faltas (R$)',        melhor:'min_abs',  fmt:'brl'},
      {key:'saldoLiquido',  label:'Saldo Líquido (R$)',       melhor:'min_abs',  fmt:'brl'},
      {key:'taxaRuptura',   label:'Taxa de Ruptura (%)',      melhor:'min',      fmt:'pct'},
      {key:'totalRupturas', label:'Itens em Ruptura',         melhor:'min',      fmt:'num'},
      {key:'rupturaA',      label:'Rupturas Curva A',         melhor:'min',      fmt:'num'},
      {key:'coberturaGeral',label:'Cobertura (dias)',          melhor:'target30', fmt:'num'},
      {key:'coberturaA',    label:'Cobertura Curva A (dias)',  melhor:'target15', fmt:'num'},
      {key:'semGiro',       label:'Itens Sem Giro',           melhor:'min',      fmt:'num'},
      {key:'valorExcesso',  label:'Valor em Excesso (R$)',    melhor:'min',      fmt:'brl'},
      {key:'totalInvest',   label:'Valor Total Estoque (R$)', melhor:'info',     fmt:'brl'},
      {key:'perdaFatDia',   label:'Perda Fat./Dia (R$)',      melhor:'min',      fmt:'brl'},
      {key:'perdaMensal',   label:'Perda Mensal (R$)',        melhor:'min',      fmt:'brl'}
    ];

    metricasDefs.forEach(function(def){
      var vals = rows.map(function(r){ return {unidade:r.unidade, valor:r[def.key]}; })
        .filter(function(v){ return v.valor !== null && v.valor !== undefined; });
      if(vals.length < 2) return;

      var sorted;
      if(def.melhor === 'max'){
        sorted = vals.slice().sort(function(a,b){ return b.valor - a.valor; });
      } else if(def.melhor === 'min_abs'){
        sorted = vals.slice().sort(function(a,b){ return Math.abs(a.valor) - Math.abs(b.valor); });
      } else if(def.melhor === 'target30' || def.melhor === 'target15'){
        var alvo = def.melhor === 'target30' ? 23 : 10;
        sorted = vals.slice().sort(function(a,b){ return Math.abs(a.valor-alvo) - Math.abs(b.valor-alvo); });
      } else {
        sorted = vals.slice().sort(function(a,b){ return a.valor - b.valor; });
      }

      comp.rankings.push({
        key: def.key, label: def.label, fmt: def.fmt,
        valores: vals,
        melhor: def.melhor !== 'info' ? sorted[0].unidade : null,
        pior:   def.melhor !== 'info' ? sorted[sorted.length-1].unidade : null
      });
    });

    /* SKU overlap entre unidades */
    if(rows.length >= 2){
      var skuSets = unidades.map(function(u){
        var set = {};
        if(u.results.critica){
          u.results.critica.items.forEach(function(it){ set[it.sku] = true; });
        }
        return {unidade: u.info.unidade, set: set, count: Object.keys(set).length};
      });
      var allSkus = {};
      skuSets.forEach(function(ss){ Object.keys(ss.set).forEach(function(s){ allSkus[s] = (allSkus[s]||0) + 1; }); });
      var totalUnique = Object.keys(allSkus).length;
      var emComum = Object.keys(allSkus).filter(function(s){ return allSkus[s] === skuSets.length; }).length;
      comp.skuOverlap = {
        totalUnique: totalUnique,
        emComum: emComum,
        pctComum: totalUnique ? round2(emComum/totalUnique*100) : 0,
        porUnidade: skuSets.map(function(s){ return {unidade:s.unidade, total:s.count}; })
      };
    }

    return comp;
  }

  /* ========== 7. HISTÓRICO TEMPORAL ========== */
  /* Recebe array de {data, results, info} da mesma unidade em datas diferentes.
     Retorna séries temporais e tendências. Será consumido pela UI na Sessão 3. */
  function calcHistorico(analises){
    if(!analises || analises.length < 2) return null;

    /* Ordenar por data (dd/mm/yyyy -> yyyymmdd) */
    var sorted = analises.slice().sort(function(a,b){
      var dA = a.info.dataInventario || '';
      var dB = b.info.dataInventario || '';
      var pA = dA.split('/'); var pB = dB.split('/');
      var nA = pA.length===3 ? (pA[2]+pA[1]+pA[0]) : dA;
      var nB = pB.length===3 ? (pB[2]+pB[1]+pB[0]) : dB;
      return nA < nB ? -1 : (nA > nB ? 1 : 0);
    });

    var series = {
      datas:[], acuracidade:[], taxaRuptura:[],
      coberturaGeral:[], perdaMensal:[], totalInvest:[]
    };

    sorted.forEach(function(a){
      var r = a.results || {};
      series.datas.push(a.info.dataInventario || '?');
      series.acuracidade.push(r.critica ? r.critica.acuracidade : null);
      series.taxaRuptura.push(r.ruptura ? r.ruptura.taxaRuptura : null);
      series.coberturaGeral.push(r.dias ? r.dias.coberturaGeral : null);
      series.perdaMensal.push(r.perda ? r.perda.perdaMensal : null);
      series.totalInvest.push(r.abc ? r.abc.totalInvest : null);
    });

    /* Tendência: último vs primeiro */
    var tendencias = {};
    ['acuracidade','taxaRuptura','coberturaGeral','perdaMensal','totalInvest'].forEach(function(k){
      var vals = series[k].filter(function(v){ return v !== null; });
      if(vals.length >= 2){
        var primeiro = vals[0], ultimo = vals[vals.length-1];
        var dif = ultimo - primeiro;
        /* Para taxaRuptura e perdaMensal, diminuir é bom */
        var positivo = (k === 'taxaRuptura' || k === 'perdaMensal') ? dif < 0 : dif > 0;
        tendencias[k] = {
          primeiro: primeiro, ultimo: ultimo,
          dif: round2(dif), positivo: positivo,
          estavel: Math.abs(dif) < 0.5
        };
      }
    });

    return {series:series, tendencias:tendencias, totalPeriodos:sorted.length};
  }

  // ── Análise textual elaborada da Crítica (fallback sem IA — usada na tela e no PDF) ──
  function gerarAnaliseCritica(c, info) {
    info = info || {};
    function pct(v){ return (v||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'%'; }
    function brl(v){ return (v<0?'-':'')+'R$ '+Math.abs(Math.round(v||0)).toLocaleString('pt-BR'); }
    function num(v){ return (v||0).toLocaleString('pt-BR'); }
    var paragrafos = [];

    paragrafos.push(
      'A crítica do inventário comparou o saldo do sistema com a contagem física em ' + num(c.totalSKUs) + ' SKUs qualificados para a análise'
      + ' — considerando apenas produtos com estoque de sistema positivo ou negativo, contagem física maior que zero, ou venda registrada no período; itens sem estoque, sem contagem e sem venda foram excluídos por não terem relevância para o resultado.'
      + ' O valor do estoque calculado pelo sistema, antes da contagem, era de ' + brl(c.valorEstoque) + '; a contagem física localizou ' + brl(c.valorEstoqueContado) + ' em produtos — uma acuracidade de '
      + pct(c.acuracidade) + ' medida pelo valor contado sobre o valor de sistema, e não apenas pela quantidade de SKUs batendo exatamente.'
    );

    var pFaltaSobra = 'Do total analisado, ' + num(c.faltaCount) + ' SKUs apresentaram falta (contagem física menor que o sistema), somando ' + brl(c.totalFaltas) + ' em valor não localizado, enquanto '
      + num(c.sobraCount) + ' SKUs apresentaram sobra, no valor de ' + brl(c.totalSobras) + '. O saldo líquido entre as duas pontas é de ' + brl(c.saldoLiquido);
    pFaltaSobra += c.saldoLiquido < 0
      ? (', o que confirma uma perda de estoque de ' + pct(Math.abs(c.perdaEstoquePct)) + ' sobre o valor total antes da contagem — o físico encontrado é menor que o registrado no sistema.')
      : (', ou seja, o valor físico encontrado na contagem supera o registrado no sistema.');
    paragrafos.push(pFaltaSobra);

    if (c.faltaCount > 0 && c.sobraCount > 0 && c.totalSobras) {
      var razao = Math.abs(c.totalFaltas / c.totalSobras);
      if (razao > 2) {
        paragrafos.push('O valor das faltas supera em mais do dobro o valor das sobras — padrão que costuma indicar perda operacional real (quebra, vencimento, furto) além de simples erro de contagem, e que justifica investigação prioritária nos itens de maior valor individual.');
      } else if (razao < 0.5) {
        paragrafos.push('O valor das sobras supera consideravelmente o das faltas, o que costuma apontar para falhas no lançamento de entrada de mercadoria (nota fiscal não baixada, recebimento em duplicidade) mais do que para perda física de produto.');
      }
    }

    if (c.hasCategorias && c.categorias && c.categorias.length) {
      var piores = c.categorias.filter(function(x){return x.nome!=='Sem categoria';}).slice().sort(function(a,b){return a.acuracidade-b.acuracidade;});
      if (piores.length) {
        var nomes = piores[0].nome + ' (' + pct(piores[0].acuracidade) + ')';
        if (piores.length > 1) nomes += ' e ' + piores[1].nome + ' (' + pct(piores[1].acuracidade) + ')';
        paragrafos.push('Entre as categorias analisadas, ' + nomes + ' apresentam a menor acuracidade de valor e concentram a maior fragilidade de controle — são as primeiras candidatas a recontagem e revisão dos processos de entrada e baixa de mercadoria.');
      }
    }

    if (c.acuracidade < 90) {
      paragrafos.push('Com acuracidade de valor abaixo de 90%, recomenda-se recontagem completa do estoque, revisão dos processos de entrada e saída de mercadoria e auditoria pontual no sistema de gestão — o volume de divergência já compromete a confiabilidade do estoque contábil.');
    } else if (c.acuracidade < 95) {
      paragrafos.push('Com acuracidade de valor entre 90% e 95%, uma recontagem cíclica focada nos SKUs de maior divergência financeira, aliada à revisão dos procedimentos de conferência no recebimento, deve ser suficiente para aproximar o estoque físico do contábil.');
    } else {
      paragrafos.push('Com acuracidade de valor acima de 95%, o controle de estoque está dentro de um padrão saudável — o recomendável é manter os processos atuais e monitorar a acuracidade nas próximas contagens para evitar deterioração gradual.');
    }

    return paragrafos.join('\n\n');
  }

  // ── r68: Geração de recomendações baseadas nos dados ──
  function gerarRecomendacoes(results, info) {
    var recs = [];
    var diasVenda = info && info.diasVenda ? info.diasVenda : 30;

    // 1. Acuracidade / Crítica
    if (results.critica) {
      var c = results.critica;
      if (c.acuracidade < 90) {
        recs.push({dim:'critica', prioridade:1, texto:'Acuracidade crítica de ' + c.acuracidade + '%. Recomenda-se recontagem completa com foco nos setores de maior divergência, revisão dos processos de entrada e saída de mercadorias e auditoria no sistema de gestão de estoque.'});
      } else if (c.acuracidade < 95) {
        recs.push({dim:'critica', prioridade:2, texto:'Acuracidade de ' + c.acuracidade + '% — abaixo da meta de 95%. Recomenda-se recontagem cíclica focada nos SKUs com divergência e revisão dos procedimentos de conferência no recebimento.'});
      } else {
        recs.push({dim:'critica', prioridade:3, texto:'Acuracidade de ' + c.acuracidade + '% está dentro da meta. Manter os processos atuais e monitorar para garantir consistência nas próximas contagens.'});
      }
      if (c.faltaCount > c.sobraCount * 2) {
        recs.push({dim:'critica', prioridade:1, texto:'Faltas superam sobras em mais do dobro (' + c.faltaCount + ' faltas vs ' + c.sobraCount + ' sobras). Investigar possíveis perdas operacionais, furtos ou falhas de registro no sistema.'});
      }
    }

    // 2. Ruptura
    if (results.ruptura) {
      var r = results.ruptura;
      if (r.taxaRuptura > 10) {
        recs.push({dim:'ruptura', prioridade:1, texto:'Taxa de ruptura elevada: ' + r.taxaRuptura + '% (' + r.totalRupturas + ' SKUs). Ação imediata necessária: revisar processo de reposição, frequência de pedidos e lead time de fornecedores para os itens em falta.'});
      } else if (r.taxaRuptura > 5) {
        recs.push({dim:'ruptura', prioridade:2, texto:'Taxa de ruptura de ' + r.taxaRuptura + '% requer atenção. Priorizar reposição dos ' + r.totalRupturas + ' SKUs zerados, especialmente os da curva A.'});
      } else if (r.taxaRuptura > 0) {
        recs.push({dim:'ruptura', prioridade:3, texto:'Taxa de ruptura controlada em ' + r.taxaRuptura + '%. Manter monitoramento e garantir reposição dos ' + r.totalRupturas + ' itens identificados.'});
      }
      if (r.rupturaA > 0) {
        recs.push({dim:'ruptura', prioridade:1, texto:r.rupturaA + ' itens da curva A em ruptura — produtos que representam a maior parcela do faturamento. Reposição prioritária e urgente destes SKUs.'});
      }
    }

    // 3. Dias de Estoque / Cobertura
    if (results.dias) {
      var d = results.dias;
      if (d.coberturaGeral > diasVenda * 2) {
        recs.push({dim:'dias', prioridade:2, texto:'Cobertura média de ' + round2(d.coberturaGeral) + ' dias — mais que o dobro do período de venda (' + diasVenda + ' dias). Estoque excessivo eleva custos de armazenagem. Avaliar redução nos pedidos de reposição.'});
      } else if (d.coberturaGeral < 15) {
        recs.push({dim:'dias', prioridade:1, texto:'Cobertura média de apenas ' + round2(d.coberturaGeral) + ' dias — risco de desabastecimento. Reforçar pedidos de reposição para garantir continuidade.'});
      } else {
        recs.push({dim:'dias', prioridade:3, texto:'Cobertura média de ' + round2(d.coberturaGeral) + ' dias está adequada. Monitorar variações sazonais e manter equilíbrio entre estoque e demanda.'});
      }
      if (d.semGiro > 0) {
        recs.push({dim:'dias', prioridade:2, texto:d.semGiro + ' SKUs sem giro identificados (estoque parado). Avaliar promoções de liquidação, devoluções ao fornecedor ou remanejamento entre unidades.'});
      }
      if (d.excessos > 0) {
        recs.push({dim:'dias', prioridade:2, texto:d.excessos + ' SKUs com excesso de estoque. Valor imobilizado em excesso: R$ ' + formatNum(d.valorExcesso) + '. Reduzir pedidos futuros destes itens.'});
      }
    }

    // 4. Curva ABC / Investimento
    if (results.abc) {
      var a = results.abc;
      var pctA = a.totalInvest > 0 ? round2(a.investA / a.totalInvest * 100) : 0;
      if (pctA > 50) {
        recs.push({dim:'abc', prioridade:2, texto:'Curva A concentra ' + pctA + '% do investimento total (R$ ' + formatNum(a.investA) + '). Garantir que estes SKUs tenham alta acuracidade e reposição prioritária.'});
      }
      if (a.totalInvest > 0) {
        recs.push({dim:'abc', prioridade:3, texto:'Investimento total em estoque: R$ ' + formatNum(a.totalInvest) + '. Faturamento 90 dias: R$ ' + formatNum(a.totalFat) + '. Giro de estoque de ' + round2(a.totalFat / a.totalInvest) + 'x no período.'});
      }
    }

    // 5. Projeção de Perda
    if (results.perda) {
      var pe = results.perda;
      if (pe.perdaMensal > 50000) {
        recs.push({dim:'perda', prioridade:1, texto:'Projeção de perda mensal elevada: R$ ' + formatNum(pe.perdaMensal) + '. Ação urgente: investigar causas (avarias, furtos, erros de registro), reforçar controles e monitorar os ' + pe.totalSKUs + ' SKUs identificados.'});
      } else if (pe.perdaMensal > 10000) {
        recs.push({dim:'perda', prioridade:2, texto:'Projeção de perda mensal de R$ ' + formatNum(pe.perdaMensal) + '. Recomenda-se investigação focada nos SKUs de maior impacto financeiro e revisão dos processos de conferência.'});
      } else if (pe.perdaMensal > 0) {
        recs.push({dim:'perda', prioridade:3, texto:'Projeção de perda mensal de R$ ' + formatNum(pe.perdaMensal) + '. Nível dentro do aceitável, mas manter monitoramento contínuo.'});
      }
    }

    // Ordenar por prioridade (1=urgente, 3=informativo)
    recs.sort(function(a,b){ return a.prioridade - b.prioridade; });
    return recs;
  }

  function formatNum(v) {
    if (v == null || isNaN(v)) return '0';
    return Number(v).toLocaleString('pt-BR', {minimumFractionDigits:0, maximumFractionDigits:0});
  }

  return {calcCritica:calcCritica, calcRuptura:calcRuptura, calcDiasEstoque:calcDiasEstoque, calcInvestimentoABC:calcInvestimentoABC, calcProjecaoPerda:calcProjecaoPerda, calcABC:calcABC, buildItemsFromContagem:buildItemsFromContagem, buildCustoMap:buildCustoMap, resolveCusto:resolveCusto, round2:round2, roundInt:roundInt, calcComparativo:calcComparativo, calcHistorico:calcHistorico, gerarRecomendacoes:gerarRecomendacoes, gerarAnaliseCritica:gerarAnaliseCritica, formatNum:formatNum};
})();
