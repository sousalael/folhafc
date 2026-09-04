# PONTO DE SALVAMENTO — Sistema de Gestão FC
## Versão: 2026-09-03-r94

---

## 1. Visão Geral

O **Sistema de Gestão FC** é a plataforma integrada da Formula Code com três módulos: App Folha de Ponto, App Análise de Inventários e App **Análise de Preparação para Inventário** (r69 — renomeado; antes chamado "Auditoria de Operação". Arquivo, aba da planilha e nomes internos de função continuam `auditoria.html` / `Auditoria_Operacao` / `*Auditoria*` — só o texto exibido ao usuário mudou).

**Stack:** GitHub Pages (frontend) + Google Apps Script (backend) + Google Sheets (DB) + Google Drive (arquivos) + MailApp (e-mail) + Claude API Haiku (IA para relatórios)

**Deploy URL (fixo, NUNCA recriar):** `https://script.google.com/macros/s/AKfycbyHQEQanZCbRh7BHA0QIvUKdmNmtrYlDuWSYf7kZc8PCZbzfXk3qrCN7YQ0je6TpENw/exec`

**SHEET_ID:** `1vbm0Uh8Qp87svlcCKluYBV_m9euoZukB_FWltGlHnvs`

## 2. Arquivos (9 — sempre entregues juntos)

| Arquivo | Módulo |
|---|---|
| `index.html` | Folha de Ponto frontend |
| `SCRIPT_GOOGLE_FOLHA_DE_PONTO.txt` | Backend unificado |
| `auditoria.html` | Análise de Preparação para Inventário frontend |
| `relatorio.html` | Relatório despesas |
| `revisao.html` | Revisão lançamentos |
| `inventario.html` | Inventários frontend (multi-unidade + resumo) |
| `engine.js` | Motor cálculos inventário + comparativo + histórico + recomendações + textos de análise por IA |
| `export.js` | Exportação Excel/PDF/PPTX (individual + comparativo + resumo, com gráficos) |
| `app.js` | Upload, slots, interface inventário multi-unidade + resumo |

**Nota sobre versionamento (regra fixada na r88):** a cada entrega, `VERSAO_ESPERADA`/`VERSAO_SCRIPT`/`VERSAO` e o nome do zip são SEMPRE atualizados para o próximo número, independentemente de quais dos 9 arquivos mudaram de fato.

⚠️ **(r92) Falha detectada e corrigida:** entre a r90 e a r91, `VERSAO` (auditoria.html) ficou parada em `r89` enquanto `VERSAO_ESPERADA`/`VERSAO_SCRIPT` avançavam — a regra do parágrafo acima não estava sendo aplicada às 3 constantes, só a 2. Corrigido nesta versão; os 3 valores voltaram a bater. Conferir sempre os 3 antes de fechar uma entrega.

**Espelho dos arquivos-fonte no projeto (regra fixada na r89):** a cada atualização do sistema, os 9 arquivos-fonte também são gravados como docs do projeto em `codigo/<nome-do-arquivo>` — sempre substituindo a versão anterior no mesmo caminho. **Atenção:** Claude não tem uma ferramenta que escreva direto no Project Knowledge — o espelho depende do usuário anexar/substituir os arquivos manualmente na tela do Projeto após cada entrega.

## 3. Abas da Planilha

Colaboradores | Ponto | Despesas | Analise_Despesas | Analise_Inventarios | Auditoria_Operacao | Clientes_FC

Aba **Ponto** (Folha de Ponto) é na verdade 2 abas físicas: **Projetos** (`data|cliente|unidade|status`) e **Presencas** (`data|cliente|unidade|cpf|funcao|obs|timestamp|nome`). "Equipe" no vocabulário do Lael = um `Projeto`. `funcao` em Presencas assume `'SUPERVISÃO'`, `'DIRETORIA'` ou `'OPERAÇÃO'` — a linha `SUPERVISÃO` de um projeto identifica qual supervisor é o responsável por ele (usado no escopo dos contadores, ver seção 6).

## 4. Análise de Preparação para Inventário (r47–r93)

(sem alteração funcional na estrutura de auditoria em si desde a r89 — ver checkpoint r88 para o histórico completo)

**(r93) 4 correções implementadas nas diretrizes de texto do relatório** (estavam pendentes desde antes do painel de contadores — combinadas em conversa anterior, só agora codadas), todas em `montarPromptAuditoria` (`SCRIPT_GOOGLE_FOLHA_DE_PONTO.txt`):
1. **Equipe N/A (causa raiz corrigida):** o schema JSON do prompt forçava menção a equipe em 4 campos incondicionalmente (`resumo_executivo`, `analise_retaguarda`, `analise_area_vendas`, `analise_equipe`) e a regra 8 dizia "SEMPRE" analisar equipe — isso contradizia a regra 19 (nunca mencionar N/A) e fazia a IA inventar frases genéricas tipo "a equipe de pesagem está disponível para apoiar" mesmo com a dimensão marcada N/A. Agora tudo isso é condicional a `var temEquipe = notasEquipe.length > 0` (calculado a partir dos próprios critérios avaliados): sem nenhum dado de equipe, os 4 campos deixam de mencionar equipe (e `analise_equipe` nem aparece no schema), a regra 8 vira proibição explícita, e a regra 12 (calibragem) para de exigir frases de equipe.
2. **Tom do fechamento de prazo (regra 18) suavizado:** trocado o "as adequações devem ser concluídas antes do início da contagem oficial, para não comprometer o cronograma do inventário" (impositivo) por uma orientação de sugestão/parceria, com proibição explícita de formulação de ameaça/cobrança sobre prazo.
3. **Exemplo da regra 16 corrigido:** "a padronização do empilhamento por SKU único" → "...com produtos organizados por código de barras".
4. **Nova regra permanente 16c:** proibido exigir/sugerir "SKU único por pallet" (ou variações) em qualquer seção — mix de produtos por loja é grande demais pra isso ser viável na prática.

`analise_equipe` (campo do schema) foi conferido e confirmado como não consumido em nenhum lugar do frontend (`grep` zero ocorrências em todos os .html/.js) — seguro remover condicionalmente sem quebrar nada.

## 5. Análise de Inventários (r64–r94)

(sem alteração nesta versão — ver checkpoint r89 para a última mudança funcional: universo da faixa "Sem giro" em Dias de Estoque)

## 6. Folha de Ponto (r64–r94)

Header com logo (r55+). Relatório Múltiplo com filtros Período+Cliente (r53+). Botão "Análise de Preparação para Inventário" (r47+, renomeado r69).

**(r94) BUG CRÍTICO corrigido — Diretor via só as próprias despesas.** Causa: `popularFiltrosIniciais()` (pré-preenchia o filtro Supervisor com `[getSupervisorLogado()]`, "pelo menos ele mesmo") era chamada SEM `await` logo antes de `buscarDespesasGeral()` (a busca real, que já trazia todo mundo corretamente) — as duas rodavam em paralelo. Como `popularFiltrosIniciais` faz 2 chamadas de API antes de popular o filtro, ela às vezes terminava DEPOIS da busca real e sobrescrevia o filtro Supervisor (que já estava certo, com todos) de volta para "só eu mesmo" — restringindo a visão do Diretor como se fosse Supervisor. O backend (`podeVerRegistroFinanceiro`) sempre esteve correto (`if (ehDiretor) return true`) — o bug era 100% no frontend. Corrigido removendo a chamada a `popularFiltrosIniciais()` do fluxo automático — os 4 filtros agora nascem exclusivamente da busca real (MTD), sem nenhum pré-preenchimento paralelo.

**(r90–r92) Painel de Contadores.** Cards no topo das 3 abas (Ponto, Despesas, Financeiro), carregando sozinhos ao abrir a tela — sem precisar clicar em buscar.

- **Ponto:** "Equipes no período" (Projetos distintos) + "Participantes no período" (total de presenças, conta a mesma pessoa 2x se ela apareceu em 2 equipes). Seletor de período próprio (mês/ini/fim/🔍), MTD por padrão.
- **Despesas:** "Lançamentos no período" + "Valor no período" — reagem aos 4 filtros já existentes (Supervisor/Cliente/Grupo/Forma) e ao período pesquisado na "Todas as Despesas".
- **Financeiro:** "Adiantamentos" (qtd+valor, só Diretoria — mesma regra que já existia em `getAdiantamentos`), "Provisionamentos" (qtd+valor) e "Total de lançamentos". Reage ao mesmo seletor de período que já existia na aba (`fin-ini`/`fin-fim`, o mesmo do card "Despesas × Provisionamento").
- **Escopo (as 3 abas):** Diretor vê tudo da empresa; Supervisor vê só o que é dele. Em Ponto isso é calculado varrendo `Presencas` por linha `funcao='SUPERVISÃO'` do CPF logado (não existe coluna de supervisor na aba Projetos).
- Backend: `getContadoresPonto(cpf, filtros)` e `getContadoresFinanceiro(cpf, filtros)` — `filtros` opcional, cai no mês corrente (MTD) quando omitido.

**(r90) NOVO: filtros de Despesas em cascata hierárquica.** Os 4 filtros (Supervisor/Cliente/Grupo/Forma) da "Todas as Despesas":
- Carregam sozinhos ao abrir a aba (busca automática do mês corrente, sem precisar clicar 🔍).
- Só mostram valores que realmente têm despesa lançada (nunca vêm de cadastro/lista mestre).
- Cascata nos dois sentidos: marcar valor(es) numa categoria restringe as opções das outras 3, com base no que coexiste nos dados. 2+ valores marcados na mesma categoria = união entre eles pra esse cruzamento.
- **(r92) Regra de "filtro zerado":** uma categoria com **zero** opções marcadas (ex: usuário clicou em "Selecionar todas" pra desmarcar tudo antes de escolher manualmente) NÃO restringe as outras 3 categorias nem zera a tabela/cards — é tratada como "sem filtro" (mantém o resultado de antes) até o usuário marcar a primeira opção específica.
- **(r91) Regra de reset por período:** toda nova busca de período em Despesas zera o estado dos 4 filtros antes de recalcular a cascata — sem isso, seleções específicas de um período (ex: 1 cliente) tentavam se cruzar com o período novo pesquisado e, como raramente coincidem, zeravam os resultados.

**Fora do escopo desta rodada (combinado com o cliente):** aplicar filtro+cascata nos 5 modais de relatório que só têm filtro de Período hoje (Relatório de Despesas, BI de Despesas, Relatório de Múltiplos Projetos, Relatório de Provisionamento, Visão de Adiantamentos) — nenhum deles tem as categorias de filtro hoje, seria criar do zero.

**Versão exibida na tela de login** — `VERSAO_ESPERADA`/`VERSAO_SCRIPT`/`VERSAO` sincronizados em `2026-09-03-r94` / `r94` (pendente publicar — ver seção 11).

## 7. Princípios Técnicos

- `getActiveSpreadsheet()` NUNCA — sempre `openById(SHEET_ID)`
- Deploy fixo — NUNCA recriar
- Fetch sem Content-Type (CORS)
- `!important` em `display:none` proibido **como estilo inline direto no elemento** — a classe utilitária `.hidden { display:none !important; }` já existe no CSS do sistema e é a forma correta de esconder/mostrar via `classList.toggle('hidden', condição)`.
- Scores: `parseFloat()` (número)
- Datas: `Utilities.formatDate` antes de string ops; datas vindas de `<input type="date">` (formato `yyyy-MM-dd`) convertidas com `htmlDateToInt`; datas de planilha (formato `dd/mm/yyyy`) convertidas com `ddmmyyyyToInt` — os dois produzem o mesmo inteiro `yyyyMMdd`, comparáveis entre si.
- `MailApp` (não GmailApp)
- Fotos: upload separado da finalização
- `State` é alias para `Units[activeUnitIdx]` — compatibilidade total
- Resumo Executivo é 100% dinâmico — só exibe dimensões com dados calculados
- PDF de verdade a partir de PPTX: **API v3** — `Drive.Files.create(...)` + `DriveApp.getFileById(id).getAs('application/pdf')`. **`Drive.Files.insert` (API v2) não existe neste projeto — nunca usar.**
- Ao editar/reeditar um registro que guarda campos como `fotosIds`, sempre reenviar o valor atual no payload de salvamento
- **(r76) Regra permanente (Inventário):** nenhum texto gerado pelo sistema recomenda recontagem de estoque.
- **(r77) Regra permanente (Auditoria):** proibido recomendar "protocolo de comunicação entre equipes".
- **(r78) Regra permanente (Inventário):** proibido recomendar revisão de "procedimentos de contagem/lançamento".
- **(r79)** Qualquer arquivo de entrada do Inventário pode conter múltiplas linhas para o mesmo SKU — toda agregação por SKU deve somar, nunca sobrescrever.
- **(r80) Regra permanente (Auditoria):** proibido dar a entender que a contagem da FC foi afetada.
- **(r81)** `normalizeSKU()` (em `app.js`) é o único ponto de normalização de SKU do sistema.
- **(r82) Regra permanente (todo o sistema):** proibido dar a entender que a qualidade da contagem física ou a confiabilidade do estoque está comprometida.
- **(r85) Regra permanente (Auditoria):** Apresentação Executiva nunca pode divergir de relatório já publicado.
- **(r86/r87) Regra permanente (todo o sistema):** termos proibidos + reenquadramento construtivo obrigatório.
- **(r87) Regra permanente (Auditoria):** toda ação recomendada é sempre do cliente, nunca da Formula Code.
- **(r83/r88) Nomenclatura fixa:** "Ruptura Loja x Depósito".
- **(r88) Regra de versionamento (ativa):** a cada entrega, SEMPRE atualizar `VERSAO_ESPERADA` (index.html), `VERSAO_SCRIPT` (SCRIPT_GOOGLE_FOLHA_DE_PONTO.txt) e `VERSAO` (auditoria.html) e o nome do arquivo zip para o próximo número — mesmo quando a mudança está só em arquivo client-side sem essa constante. Pedido explícito do cliente. **(r92) Reforço:** essa checagem falhou de fato entre r90–r91 (só 2 das 3 constantes foram sincronizadas) — sempre grep pelas 3 antes de fechar uma entrega, não confiar só na memória.
- **(r89) Regra de universo (Dias de Estoque):** a faixa "Sem giro" é exclusiva de SKU efetivamente contado (`qtdContada>0`) e sem venda no período.
- **(r89) Regra de espelho no projeto (ativa):** a cada atualização, os 9 arquivos-fonte também são gravados/atualizados em `codigo/<nome-do-arquivo>` no projeto Claude — feito manualmente pelo usuário, Claude não escreve no Project Knowledge diretamente.
- **(r90) Boa prática de reconstrução:** ao reconstruir um arquivo grande a partir do texto lido do projeto, todo conteúdo é passado por uma etapa mecânica de decodificação (nunca "retype" manual de trechos longos e sensíveis a caractere) — e, ao final, roda-se sempre checagem de sintaxe (`new Function()`/`node --check`), chaves balanceadas, tags HTML e sequências `\"` suspeitas antes da entrega.
- **(r90) Escopo Diretor/Supervisor sem coluna dedicada:** quando uma tabela não guarda "de quem é" um registro (ex: aba Projetos), o dono é inferido por outra aba relacionada (ex: linha `funcao='SUPERVISÃO'` em Presencas) — sempre verificar o dado real antes de assumir que uma coluna existe.
- **(r90) Regra de acesso (Financeiro, já existente, documentada agora):** Adiantamento é visível SOMENTE para Diretoria (`getAdiantamentos` já bloqueia Supervisor até dos próprios); Provisionamento segue a regra padrão Diretor-vê-tudo/Supervisor-vê-o-seu. As duas regras são diferentes uma da outra — não assumir simetria.
- **(r91) Regra de cascata + período:** sempre que uma tela recalcula filtros em cascata a partir de um dataset, qualquer nova busca que troque o dataset (novo período, por exemplo) precisa **resetar o estado dos filtros antes de recalcular** — carregar seleções antigas para cruzar com dado novo tende a zerar tudo.
- **(r92) Regra de cascata + filtro zerado:** uma categoria de filtro com zero opções marcadas deve ser tratada como "sem restrição" (não como "restringe a nada") tanto no cálculo da cascata quanto no filtro final da tabela — do contrário, o simples ato de desmarcar tudo pra escolher de novo apaga as opções dos outros filtros.
- **(r94) Regra de concorrência (crítica):** nunca disparar duas funções `async` que escrevem no mesmo estado compartilhado (ex: `filtrosGeralEstado`) uma logo após a outra sem `await` — mesmo que pareçam independentes, a ordem de retorno das chamadas de API não é garantida, e a que "chegar por último" vence e sobrescreve a outra. Se uma virou redundante (como `popularFiltrosIniciais` ficou depois do auto-carregamento real), o caminho mais seguro é remover a chamada, não tentar sincronizar as duas.
- **Padrão consolidado:** todo texto gerado por IA ou fallback local deve ter a atribuição de responsabilidade e a ação corretiva corretas explícitas no prompt de sistema.
- **Atenção a versões concorrentes:** antes de entregar uma nova versão, vale conferir se `VERSAO_ESPERADA`/`VERSAO_SCRIPT`/`VERSAO` nos arquivos já entregues ao usuário batem com a versão atual desta conversa.

## 8. Paleta

Navy #002B50 | Green #61CF00 | Orange #E8872B | Dark #051323

## 9. Versões r47–r94

r47–r88: (ver checkpoints anteriores para o histórico completo)
r89: Dias de Estoque — corrigido o universo da faixa "Sem giro" (exige contagem física efetiva). Nova regra permanente: espelho dos 9 arquivos-fonte no projeto Claude.
r90: Sincronização de versão (pedido do cliente), sem mudança de comportamento **+** criação do Painel de Contadores nas 3 abas da Folha de Ponto (Ponto/Despesas/Financeiro), todos fixos em MTD, e dos 4 filtros de Despesas ganhando carregamento automático + cascata hierárquica.
r91: Correção de bug crítico — nova busca de período em Despesas zerava todos os lançamentos (cascata cruzava seleção do período anterior com o novo). Extensão de escopo: Ponto ganhou seletor de período próprio (não existia) e Financeiro passou a reagir ao seletor de período já existente na aba — os dois deixaram de ser MTD fixo.
r92: Correção de bug — filtro de Despesas com zero opções marcadas esvaziava as opções dos outros 3 filtros e zerava a tabela/cards; agora é tratado como "sem restrição" até o usuário marcar algo. Também corrigida falha própria: `VERSAO` do `auditoria.html` estava parada em r89 desde a r90 — sincronizada.
r93: Implementadas as 4 correções pendentes nas diretrizes de texto do relatório de Auditoria (ver seção 4): causa raiz do bug de "equipe N/A sendo mencionada" corrigida na raiz (schema condicional a `temEquipe`), tom do fechamento de prazo suavizado, exemplo "SKU único" trocado por "código de barras", e nova regra permanente proibindo exigir SKU único por pallet.
r94: Correção de bug crítico — Diretor via só as próprias despesas lançadas (mesma restrição do Supervisor). Causa: condição de corrida entre `popularFiltrosIniciais()` (sem `await`) e a busca automática MTD introduzida na r90 — a que terminasse por último vencia, e às vezes `popularFiltrosIniciais` sobrescrevia o filtro Supervisor de volta pra "só o usuário logado". Corrigido removendo a chamada, que já era redundante desde a r90.

## 10. Materiais Comerciais

Apresentação comercial "Inteligência que Transforma Operação em Resultado" (7 slides, PPTX + PDF) criada na r68 — ver checkpoints anteriores.

## 11. Pendências

- [ ] **(r94)** Publicar os 9 arquivos: GitHub Pages (frontend) + Nova Versão no Apps Script (Implantar → Gerenciar implantações → lápis → Nova versão → Implantar) — os três lados (`VERSAO_ESPERADA`, `VERSAO_SCRIPT`, `VERSAO` do auditoria.html) precisam ficar em r94 ao mesmo tempo.
- [ ] **(r94)** Usuário precisa atualizar manualmente os 9 arquivos em `codigo/<nome-do-arquivo>` no Project Knowledge (Claude não escreve lá diretamente) — inclusive este checkpoint.
- [ ] **(r94) CRÍTICO:** logar como Diretor em produção e confirmar que a aba Despesas mostra lançamentos de TODOS os supervisores no filtro Supervisor (não só o próprio Diretor) — era o bug relatado.
- [ ] **(r93)** Gerar um relatório de Auditoria real com a dimensão Equipe 100% marcada N/A e confirmar que nenhuma menção a equipe aparece em nenhuma seção (era o bug relatado originalmente).
- [ ] **(r93)** Gerar um relatório de Auditoria real com pallets abaixo de 7 e conferir o novo texto ("código de barras" em vez de "SKU único") e o novo tom do fechamento de prazo.
- [ ] **(r90)** Testar em produção o Painel de Contadores nas 3 abas (Ponto/Despesas/Financeiro), inclusive a visão Supervisor vs Diretor.
- [ ] **(r91/r92)** Testar em produção a cascata de filtros de Despesas: trocar de período, desmarcar "Selecionar todas", marcar itens específicos — nessa ordem, em sequência, pra confirmar as duas correções juntas.
- [ ] **(fora de escopo, combinado)** Avaliar se/quando aplicar filtro+cascata nos 5 modais que só têm Período hoje (Relatório de Despesas, BI, Múltiplos Projetos, Provisionamento, Adiantamento).
- [ ] Testar aba Resumo Executivo em produção (todas as combinações de arquivos)
- [ ] Testar exportação PDF Documento e PPTX do resumo
- [ ] Testar multi-unidade em produção
- [ ] Testar IA comparativa em produção
- [ ] UI e gráficos do histórico temporal
- [ ] Testar relatório Claude API individual em produção
- [ ] PDF idêntico ao HTML (print-color-adjust)
- [ ] Feedback do cliente (feedback.html)
- [ ] BI percepção do cliente
- [ ] Gráficos PRÉ vs DURANTE
- [ ] Recriar trigger analisarDespesasPendentes
- [ ] **(r69/r70)** Testar em produção a Apresentação Executiva da auditoria
- [ ] **(r74)** Confirmar que as fotos se mantêm numa auditoria reeditada
- [ ] **(r75)** Testar em produção a nova fórmula de acuracidade
- [ ] **(r76)** Validar tom da recomendação entrada/transformação/saída
- [ ] **(r78)** Confirmar que a recomendação sobre "procedimentos de contagem/lançamento" não aparece mais
- [ ] **(r79)** Testar em produção com um arquivo real de Estoque com múltiplas linhas por SKU
- [ ] **(r81) CRÍTICO:** testar em produção com o arquivo real de Contagem que originou o relato de SKU não somado
- [ ] **(r82)** Gerar uma Crítica de Inventário real com acuracidade baixa e confirmar que o Resumo Executivo não insinua problema na contagem
- [ ] **(r85)** Testar geração de Apresentação Executiva para uma auditoria que já tem relatório completo publicado
- [ ] **(r86/r87)** Gerar um relatório/apresentação real (Auditoria) com scores baixos e conferir tom
- [ ] **(r89)** Testar em produção a aba Dias de Estoque com um arquivo real que tenha SKU só de sistema (sem contagem, sem venda) e confirmar que ele não aparece mais como "Sem giro"
