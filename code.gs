// ============================================================
//  GOOGLE APPS SCRIPT — Indica Aí | Castello Forte Barbearia
//  Versão 2.0
//
//  COMO PUBLICAR:
//  1. Acesse https://script.google.com/
//  2. Novo projeto → cole este código inteiro
//  3. Implantar → Nova implantação
//     - Tipo: App da Web
//     - Executar como: Eu
//     - Quem pode acessar: Qualquer pessoa
//  4. Copie a URL gerada e cole no index.html (var SCRIPT_URL)
// ============================================================

// ── Configurações ──────────────────────────────────────────
var SPREADSHEET_ID      = "1uun-641r9XMwCLrcrrDT0h-Iurh0_kP_VipnsjKQ_IY";

// Aba 1 — Resumo de envios (uma linha por assinante que enviou)
var ABA_INDICACOES      = "Indicações";

// Aba 2 — Controle de abordagem (uma linha por indicado)
// Identificada pelo GID interno da planilha (gid=273016962)
var GID_CONTROLE        = 273016962;

// Na aba de controle, o cabeçalho está na linha 3 e dados a partir da linha 4
var LINHA_DADOS_CONTROLE = 4;

// ── Cabeçalhos Aba 1 ───────────────────────────────────────
var CABECALHO_INDICACOES = [
  "ID",
  "Data/Hora",
  "Nome Assinante",
  "WhatsApp Assinante",
  "Unidade",
  "Ind.1 — Nome", "Ind.1 — Tel",
  "Ind.2 — Nome", "Ind.2 — Tel",
  "Ind.3 — Nome", "Ind.3 — Tel",
  "Ind.4 — Nome", "Ind.4 — Tel",
  "Ind.5 — Nome", "Ind.5 — Tel",
  "Total Indicados"
];

// ── Cabeçalhos Aba 2 (apenas referência — a aba já existe) ─
// A: ID | B: ID Indicação | C: Nome Indicado | D: Unidade
// E: Recepcionista | F: Telefone | G: Indicado Por
// H: 1º Contato (Data) | I: Resultado 1º | J: 2º Contato | K: Resultado 2º | L: Status Final

// ──────────────────────────────────────────────────────────
//  GET — healthcheck
// ──────────────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "online", versao: "2.0" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ──────────────────────────────────────────────────────────
//  POST — recebe dados do formulário
// ──────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Grava resumo na Aba 1
    var abaIndicacoes = obterOuCriarAbaIndicacoes(ss);
    var idIndicacao   = gravarResumoAssinante(abaIndicacoes, dados);

    // 2. Grava cada indicado na Aba 2 (Controle de Abordagem)
    var abaControle = obterAbaPorGid(ss, GID_CONTROLE);
    gravarIndicadosNoControle(abaControle, dados, idIndicacao);

    return responder({ sucesso: true, id: idIndicacao });

  } catch (err) {
    Logger.log("ERRO doPost: " + err.toString());
    return responder({ sucesso: false, erro: err.toString() });
  }
}

// ──────────────────────────────────────────────────────────
//  ABA 1 — Resumo por assinante
// ──────────────────────────────────────────────────────────
function obterOuCriarAbaIndicacoes(ss) {
  var aba = ss.getSheetByName(ABA_INDICACOES);
  if (!aba) {
    aba = ss.insertSheet(ABA_INDICACOES);
    var cab = aba.getRange(1, 1, 1, CABECALHO_INDICACOES.length);
    cab.setValues([CABECALHO_INDICACOES]);
    cab.setBackground("#111111");
    cab.setFontColor("#C9A84C");
    cab.setFontWeight("bold");
    cab.setFontSize(10);
    cab.setHorizontalAlignment("center");
    aba.setFrozenRows(1);
    aba.autoResizeColumns(1, CABECALHO_INDICACOES.length);
  }
  return aba;
}

function gravarResumoAssinante(aba, dados) {
  var agora = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss"
  );

  // Calcula próximo ID (número de linhas de dados + 1)
  var ultimaLinha = aba.getLastRow();
  var novoId      = (ultimaLinha < 1 ? 1 : ultimaLinha); // linha 1 = cabeçalho → ID começa em 1

  // Conta quantos indicados foram preenchidos
  var totalIndicados = 0;
  for (var i = 1; i <= 5; i++) {
    if (dados["ind" + i + "_nome"]) totalIndicados++;
  }

  var linha = [
    novoId,
    agora,
    dados.nome_assinante || "",
    dados.tel_assinante  || "",
    dados.unidade        || "",
    dados.ind1_nome || "", dados.ind1_tel || "",
    dados.ind2_nome || "", dados.ind2_tel || "",
    dados.ind3_nome || "", dados.ind3_tel || "",
    dados.ind4_nome || "", dados.ind4_tel || "",
    dados.ind5_nome || "", dados.ind5_tel || "",
    totalIndicados
  ];

  aba.appendRow(linha);

  // Zebrado
  var ult = aba.getLastRow();
  aba.getRange(ult, 1, 1, CABECALHO_INDICACOES.length)
     .setBackground(ult % 2 === 0 ? "#F7F7F7" : "#FFFFFF");

  return novoId;
}

// ──────────────────────────────────────────────────────────
//  ABA 2 — Controle de abordagem por indicado
// ──────────────────────────────────────────────────────────
function obterAbaPorGid(ss, gid) {
  var abas = ss.getSheets();
  for (var i = 0; i < abas.length; i++) {
    if (abas[i].getSheetId() === gid) return abas[i];
  }
  // Se não encontrar pelo gid, tenta criar uma nova aba de controle
  Logger.log("Aba com gid=" + gid + " não encontrada. Criando nova aba.");
  return criarAbaControle(ss);
}

function criarAbaControle(ss) {
  var aba = ss.insertSheet("Controle_Indicados");

  // Linha 1 — Título
  var titulo = aba.getRange(1, 1, 1, 12);
  titulo.merge();
  titulo.setValue("📞 CONTROLE DE ABORDAGEM DOS CLIENTES INDICADOS");
  titulo.setBackground("#111111");
  titulo.setFontColor("#C9A84C");
  titulo.setFontWeight("bold");
  titulo.setFontSize(14);
  titulo.setHorizontalAlignment("center");

  // Linha 2 — Descrição
  var desc = aba.getRange(2, 1, 1, 12);
  desc.merge();
  desc.setValue("Registre cada contato feito com os indicados. Atualize o status conforme o avanço.");
  desc.setBackground("#222222");
  desc.setFontColor("#AAAAAA");
  desc.setFontSize(10);
  desc.setHorizontalAlignment("center");

  // Linha 3 — Cabeçalhos
  var cabecalhos = [
    "ID", "ID\nINDICAÇÃO", "NOME INDICADO", "Unidade",
    "recepcionista", "TELEFONE", "INDICADO POR",
    "1º CONTATO\n(DATA)", "RESULTADO\n1º CONTATO",
    "2º CONTATO\n(DATA)", "RESULTADO\n2º CONTATO",
    "STATUS FINAL"
  ];
  var cab = aba.getRange(3, 1, 1, cabecalhos.length);
  cab.setValues([cabecalhos]);
  cab.setBackground("#1A1A1A");
  cab.setFontColor("#C9A84C");
  cab.setFontWeight("bold");
  cab.setFontSize(9);
  cab.setHorizontalAlignment("center");
  cab.setWrap(true);
  aba.setRowHeight(3, 40);
  aba.setFrozenRows(3);
  aba.autoResizeColumns(1, cabecalhos.length);
  return aba;
}

function gravarIndicadosNoControle(aba, dados, idIndicacao) {
  var agora = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );

  // Busca próximo ID disponível na aba de controle
  // (conta linhas com dados a partir de LINHA_DADOS_CONTROLE)
  var proximoId = obterProximoIdControle(aba);

  for (var i = 1; i <= 5; i++) {
    var nome = dados["ind" + i + "_nome"] ? dados["ind" + i + "_nome"].trim() : "";
    var tel  = dados["ind" + i + "_tel"]  ? dados["ind" + i + "_tel"].trim()  : "";

    if (!nome) continue; // Pula indicados não preenchidos

    // Colunas: A=ID | B=ID_Indicação | C=Nome | D=Unidade | E=Recepcionista (vazio)
    //          F=Telefone | G=Indicado Por | H..K = operador (vazio) | L=Status Final
    var linha = [
      proximoId,                           // A — ID (auto)
      idIndicacao,                          // B — ID Indicação (ref. Aba 1)
      nome,                                 // C — Nome Indicado
      dados.unidade || "",                  // D — Unidade
      "",                                   // E — Recepcionista (operador preenche)
      tel,                                  // F — Telefone
      dados.nome_assinante || "",           // G — Indicado Por
      "",                                   // H — 1º Contato Data (operador)
      "",                                   // I — Resultado 1º Contato (operador)
      "",                                   // J — 2º Contato Data (operador)
      "",                                   // K — Resultado 2º Contato (operador)
      "Pendente Contato"                    // L — Status Final (padrão inicial)
    ];

    // Encontra a próxima linha vazia a partir de LINHA_DADOS_CONTROLE
    var linhaAlvo = encontrarProximaLinhaVazia(aba, LINHA_DADOS_CONTROLE, 3); // col C = Nome
    aba.getRange(linhaAlvo, 1, 1, linha.length).setValues([linha]);

    // Estilo da célula de Status Final
    estilizarStatusFinal(aba, linhaAlvo);

    proximoId++;
  }

  aba.autoResizeColumns(1, 12);
}

// Retorna o próximo ID a usar na aba de controle
function obterProximoIdControle(aba) {
  var ultima = aba.getLastRow();
  if (ultima < LINHA_DADOS_CONTROLE) return 1;

  // Lê coluna A (IDs) a partir da linha de dados
  var vals = aba.getRange(LINHA_DADOS_CONTROLE, 1, ultima - LINHA_DADOS_CONTROLE + 1, 1).getValues();
  var maxId = 0;
  vals.forEach(function(row) {
    var v = parseInt(row[0], 10);
    if (!isNaN(v) && v > maxId) maxId = v;
  });
  return maxId + 1;
}

// Encontra a próxima linha vazia na coluna indicada, a partir de startRow
function encontrarProximaLinhaVazia(aba, startRow, colRef) {
  var ultima = aba.getLastRow();
  if (ultima < startRow) return startRow;

  var vals = aba.getRange(startRow, colRef, ultima - startRow + 1, 1).getValues();
  for (var r = 0; r < vals.length; r++) {
    if (!vals[r][0] || vals[r][0].toString().trim() === "") {
      return startRow + r;
    }
  }
  return ultima + 1; // Adiciona após a última linha se todas estiverem ocupadas
}

// Aplica cor de fundo e estilo à célula de Status Final (coluna L = 12)
function estilizarStatusFinal(aba, linha) {
  var celStatus = aba.getRange(linha, 12);
  celStatus.setBackground("#1B6CA8");
  celStatus.setFontColor("#FFFFFF");
  celStatus.setFontWeight("bold");
  celStatus.setHorizontalAlignment("center");
  celStatus.setFontSize(9);
}

// ──────────────────────────────────────────────────────────
//  Helper — retorna resposta JSON
// ──────────────────────────────────────────────────────────
function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
