// ============================================================
//  GOOGLE APPS SCRIPT — Indica Aí | Castello Forte Barbearia
//  Versão 3.0 (Otimizada para IDs Pré-preenchidos)
// ============================================================

var SPREADSHEET_ID = "1lbFXxsPdTBYZ0MFeJKfs__P-qnwr7ve65SSSIXdximo";

// IDs internos (GIDs) das abas
var GID_INDICACOES = 1471518501; // Aba 1 — Indicações
var GID_CONTROLE   = 273016962;  // Aba 2 — Abordagem

// ── GET — Teste de conexão ──────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "online", versao: "3.0" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST — Recebe dados do formulário ───────────────────────
function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Salva na Aba 1 — "Indicações"
    var abaIndicacoes = obterAbaPorGid(ss, GID_INDICACOES);
    var idIndicacao   = gravarAba1(abaIndicacoes, dados);

    // 2. Salva na Aba 2 — "Abordagem"
    var abaControle = obterAbaPorGid(ss, GID_CONTROLE);
    gravarAba2(abaControle, dados, idIndicacao);

    return responder({ sucesso: true, id: idIndicacao });

  } catch (err) {
    return responder({ sucesso: false, erro: err.toString() });
  }
}

// ── Gravar na Aba 1 (Indicações) ──────────────────────────
function gravarAba1(aba, dados) {
  var agora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  
  // A partir da linha 5 (onde começam os dados de verdade)
  var startRow = 5; 
  var ultimaLinha = aba.getLastRow();
  var linhaDestino = startRow;

  // Busca a primeira linha onde a coluna C (Nome do Assinante) está vazia
  if (ultimaLinha >= startRow) {
    var valores = aba.getRange(startRow, 3, ultimaLinha - startRow + 1, 1).getValues();
    for (var i = 0; i < valores.length; i++) {
      if (!valores[i][0] || valores[i][0].toString().trim() === "") {
        linhaDestino = startRow + i;
        break;
      }
    }
    if (linhaDestino === startRow && valores[valores.length - 1][0] !== "") {
      linhaDestino = ultimaLinha + 1;
    }
  }

  // Pega o ID que já está preenchido na coluna A, ou gera um novo caso necessário
  var idCel = aba.getRange(linhaDestino, 1).getValue();
  var idFinal = (idCel !== "") ? idCel : (linhaDestino - 4);

  // Prepara os dados (16 colunas: da B até a Q)
  var dadosLinha = [
    agora,                       // B: Data
    dados.nome_assinante || "",  // C: Nome
    "",                          // D: CPF (em branco)
    dados.tel_assinante  || "",  // E: Telefone
    dados.unidade        || "",  // F: Unidade
    "",                          // G: Recepcionista (em branco)
    dados.ind1_nome || "", dados.ind1_tel || "", // H, I: Indicado 1
    dados.ind2_nome || "", dados.ind2_tel || "", // J, K: Indicado 2
    dados.ind3_nome || "", dados.ind3_tel || "", // L, M: Indicado 3
    dados.ind4_nome || "", dados.ind4_tel || "", // N, O: Indicado 4
    dados.ind5_nome || "", dados.ind5_tel || "", // P, Q: Indicado 5
  ];

  aba.getRange(linhaDestino, 2, 1, dadosLinha.length).setValues([dadosLinha]);
  return idFinal;
}

// ── Gravar na Aba 2 (Abordagem) ───────────────────────────
function gravarAba2(aba, dados, idIndicacao) {
  var startRow = 4; // Dados começam na linha 4
  
  for (var i = 1; i <= 5; i++) {
    var nome = dados["ind" + i + "_nome"] ? dados["ind" + i + "_nome"].trim() : "";
    var tel  = dados["ind" + i + "_tel"]  ? dados["ind" + i + "_tel"].trim()  : "";

    if (!nome) continue; // Pula se o indicado estiver vazio

    var ultimaLinha = aba.getLastRow();
    var linhaDestino = startRow;

    // Busca a primeira linha onde a coluna C (Nome Indicado) está vazia
    if (ultimaLinha >= startRow) {
      var valores = aba.getRange(startRow, 3, ultimaLinha - startRow + 1, 1).getValues();
      var achou = false;
      for (var j = 0; j < valores.length; j++) {
        if (!valores[j][0] || valores[j][0].toString().trim() === "") {
          linhaDestino = startRow + j;
          achou = true;
          break;
        }
      }
      if (!achou) {
        linhaDestino = ultimaLinha + 1;
      }
    }

    // Prepara dados para Aba 2 (11 colunas: da B até a L)
    var dadosLinha = [
      idIndicacao,                // B: ID Indicação
      nome,                       // C: Nome Indicado
      dados.unidade || "",        // D: Unidade
      "",                         // E: Recepcionista (em branco)
      tel,                        // F: Telefone
      dados.nome_assinante || "", // G: Indicado Por
      "",                         // H: 1º Contato (Data)
      "",                         // I: Resultado 1º
      "",                         // J: 2º Contato (Data)
      "",                         // K: Resultado 2º
      "Pendente Contato"          // L: Status Final
    ];

    aba.getRange(linhaDestino, 2, 1, dadosLinha.length).setValues([dadosLinha]);
    
    // Formata o status (Col L) com fundo azul
    var celStatus = aba.getRange(linhaDestino, 12);
    celStatus.setBackground("#1B6CA8");
    celStatus.setFontColor("#FFFFFF");
    celStatus.setFontWeight("bold");
    celStatus.setHorizontalAlignment("center");
  }
}

// ── Obter Aba pelo ID ──────────────────────────────────────
function obterAbaPorGid(ss, gid) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) {
      return sheets[i];
    }
  }
  throw new Error("Aba com GID " + gid + " não encontrada na planilha.");
}

// ── Retorna Resposta JSON ─────────────────────────────────
function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
