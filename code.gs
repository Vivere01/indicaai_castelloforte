// ============================================================
//  GOOGLE APPS SCRIPT — Recebe indicações do formulário HTML
//  e salva na planilha Controle_Indicacoes_Barbearia
// ============================================================
//
//  COMO PUBLICAR:
//  1. Abra https://script.google.com/
//  2. Crie um novo projeto e cole este código
//  3. Clique em "Implantar" → "Nova implantação"
//  4. Tipo: "App da Web"
//     - Executar como: "Eu"
//     - Quem pode acessar: "Qualquer pessoa"
//  5. Copie a URL gerada e cole no index.html (variável SCRIPT_URL)
// ============================================================

var SPREADSHEET_ID = "1uun-641r9XMwCLrcrrDT0h-Iurh0_kP_VipnsjKQ_IY";
var SHEET_NAME     = "Indicações";  // Nome da aba — será criada se não existir

// Cabeçalhos das colunas
var HEADERS = [
  "Data/Hora",
  "Nome Assinante",
  "WhatsApp Assinante",
  "Indicado 1 — Nome",
  "Indicado 1 — WhatsApp",
  "Indicado 2 — Nome",
  "Indicado 2 — WhatsApp",
  "Indicado 3 — Nome",
  "Indicado 3 — WhatsApp",
  "Indicado 4 — Nome",
  "Indicado 4 — WhatsApp",
  "Indicado 5 — Nome",
  "Indicado 5 — WhatsApp",
  "Status"
];

// ── GET (healthcheck) ──────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "online" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST (recebe dados do formulário) ──────────────────────
function doPost(e) {
  try {
    var dados = JSON.parse(e.postData.contents);
    var sheet = obterOuCriarAba();

    var agora = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "dd/MM/yyyy HH:mm:ss"
    );

    var linha = [
      agora,
      dados.nome_assinante || "",
      dados.tel_assinante  || "",
      dados.ind1_nome || "", dados.ind1_tel || "",
      dados.ind2_nome || "", dados.ind2_tel || "",
      dados.ind3_nome || "", dados.ind3_tel || "",
      dados.ind4_nome || "", dados.ind4_tel || "",
      dados.ind5_nome || "", dados.ind5_tel || "",
      "Pendente"
    ];

    sheet.appendRow(linha);
    formatarUltimaLinha(sheet);

    return responder({ sucesso: true, mensagem: "Indicação salva com sucesso!" });

  } catch (err) {
    return responder({ sucesso: false, erro: err.toString() });
  }
}

// ── Helpers ───────────────────────────────────────────────

function obterOuCriarAba() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Cabeçalho
    sheet.appendRow(HEADERS);

    // Estilo do cabeçalho
    var cabRange = sheet.getRange(1, 1, 1, HEADERS.length);
    cabRange.setBackground("#111111");
    cabRange.setFontColor("#C9A84C");
    cabRange.setFontWeight("bold");
    cabRange.setFontSize(10);
    cabRange.setHorizontalAlignment("center");

    // Congela linha do cabeçalho
    sheet.setFrozenRows(1);

    // Largura automática
    sheet.autoResizeColumns(1, HEADERS.length);
  }

  return sheet;
}

function formatarUltimaLinha(sheet) {
  var ultima = sheet.getLastRow();
  var range  = sheet.getRange(ultima, 1, 1, HEADERS.length);

  // Zebrado claro
  if (ultima % 2 === 0) {
    range.setBackground("#F7F7F7");
  } else {
    range.setBackground("#FFFFFF");
  }
  range.setVerticalAlignment("middle");
  sheet.autoResizeColumns(1, HEADERS.length);
}

function responder(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
