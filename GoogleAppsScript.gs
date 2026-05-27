const SHEET_NAME = "Respostas";

const HEADERS = [
  "data_recebimento",
  "nome",
  "email",
  "telefone",
  "cargo",
  "tipo_organizacao",
  "qtd_clientes",
  "equipe_fiscal",
  "regimes",
  "volume_documentos",
  "documentos",
  "problemas_documentos",
  "dor_organizacao",
  "dor_cadastro",
  "dor_fechamento",
  "retrababalho",
  "horas_retrabalho",
  "ferramentas",
  "controle_pendencias",
  "preparacao_reforma",
  "preocupacoes_reforma",
  "modulos_interesse",
  "valor_percebido",
  "modelo_preferido",
  "preco_diagnostico",
  "preco_mensalidade",
  "interesse_piloto",
  "maior_dor",
  "valor_imediato",
  "comentarios",
  "origem",
  "user_agent"
];

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: "online",
      message: "Web App Autfisca ativo"
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error("Apps Script nao esta vinculado a uma planilha. Crie o Apps Script dentro do Google Sheets em Extensoes > Apps Script.");
    }

    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);

      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight("bold")
        .setBackground("#1f4e79")
        .setFontColor("#ffffff");

      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, HEADERS.length);
    }

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Nenhum dado recebido no POST.");
    }

    const data = JSON.parse(e.postData.contents);
    data.data_recebimento = new Date();

    const row = HEADERS.map(function(h) {
      return data[h] !== undefined && data[h] !== null ? data[h] : "";
    });

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({
        status: "success",
        message: "Resposta salva com sucesso"
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: "error",
        message: String(error)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
