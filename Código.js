
/*
 * CENTRAL DE FEEDBACK DO MOSTEIRO
 * Painel administrativo privado
 */

const SPREADSHEET_ID = "1MULLLzuk9pyLwjyBoRZ6iFje46SzjBpqbaJATw_Arzk";
const ABA_FEEDBACK = "Feedback_Central";

/*
 * A proteção principal deste painel será feita
 * pela configuração de acesso da implantação:
 * "Somente eu".
 */

function doGet(e) {
  const modulo = e && e.parameter ? e.parameter.module : "";

  // O módulo é servido separadamente para ser incorporado
  // na área de trabalho da interface principal.
  if (modulo === "feedback") {
    return HtmlService
      .createHtmlOutputFromFile("Feedback")
      .setTitle("Feedback · Centralce")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService
    .createHtmlOutputFromFile("Painel")
    .setTitle("Centralce · Mosteiro da Santa Cruz")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}


/**
 * Consulta os comentários registrados.
 */
function obterFeedbacks() {
  const planilha = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = planilha.getSheetByName(ABA_FEEDBACK);

  if (!aba || aba.getLastRow() < 2) {
    return [];
  }

  const dados = aba.getRange(
    2, 1, aba.getLastRow() - 1, 7
  ).getValues();

  return dados.reverse().map(linha => ({
    data: linha[0] instanceof Date
      ? Utilities.formatDate(
          linha[0],
          Session.getScriptTimeZone(),
          "dd/MM/yyyy HH:mm"
        )
      : String(linha[0]),

    id: String(linha[1]),
    origem: String(linha[2]),
    pagina: String(linha[3]),
    comentario: String(linha[4]),
    situacao: String(linha[5]),
    observacoes: String(linha[6])
  }));
}


/**
 * Atualiza a situação e as observações.
 */
function atualizarFeedback(id, situacao, observacoes) {
  const situacoesPermitidas = [
    "Recebido",
    "Em análise",
    "Atendido",
    "Arquivado"
  ];

  if (!situacoesPermitidas.includes(situacao)) {
    throw new Error("Situação inválida.");
  }

  const planilha = SpreadsheetApp.openById(SPREADSHEET_ID);
  const aba = planilha.getSheetByName(ABA_FEEDBACK);

  if (!aba || aba.getLastRow() < 2) {
    throw new Error("Nenhum feedback encontrado.");
  }

  const dados = aba.getRange(
    2, 1, aba.getLastRow() - 1, 7
  ).getValues();

  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][1]) === String(id)) {
      aba.getRange(i + 2, 6).setValue(situacao);
      aba.getRange(i + 2, 7).setValue(
        String(observacoes || "").slice(0, 2000)
      );

      return true;
    }
  }

  throw new Error("Feedback não encontrado.");
}
