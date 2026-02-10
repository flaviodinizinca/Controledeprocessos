/**
 * Log de Justificativa
 * Envia colunas N, O, P para a planilha de Logs
 */
function enviarLogJustificativa() {
  // --- CONFIGURAÇÃO MANUAL DENTRO DA FUNÇÃO ---
  // ID atualizado conforme sua solicitação
  const ID_DESTINO = "1xl9phkFAFpfC1eXclB4SrIxtfwiYcJr2h-rdhFYFbys"; 
  const NOME_ABA = "LogJustificativa";
  
  // Mapeamento de colunas (A=1, N=14, O=15, P=16)
  const C_PROC = 1;
  const C_DATA = 14;
  const C_JUST = 15;
  const C_ACAO = 16;
  // ---------------------------------------------

  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const linha = sheet.getActiveCell().getRow();

  if (linha < 2) {
    ui.alert("Por favor, selecione uma linha de dados válida.");
    return;
  }

  const processo = sheet.getRange(linha, C_PROC).getValue();
  const dataJust = sheet.getRange(linha, C_DATA).getValue();
  const just = sheet.getRange(linha, C_JUST).getValue();
  const acao = sheet.getRange(linha, C_ACAO).getValue();

  // PEGA O NOME DA GUIA (USUÁRIO)
  const nomeUsuario = sheet.getName();

  if (processo === "") {
    ui.alert("A linha selecionada não tem número de processo (Coluna A vazia).");
    return;
  }

  try {
    const ssLog = SpreadsheetApp.openById(ID_DESTINO.trim());
    const abaLog = ssLog.getSheetByName(NOME_ABA);

    if (!abaLog) {
      ui.alert(`ERRO: A aba "${NOME_ABA}" não existe na planilha de ID informado.`);
      return;
    }

    abaLog.appendRow([
      processo,
      dataJust,
      just,
      acao,
      nomeUsuario, // Agora envia o nome da guia em vez do e-mail
      new Date()
    ]);

    ui.alert("✅ Log de Justificativa enviado com sucesso!");

  } catch (e) {
    ui.alert("ERRO AO ABRIR PLANILHA DE LOGS:\n" + e.message + "\n\nVerifique se o ID está correto.");
  }
}

/**
 * Log Requisitante
 * Envia colunas R, S e calcula dias
 */
function enviarLogRequisitante() {
  // --- CONFIGURAÇÃO MANUAL DENTRO DA FUNÇÃO ---
  // Atualizei o ID aqui também para garantir
  const ID_DESTINO = "1xl9phkFAFpfC1eXclB4SrIxtfwiYcJr2h-rdhFYFbys"; 
  const NOME_ABA = "LogRequisitante";

  // Mapeamento de colunas (A=1, R=18, S=19)
  const C_PROC = 1;
  const C_ENVIO = 18;
  const C_RETORNO = 19;
  // ---------------------------------------------

  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const linha = sheet.getActiveCell().getRow();

  if (linha < 2) {
    ui.alert("Por favor, selecione uma linha de dados válida.");
    return;
  }

  const processo = sheet.getRange(linha, C_PROC).getValue();
  const dataEnvio = sheet.getRange(linha, C_ENVIO).getValue();
  const dataRetorno = sheet.getRange(linha, C_RETORNO).getValue();
  
  // Se quiser manter o padrão, aqui também pega o nome da guia
  const nomeUsuario = sheet.getName(); 

  if (processo === "") {
    ui.alert("A linha selecionada não tem número de processo (Coluna A vazia).");
    return;
  }

  // Cálculo de Dias
  let totalDias = "";
  if (dataEnvio instanceof Date && dataRetorno instanceof Date) {
    const diffTime = Math.abs(dataRetorno - dataEnvio);
    totalDias = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
  } else {
    totalDias = "Pendentes";
  }

  try {
    const ssLog = SpreadsheetApp.openById(ID_DESTINO.trim());
    const abaLog = ssLog.getSheetByName(NOME_ABA);

    if (!abaLog) {
      ui.alert(`ERRO: A aba "${NOME_ABA}" não existe na planilha de ID informado.`);
      return;
    }

    abaLog.appendRow([
      processo,
      nomeUsuario, // Nome da guia
      dataEnvio,
      dataRetorno,
      totalDias,
      new Date() 
    ]);

    ui.alert(`✅ Log Requisitante enviado! (Dias: ${totalDias})`);

  } catch (e) {
    ui.alert("ERRO AO ABRIR PLANILHA DE LOGS:\n" + e.message);
  }
}