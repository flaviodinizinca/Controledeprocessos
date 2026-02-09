/**
 * 02_Guias.gs
 * Gerencia a criação e estruturação das guias (Padrão e Saneamento).
 */

function criarGuiaComprador(nomeGuia, tipo = "PADRAO") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (ss.getSheetByName(nomeGuia)) {
    return ss.getSheetByName(nomeGuia); 
  }

  const novaGuia = ss.insertSheet(nomeGuia);

  if (tipo === "SANEAMENTO") {
    configurarEstruturaSaneamento(novaGuia);
  } else {
    configurarEstruturaGuia(novaGuia, ss);
  }
  
  return novaGuia;
}

/** ESTRUTURA PADRÃO */
function configurarEstruturaGuia(novaGuia, ss) {
  const cabecalhos = [
    "Processo", "Marcador", "Especificação", "Objeto", "Modalidade",
    "Qtd Itens", "Valor Estimado", "Status Atual", "Fase SEI", "Prioridade",
    "Data Rec. Comprador", "Início Pesquisa", "Fim Pesquisa", "Envio Área Requisitante",
    "Retorno Área Requisitante", "Envio p/ Licitação", "Dias na Pesquisa",
    "Dias com Requisitante", "Total dias Comprador", "Observações"
  ];

  const rangeCabecalho = novaGuia.getRange(1, 1, 1, cabecalhos.length);
  rangeCabecalho.setValues([cabecalhos]);
  rangeCabecalho.setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  novaGuia.setRowHeight(1, 45);

  novaGuia.getRange(1, 1, 1, 3).setBackground("#EEEEEE"); 
  novaGuia.getRange(1, 4, 1, 4).setBackground("#CFE2F3"); 
  novaGuia.getRange(1, 8, 1, 3).setBackground("#FFF2CC"); 
  novaGuia.getRange(1, 11, 1, 6).setBackground("#D9EAD3"); 
  novaGuia.getRange(1, 17, 1, 3).setBackground("#F4CCCC"); 
  novaGuia.getRange(1, 20, 1, 1).setBackground("#EA9999").setFontColor("white");

  novaGuia.setFrozenRows(1);
  novaGuia.setFrozenColumns(4);

  const guiaConfig = ss.getSheetByName("Modal_Config");
  if (guiaConfig) {
    const regra = SpreadsheetApp.newDataValidation().requireValueInRange(guiaConfig.getRange("A2:A"), true).build();
    novaGuia.getRange(2, 5, 999, 1).setDataValidation(regra);
  }

  novaGuia.autoResizeColumns(1, cabecalhos.length);
  novaGuia.setColumnWidth(1, 140); 
  novaGuia.setColumnWidth(2, 100);
  novaGuia.setColumnWidth(3, 200);
  novaGuia.setColumnWidth(4, 300);
}

/** ESTRUTURA SANEAMENTO */
function configurarEstruturaSaneamento(novaGuia) {
  const cabecalhos = [
    "PROCESSO", "Data de Chegada", "PROTOCOLO", "PARECER/ NOTA/ COTA", "OBJETO",
    "CÉLULA", "MODALIDADE", "DATA DO STATUS", "SANEAMENTO ENCERRADO?", "LOCALIZAÇÃO", "STATUS"
  ];

  const rangeCabecalho = novaGuia.getRange(1, 1, 1, cabecalhos.length);
  rangeCabecalho.setValues([cabecalhos]);
  rangeCabecalho.setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  novaGuia.setRowHeight(1, 45);

  novaGuia.getRange(1, 1, 1, 11).setBackground("#FCE5CD");
  novaGuia.getRange(1, 1, 1, 1).setBackground("#E69138").setFontColor("white");
  novaGuia.getRange(1, 9, 1, 1).setBackground("#EA9999");

  novaGuia.setFrozenRows(1);
  novaGuia.setFrozenColumns(2);

  const regraSimNao = SpreadsheetApp.newDataValidation().requireValueInList(["SIM", "NÃO"], true).setAllowInvalid(false).build();
  novaGuia.getRange(2, 9, 999, 1).setDataValidation(regraSimNao);

  const regraData = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  novaGuia.getRange(2, 2, 999, 1).setDataValidation(regraData);
  novaGuia.getRange(2, 8, 999, 1).setDataValidation(regraData);

  novaGuia.autoResizeColumns(1, cabecalhos.length);
  novaGuia.setColumnWidth(1, 150);
  novaGuia.setColumnWidth(5, 250);
  novaGuia.setColumnWidth(4, 200);
}