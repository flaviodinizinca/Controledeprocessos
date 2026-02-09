/**
 * Cria uma nova aba para o comprador com a estrutura atualizada (Incluindo Prioridade).
 */
function criarGuiaComprador() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Novo Comprador', 'Digite o nome do comprador (será o nome da guia):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }

  const nomeGuia = response.getResponseText().trim();
  
  if (nomeGuia === "") {
    ui.alert('O nome não pode estar vazio.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (ss.getSheetByName(nomeGuia)) {
    ui.alert('Já existe uma guia com este nome: ' + nomeGuia);
    return;
  }

  const novaGuia = ss.insertSheet(nomeGuia);
  configurarEstruturaGuia(novaGuia, ss);
}

/**
 * Configura as colunas, validações e cores da nova guia.
 */
function configurarEstruturaGuia(novaGuia, ss) {
  const ui = SpreadsheetApp.getUi();
  
  // Definição dos Cabeçalhos (20 Colunas agora)
  const cabecalhos = [
    // Bloco 1: Geral (Colunas A-D)
    "Processo", "Objeto", "Modalidade", "Qtd Itens", 
    
    // Bloco 2: Datas SECOM/ASTEC (Colunas E-H)
    "Data de recebimento", "Data Envio ASTEC", "Data Rec. Secr.", "Data Rec. Serv.", 
    
    // Bloco 3: Comprador e Pesquisa (Colunas I-P)
    "Rec. Comprador", "Inicio Pesquisa",
    "Data", "Justificativa 01", "Ação 01",
    "Data", "Justificativa 02", "Ação 02",
    
    // Bloco 4: Finalização (Colunas Q-R)
    "Envio Chefia", "Envio COAGE",
    
    // Bloco 5: Automático
    "Prazo Limite (60 dias)",
    
    // Bloco 6: Prioridade (Coluna T - Nova)
    "STATUS PRIORIDADE"
  ];

  // Aplica cabeçalhos na linha 1
  const rangeCabecalho = novaGuia.getRange(1, 1, 1, cabecalhos.length);
  rangeCabecalho.setValues([cabecalhos]);
  rangeCabecalho.setFontWeight("bold");
  rangeCabecalho.setHorizontalAlignment("center");
  rangeCabecalho.setVerticalAlignment("middle");
  rangeCabecalho.setWrap(true);

  // --- Aplicação de Cores ---
  novaGuia.getRange(1, 1, 1, 4).setBackground("#EEEEEE"); // Geral
  novaGuia.getRange(1, 5, 1, 4).setBackground("#CFE2F3"); // SECOM
  novaGuia.getRange(1, 9, 1, 8).setBackground("#FFF2CC"); // Comprador
  novaGuia.getRange(1, 17, 1, 2).setBackground("#D9EAD3"); // Finalização
  novaGuia.getRange(1, 19, 1, 1).setBackground("#F4CCCC"); // Prazo
  novaGuia.getRange(1, 20, 1, 1).setBackground("#EA9999").setFontColor("white"); // Prioridade (Vermelho Escuro)

  // Congela painéis
  novaGuia.setFrozenRows(1);
  novaGuia.setFrozenColumns(4);

  // --- Validação (Modalidade) ---
  const guiaConfig = ss.getSheetByName("Modal_Config");
  if (guiaConfig) {
    const rangeValidacao = guiaConfig.getRange("A2:A");
    const regraValidacao = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeValidacao, true).setAllowInvalid(false).build();
    novaGuia.getRange(2, 3, 999, 1).setDataValidation(regraValidacao);
  }

  // Ajustes Finais
  novaGuia.autoResizeColumns(1, cabecalhos.length);
  novaGuia.setColumnWidth(2, 300); // Objeto
  novaGuia.setColumnWidth(20, 150); // Coluna Prioridade
}