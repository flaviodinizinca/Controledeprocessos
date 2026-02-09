/**
 * Cria uma nova aba para o comprador com a estrutura atualizada.
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
 * Inclui a coluna "Marcador" na Coluna B e "Especificação" na Coluna C.
 */
function configurarEstruturaGuia(novaGuia, ss) {
  const ui = SpreadsheetApp.getUi();
  
  // Definição dos Cabeçalhos Atualizada
  const cabecalhos = [
    "Processo",           // Coluna A
    "Marcador",            // Coluna B (Nova)
    "Especificação",       // Coluna C (Anteriormente na Coluna B)
    "Objeto",              // Coluna D
    "Modalidade",          // Coluna E
    "Qtd Itens",           // Coluna F
    "Valor Estimado",      // Coluna G
    "Status Atual",        // Coluna H
    "Fase SEI",            // Coluna I
    "Prioridade",          // Coluna J
    "Data Rec. Comprador", // Coluna K
    "Início Pesquisa",     // Coluna L
    "Fim Pesquisa",        // Coluna M
    "Envio Área Requisitante",   // Coluna N
    "Retorno Área Requisitante", // Coluna O
    "Envio p/ Licitação",        // Coluna P
    "Dias na Pesquisa",          // Coluna Q
    "Dias com Requisitante",     // Coluna R
    "Total dias Comprador",      // Coluna S
    "Observações"                // Coluna T
  ];

  // Aplica cabeçalhos na linha 1
  const rangeCabecalho = novaGuia.getRange(1, 1, 1, cabecalhos.length);
  rangeCabecalho.setValues([cabecalhos]);
  rangeCabecalho.setFontWeight("bold");
  rangeCabecalho.setHorizontalAlignment("center");
  rangeCabecalho.setVerticalAlignment("middle");
  rangeCabecalho.setWrap(true);

  // --- Aplicação de Cores de Fundo ---
  novaGuia.getRange(1, 1, 1, 3).setBackground("#EEEEEE"); // Processo, Marcador e Especificação
  novaGuia.getRange(1, 4, 1, 4).setBackground("#CFE2F3"); // Objeto, Modalidade, Qtd, Valor
  novaGuia.getRange(1, 8, 1, 3).setBackground("#FFF2CC"); // Status, Fase, Prioridade
  novaGuia.getRange(1, 11, 1, 6).setBackground("#D9EAD3"); // Datas de acompanhamento
  novaGuia.getRange(1, 17, 1, 3).setBackground("#F4CCCC"); // Cálculos de dias
  novaGuia.getRange(1, 20, 1, 1).setBackground("#EA9999").setFontColor("white"); // Observações

  // Congela a primeira linha para facilitar a navegação
  novaGuia.setFrozenRows(1);
  
  // Congela as primeiras colunas (até Objeto) para referência visual
  novaGuia.setFrozenColumns(4);

  // --- Validação de Dados (Exemplo para Modalidade) ---
  const guiaConfig = ss.getSheetByName("Modal_Config");
  if (guiaConfig) {
    const rangeValidacao = guiaConfig.getRange("A2:A");
    const regraValidacao = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeValidacao, true)
      .setAllowInvalid(false)
      .build();
    // Aplica a validação na Coluna E (Modalidade)
    novaGuia.getRange(2, 5, 999, 1).setDataValidation(regraValidacao);
  }

  // Ajustes de largura de colunas
  novaGuia.autoResizeColumns(1, cabecalhos.length);
  novaGuia.setColumnWidth(4, 300); // Largura maior para o Objeto
  novaGuia.setColumnWidth(20, 250); // Largura maior para Observações
}