// --- FUNÇÕES DE CRIAÇÃO DE GUIAS ---

/**
 * 1. Cria a Guia de Abertura de Processos (Genérica)
 */
function criarGuiaAberturaProcessos() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt('Nova Guia Abertura', 'Qual o nome do usuário/guia?', ui.ButtonSet.OK_CANCEL);

  if (resposta.getSelectedButton() == ui.Button.OK) {
    const nomeGuia = resposta.getResponseText();
    
    const cabecalhos = [
      "Processo", "Marcador", "Especificação", "Objeto", "Modalidade", 
      "Qtd Itens", "Valor Estimado notes", "Data de Chegada do Notes", 
      "Devolução Notes", "Retorno Notes", "Dara Envio Pendencias Iniciais", 
      "Data Retorno Processo", "Tem Marcas (Sim ou Não)", 
      "Avaliação da Amostra(Física/Catalogo/Ambos)", "TR/Anexos", "Atribuição ao servidor"
    ];

    configurarNovaGuia(nomeGuia, cabecalhos);
  }
}

/**
 * 2. Cria a Guia de IRP
 */
function criarGuiaIRP() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt('Nova Guia IRP', 'Qual o nome do usuário/guia?', ui.ButtonSet.OK_CANCEL);

  if (resposta.getSelectedButton() == ui.Button.OK) {
    const nomeGuia = resposta.getResponseText();
    
    const cabecalhos = [
      "Processo", "Marcador", "Especificação", "Objeto", "Modalidade", 
      "Qtd Itens", "Valor Estimado", "Data Rec.", "Publicação para manifestação IRP", 
      "Finalização da mnifestação IRP", "Solicitação de confirmação", 
      "Confirmação", "Qtd de Particitantes", "Atribuição ao servidor"
    ];

    configurarNovaGuia(nomeGuia, cabecalhos);
  }
}

/**
 * 3. Cria a Guia Comprador (NOVA)
 */
function criarGuiaComprador() {
  const ui = SpreadsheetApp.getUi();
  const resposta = ui.prompt('Nova Guia Comprador', 'Qual o nome do usuário/guia?', ui.ButtonSet.OK_CANCEL);

  if (resposta.getSelectedButton() == ui.Button.OK) {
    const nomeGuia = resposta.getResponseText();
    
    // Lista exata das 24 colunas solicitadas
    const cabecalhos = [
      "Processo",                   // A - 1
      "Marcador",                   // B - 2
      "Especificação",              // C - 3
      "Objeto",                     // D - 4
      "Modalidade",                 // E - 5 (Lista Suspensa da Modal_Config)
      "Qtd Itens",                  // F - 6
      "Valor Estimado",             // G - 7
      "Status Atual",               // H - 8
      "Fase SEI",                   // I - 9
      "Prioridade",                 // J - 10
      "Data Rec. Comprador",        // K - 11
      "Início Pesquisa",            // L - 12
      "Param. IN 65/2021",          // M - 13
      "Data da Justificativa",      // N - 14
      "Justificativa",              // O - 15
      "Ação",                       // P - 16
      "Fim Pesquisa",               // Q - 17
      "Envio Área Requisitante",    // R - 18
      "Retorno Área Requisitante",  // S - 19
      "Envio p/ Licitação",         // T - 20
      "Dias na Pesquisa",           // U - 21
      "Dias com Requisitante",      // V - 22
      "Total de dias com Comprador",// W - 23
      "Observações"                 // X - 24
    ];

    configurarNovaGuia(nomeGuia, cabecalhos);
  }
}

// --- FUNÇÕES AUXILIARES ---

function configurarNovaGuia(nomeGuia, cabecalhos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (ss.getSheetByName(nomeGuia)) {
    SpreadsheetApp.getUi().alert('Já existe uma guia com este nome!');
    return;
  }

  const novaGuia = ss.insertSheet(nomeGuia);
  
  // Define cabeçalhos e formata
  novaGuia.getRange(1, 1, 1, cabecalhos.length)
    .setValues([cabecalhos])
    .setFontWeight("bold")
    .setBackground("#EFEFEF"); // Um cinza claro para destacar o cabeçalho
  
  novaGuia.setFrozenRows(1);

  // --- CONFIGURAÇÃO COLUNA B (MARCADOR) ---
  // Usa a lista do Config.gs
  const regraMarcador = SpreadsheetApp.newDataValidation()
    .requireValueInList(LISTA_MARCADORES, true)
    .setAllowInvalid(true)
    .build();
  
  novaGuia.getRange(2, 2, 999, 1).setDataValidation(regraMarcador);

  // --- CONFIGURAÇÃO COLUNA E (MODALIDADE) ---
  // Tenta pegar da aba "Modal_Config", se não existir, usa a lista padrão do Config
  let regraModalidade;
  const abaModalConfig = ss.getSheetByName("Modal_Config");

  if (abaModalConfig) {
    // Pega o intervalo A2:A da aba Modal_Config para permitir expansão futura
    // A referência fica dinâmica (se você adicionar lá, aparece aqui)
    const intervaloModalidades = abaModalConfig.getRange("A2:A");
    regraModalidade = SpreadsheetApp.newDataValidation()
      .requireValueInRange(intervaloModalidades, true) // Lê direto da aba
      .setAllowInvalid(true)
      .build();
  } else {
    // Fallback caso a aba não exista (usa a lista do código)
    regraModalidade = SpreadsheetApp.newDataValidation()
      .requireValueInList(LISTA_MODALIDADES || ["Pregão Eletrônico", "Dispensa"], true)
      .setAllowInvalid(true)
      .build();
  }

  novaGuia.getRange(2, 5, 999, 1).setDataValidation(regraModalidade);

  // Aplica cores nos marcadores
  aplicarCoresCondicionais(novaGuia);

  SpreadsheetApp.getUi().alert(`Guia "${nomeGuia}" criada com sucesso!`);
}

function aplicarCoresCondicionais(sheet) {
  const range = sheet.getRange(2, 2, 999, 1);
  const regras = [];

  for (let marcador in CORES_MARCADORES) {
    let corFundo = CORES_MARCADORES[marcador];
    let corFonte = "#000000";

    if (marcador === "Medicamento - Emergencial") {
      corFonte = "#FFFFFF";
    }

    let regra = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(marcador)
      .setBackground(corFundo)
      .setFontColor(corFonte)
      .setRanges([range])
      .build();
    
    regras.push(regra);
  }

  sheet.setConditionalFormatRules(regras);
}