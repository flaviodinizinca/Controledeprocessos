/**
 * 02_Guias.gs
 * Gerencia a criação e estruturação das guias dos compradores.
 */

/**
 * Cria uma nova aba para o comprador com a estrutura atualizada.
 * Acionada manualmente pelo menu (se houver) ou chamada por outros scripts.
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
 * Atualizado: Inclui "Marcador" (Col B) e move "Especificação" para (Col C).
 */
function configurarEstruturaGuia(novaGuia, ss) {
  // --- Definição dos Cabeçalhos ---
  const cabecalhos = [
    "Processo",              // Coluna A (1)
    "Marcador",              // Coluna B (2) - NOVA
    "Especificação",         // Coluna C (3) - Antes era B
    "Objeto",                // Coluna D (4)
    "Modalidade",            // Coluna E (5)
    "Qtd Itens",             // Coluna F (6)
    "Valor Estimado",        // Coluna G (7)
    "Status Atual",          // Coluna H (8)
    "Fase SEI",              // Coluna I (9)
    "Prioridade",            // Coluna J (10)
    "Data Rec. Comprador",   // Coluna K (11)
    "Início Pesquisa",       // Coluna L (12)
    "Fim Pesquisa",          // Coluna M (13)
    "Envio Área Requisitante",   // Coluna N (14)
    "Retorno Área Requisitante", // Coluna O (15)
    "Envio p/ Licitação",        // Coluna P (16)
    "Dias na Pesquisa",          // Coluna Q (17)
    "Dias com Requisitante",     // Coluna R (18)
    "Total dias Comprador",      // Coluna S (19)
    "Observações"                // Coluna T (20)
  ];

  // Aplica cabeçalhos na linha 1
  const rangeCabecalho = novaGuia.getRange(1, 1, 1, cabecalhos.length);
  rangeCabecalho.setValues([cabecalhos]);
  rangeCabecalho.setFontWeight("bold");
  rangeCabecalho.setHorizontalAlignment("center");
  rangeCabecalho.setVerticalAlignment("middle");
  rangeCabecalho.setWrap(true);
  
  // Ajusta altura da linha do cabeçalho
  novaGuia.setRowHeight(1, 45);

  // --- Aplicação de Cores de Fundo ---
  // Grupo 1: Identificação (Processo, Marcador, Espec) - Cinza Claro
  novaGuia.getRange(1, 1, 1, 3).setBackground("#EEEEEE");
  
  // Grupo 2: Detalhes (Objeto, Mod, Qtd, Valor) - Azul Claro
  novaGuia.getRange(1, 4, 1, 4).setBackground("#CFE2F3"); 
  
  // Grupo 3: Status (Status, Fase, Prioridade) - Amarelo Claro
  novaGuia.getRange(1, 8, 1, 3).setBackground("#FFF2CC");
  
  // Grupo 4: Datas (K até P) - Verde Claro
  novaGuia.getRange(1, 11, 1, 6).setBackground("#D9EAD3"); 
  
  // Grupo 5: Cálculos (Q até S) - Vermelho Claro
  novaGuia.getRange(1, 17, 1, 3).setBackground("#F4CCCC");
  
  // Grupo 6: Obs (T) - Vermelho Escuro/Texto Branco
  novaGuia.getRange(1, 20, 1, 1).setBackground("#EA9999").setFontColor("white");

  // --- Congelamento ---
  novaGuia.setFrozenRows(1);
  novaGuia.setFrozenColumns(4); // Congela até a coluna D (Objeto)

  // --- Validação de Dados (Modalidade) ---
  // Tenta buscar a aba de configuração, se não existir, não quebra o script
  const guiaConfig = ss.getSheetByName("Modal_Config");
  if (guiaConfig) {
    const rangeValidacao = guiaConfig.getRange("A2:A");
    const regraValidacao = SpreadsheetApp.newDataValidation()
      .requireValueInRange(rangeValidacao, true)
      .setAllowInvalid(false)
      .build();
    // Aplica na Coluna E (Modalidade) - Linha 2 até o fim
    novaGuia.getRange(2, 5, novaGuia.getMaxRows() - 1, 1).setDataValidation(regraValidacao);
  }

  // --- Ajustes de Largura de Colunas ---
  novaGuia.autoResizeColumns(1, cabecalhos.length);
  
  // Ajustes finos manuais para melhor visualização
  novaGuia.setColumnWidth(1, 140); // Processo
  novaGuia.setColumnWidth(2, 100); // Marcador
  novaGuia.setColumnWidth(3, 200); // Especificação
  novaGuia.setColumnWidth(4, 300); // Objeto (Mais largo)
  novaGuia.setColumnWidth(20, 250); // Observações
}