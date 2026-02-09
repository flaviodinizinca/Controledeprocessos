/**
 * 01_Menu.gs
 * Centraliza os menus da Planilha de Controle de Processos.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🚀 Controle SECOM')
    .addItem('➕ Nova Guia (Comprador)', 'acionarNovaGuiaManual')
    .addSeparator()
    .addItem('⚙️ Distribuir Processos (ToFor)', 'executarDistribuicaoToFor')
    .addToUi();
}

/**
 * Função intermediária para pedir o nome ao usuário
 */
function acionarNovaGuiaManual() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Novo Comprador', 
    'Digite o nome do comprador (será o nome da guia):', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const nomeGuia = response.getResponseText().trim();
    if (nomeGuia) {
      criarGuiaComprador(nomeGuia, "PADRAO"); 
    } else {
      ui.alert('O nome não pode estar vazio.');
    }
  }
}