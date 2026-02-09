/**
 * Cria o menu personalizado ao abrir a planilha.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestão SECOM')
    .addItem('Nova Guia de Comprador', 'criarGuiaComprador')
    .addToUi();
}