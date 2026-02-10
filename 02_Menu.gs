function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Automação SECOM')
    .addItem('Criar Guia Abertura', 'criarGuiaAberturaProcessos') // Antiga Abertura
    .addItem('Criar Guia IRP', 'criarGuiaIRP')
    .addItem('Criar Guia Comprador', 'criarGuiaComprador') // NOVA
    .addSeparator()
    .addItem('🔍 Buscar Informações (Pelo Nº Processo)', 'buscarInformacoesProcesso')
    .addSeparator()
    .addItem('📤 Enviar Log de Justificativa', 'enviarLogJustificativa')
    .addItem('📤 Enviar Log Requisitante', 'enviarLogRequisitante')
    .addToUi();
}