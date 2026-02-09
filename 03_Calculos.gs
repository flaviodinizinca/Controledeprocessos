/**
 * Gatilho que roda automaticamente ao editar qualquer célula.
 */
function onEdit(e) {
  const range = e.range;
  const guia = range.getSheet();
  const linha = range.getRow();
  const coluna = range.getColumn();
  
  // BLOQUEIO 1: Não rodar se for cabeçalho ou a guia de configuração
  if (linha === 1 || guia.getName() === "Modal_Config") return;

  // BLOQUEIO 2: Monitorar apenas a coluna "Rec. Comprador"
  // Na nova estrutura, "Rec. Comprador" é a Coluna I (índice 9)
  if (coluna === 9) {
    calcularPrazoLimite(guia, linha);
  }
}

/**
 * Calcula a data limite de 60 dias corridos a partir do recebimento.
 */
function calcularPrazoLimite(guia, linha) {
  const cellOrigem = guia.getRange(linha, 9);  // Coluna I (Rec. Comprador)
  const cellDestino = guia.getRange(linha, 19); // Coluna S (Prazo Limite - Nova Coluna)

  const dataRecebimento = cellOrigem.getValue();

  // Verifica se é uma data válida
  if (dataRecebimento instanceof Date) {
    
    const dataLimite = new Date(dataRecebimento);
    
    // Adiciona 60 dias (Corridos, padrão para prazos administrativos longos)
    dataLimite.setDate(dataRecebimento.getDate() + 60);
    
    // Escreve o resultado
    cellDestino.setValue(dataLimite);
    cellDestino.setNumberFormat("dd/MM/yyyy");
    
  } else {
    // Se a data for apagada, limpa o prazo
    cellDestino.clearContent();
  }
}