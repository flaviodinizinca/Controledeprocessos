function buscarInformacoesProcesso() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const guiaAtiva = ss.getActiveSheet();
  const celulaAtiva = guiaAtiva.getActiveCell();
  const linhaAtual = celulaAtiva.getRow();
  
  // Ignora cabeçalho
  if (linhaAtual < 2) {
    SpreadsheetApp.getUi().alert("Selecione uma linha com número de processo.");
    return;
  }

  // Pega o Processo na Coluna A da linha atual
  const numeroProcesso = guiaAtiva.getRange(linhaAtual, 1).getValue();

  if (!numeroProcesso || numeroProcesso === "") {
    SpreadsheetApp.getUi().alert("A célula da coluna 'Processo' (A) está vazia.");
    return;
  }

  // Abre a Planilha Mestra (DB)
  const ssMestra = SpreadsheetApp.openById(ID_PLANILHA_MESTRA);
  // Assume que os dados estão na primeira aba da mestra. Se não, especifique o nome: .getSheetByName("Nome")
  const guiaMestra = ssMestra.getSheets()[0]; 
  
  // Pega todos os dados da mestra para busca (performance melhor que buscar um por um)
  // Ler até a coluna G (Valor Estimado) é suficiente para os dados comuns
  const ultimaLinhaMestra = guiaMestra.getLastRow();
  const dadosMestra = guiaMestra.getRange(2, 1, ultimaLinhaMestra - 1, 7).getValues(); 

  let encontrou = false;
  let dadosEncontrados = [];

  // Loop para encontrar o processo
  for (let i = 0; i < dadosMestra.length; i++) {
    // Coluna 0 do array = Coluna A da planilha (Processo)
    if (String(dadosMestra[i][0]) === String(numeroProcesso)) {
      // Dados encontrados: [Processo, Marcador, Espec, Objeto, Mod, Qtd, Valor]
      dadosEncontrados = dadosMestra[i];
      encontrou = true;
      break;
    }
  }

  if (encontrou) {
    // Preenche a guia ativa com os dados encontrados
    // Array: 0=Proc, 1=Marc, 2=Espec, 3=Obj, 4=Mod, 5=Qtd, 6=Valor
    
    // Preenche colunas B até G (Marcador até Valor Estimado)
    // Range começa na Coluna 2 (B), pega 1 linha e 6 colunas
    const dadosParaEscrever = [[
      dadosEncontrados[1], // Marcador
      dadosEncontrados[2], // Especificação
      dadosEncontrados[3], // Objeto
      dadosEncontrados[4], // Modalidade
      dadosEncontrados[5], // Qtd Itens
      dadosEncontrados[6]  // Valor Estimado
    ]];
    
    guiaAtiva.getRange(linhaAtual, 2, 1, 6).setValues(dadosParaEscrever);
    
    SpreadsheetApp.getUi().alert("Dados importados com sucesso!");
  } else {
    SpreadsheetApp.getUi().alert("Processo não encontrado na base de dados Mestra.");
  }
}