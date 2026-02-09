/**
 * SCRIPT DE DISTRIBUIÇÃO AUTOMÁTICA
 * Local: Hospedado na Planilha de Controle de Processos (ID: 15W847YN...)
 */
function executarDistribuicaoToFor() {
  const ssControle = SpreadsheetApp.getActiveSpreadsheet();
  
  // IDs das planilhas externas
  const idDashboard = "1_Qe3fS-j1B0pM5GzS7z-Fh7v-C-fG7k-G"; // Substitua pelo ID real da sua planilha Dashboard
  const idPlanilhaUsuarios = "1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA";

  // 1. ACESSAR DADOS DA TOFOR NO DASHBOARD
  const ssDash = SpreadsheetApp.openById(idDashboard);
  const guiaToFor = ssDash.getSheetByName("ToFor");
  
  if (!guiaToFor) {
    SpreadsheetApp.getUi().alert("Erro: Aba 'ToFor' não encontrada na planilha Dashboard.");
    return;
  }
  const dadosToFor = guiaToFor.getDataRange().getValues();
  const linhasToFor = dadosToFor.slice(1); // Remove cabeçalho

  // 2. CARREGAR MAPEAMENTO DE USUÁRIOS (LOGIN -> PRIMEIRO NOME)
  const ssUsers = SpreadsheetApp.openById(idPlanilhaUsuarios);
  const guiaUsers = ssUsers.getSheetByName("User_SEI");
  const dadosUsers = guiaUsers.getDataRange().getValues();
  const mapaUsuarios = {};

  for (let i = 1; i < dadosUsers.length; i++) {
    const nomeCompleto = dadosUsers[i][0];
    const login = dadosUsers[i][1];
    if (login && nomeCompleto) {
      // Extrai apenas o primeiro nome e formata
      const primeiroNome = nomeCompleto.split(" ")[0].trim();
      const nomeFormatado = primeiroNome.charAt(0).toUpperCase() + primeiroNome.slice(1).toLowerCase();
      mapaUsuarios[login] = nomeFormatado;
    }
  }

  // 3. PROCESSAR A DISTRIBUIÇÃO
  let contagem = 0;

  linhasToFor.forEach(linha => {
    const numProcesso = linha[0];    // Coluna A (Processo)
    const loginUsuario = linha[1];   // Coluna B (Usuario/Login)
    const marcador = linha[2];       // Coluna C (Marcador)
    const especificacao = linha[3];  // Coluna D (Especificação)

    const nomeGuia = mapaUsuarios[loginUsuario];

    if (nomeGuia) {
      let abaComprador = ssControle.getSheetByName(nomeGuia);

      // Se a aba não existe, cria e ativa as configurações do 02_Guias.gs
      if (!abaComprador) {
        abaComprador = ssControle.insertSheet(nomeGuia);
        // CHAMA A FUNÇÃO EXISTENTE NO SEU PROJETO PARA APLICAR VALIDAÇÕES E CORES
        if (typeof configurarEstruturaGuia === 'function') {
          configurarEstruturaGuia(abaComprador, ssControle);
        }
      }

      // Localiza a primeira linha vazia após o cabeçalho
      // Insere: Processo (A), Marcador (B), Especificação (C)
      abaComprador.appendRow([numProcesso, marcador, especificacao]);
      contagem++;
    }
  });

  SpreadsheetApp.getUi().alert("Processo concluído! " + contagem + " registros distribuídos.");
}

/**
 * Adiciona o botão de execução ao menu para facilitar o acesso
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Distribuição')
    .addItem('Executar Distribuição ToFor', 'executarDistribuicaoToFor')
    .addToUi();
}