/**
 * SCRIPT DE DISTRIBUIÇÃO AUTOMÁTICA (DE/PARA)
 * Deve ser executado da Planilha de CONTROLE DE PROCESSOS.
 * A guia 'ToFor' deve estar nesta mesma planilha.
 */

function executarDistribuicaoToFor() {
  const ssControle = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ACESSAR DADOS DA 'TOFOR' LOCALMENTE
  const guiaToFor = ssControle.getSheetByName("ToFor");
  
  if (!guiaToFor) {
    SpreadsheetApp.getUi().alert("Erro: A guia 'ToFor' não foi encontrada nesta planilha.");
    return;
  }
  
  // Pega todos os dados da ToFor (Assume cabeçalho na linha 1)
  const dadosToFor = guiaToFor.getDataRange().getValues();
  // Se só tiver cabeçalho, para.
  if (dadosToFor.length <= 1) {
    SpreadsheetApp.getUi().alert("A guia 'ToFor' está vazia (apenas cabeçalho).");
    return;
  }

  const linhasToFor = dadosToFor.slice(1); // Remove cabeçalho da matriz

  // 2. CARREGAR MAPEAMENTO DE USUÁRIOS (User_SEI EXTERNO)
  const idPlanilhaUsuarios = "1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA";
  let ssUsers;
  try {
    ssUsers = SpreadsheetApp.openById(idPlanilhaUsuarios);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao abrir planilha de Usuários (ID incorreto ou sem permissão).");
    return;
  }

  const guiaUsers = ssUsers.getSheetByName("User_SEI");
  const dadosUsers = guiaUsers.getDataRange().getValues();
  const mapaUsuarios = {};
  
  // Cria mapa: Login -> Nome Formatado
  for (let i = 1; i < dadosUsers.length; i++) {
    const nomeCompleto = dadosUsers[i][0]; // Coluna A: Nome
    const login = dadosUsers[i][1];        // Coluna B: Login
    
    if (login && nomeCompleto) {
      // Pega o primeiro nome, capitaliza e remove espaços extras
      const primeiroNome = nomeCompleto.split(" ")[0].trim();
      const nomeFormatado = primeiroNome.charAt(0).toUpperCase() + primeiroNome.slice(1).toLowerCase();
      mapaUsuarios[login] = nomeFormatado;
    }
  }

  // 3. PROCESSAR A DISTRIBUIÇÃO
  let criados = 0;
  let distribuidos = 0;
  let erros = 0;

  linhasToFor.forEach(linha => {
    // Mapeamento das colunas da ToFor
    // A=Processo, B=Usuario, C=Marcador, D=Especificação
    const numProcesso = linha[0];    
    const usuarioLogin = linha[1];   
    const marcador = linha[2];       
    const especificacao = linha[3];  

    if (numProcesso && usuarioLogin) {
      // Busca o nome da guia pelo login
      const nomeGuia = mapaUsuarios[usuarioLogin];

      if (nomeGuia) {
        let abaComprador = ssControle.getSheetByName(nomeGuia);

        // Se a aba não existe, cria e CONFIGURA usando a função do 02_Guias.gs
        if (!abaComprador) {
          abaComprador = ssControle.insertSheet(nomeGuia);
          
          // Verifica se a função de configuração existe antes de chamar
          if (typeof configurarEstruturaGuia === 'function') {
            configurarEstruturaGuia(abaComprador, ssControle);
            criados++;
          }
        }

        // Prepara a linha para inserção seguindo a estrutura do 02_Guias
        // Coluna A: Processo
        // Coluna B: Marcador (NOVO)
        // Coluna C: Especificação (DESLOCADO)
        // Coluna D em diante: Vazio
        const novaLinha = [
          numProcesso,   // A
          marcador,      // B
          especificacao  // C
        ];

        // Adiciona na próxima linha vazia
        abaComprador.appendRow(novaLinha);
        distribuidos++;
      } else {
        // Login não encontrado no mapa
        console.log(`Login não encontrado: ${usuarioLogin}`);
        erros++;
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    `Distribuição Concluída!\n\n` +
    `🆕 Guias Criadas: ${criados}\n` +
    `📝 Processos Distribuídos: ${distribuidos}\n` +
    `⚠️ Logins não encontrados: ${erros}`
  );
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Configuração')
    .addItem('Executar Distribuição ToFor', 'executarDistribuicaoToFor')
    .addToUi();
}