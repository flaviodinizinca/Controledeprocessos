/**
 * SCRIPT DE DISTRIBUIÇÃO AUTOMÁTICA (DE/PARA)
 * Deve ser executado da Planilha de CONTROLE DE PROCESSOS.
 * ATUALIZADO: Ignora usuários que sejam Saneadores (para evitar duplicidade).
 */

function executarDistribuicaoToFor() {
  const ssControle = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ACESSAR DADOS DA 'TOFOR' (LOCAL)
  const guiaToFor = ssControle.getSheetByName("ToFor");
  if (!guiaToFor) {
    SpreadsheetApp.getUi().alert("Erro: A guia 'ToFor' não foi encontrada.");
    return;
  }
  
  const dadosToFor = guiaToFor.getDataRange().getValues();
  if (dadosToFor.length <= 1) {
    SpreadsheetApp.getUi().alert("A guia 'ToFor' está vazia.");
    return;
  }
  const linhasToFor = dadosToFor.slice(1); // Remove cabeçalho

  // 2. CONFIGURAÇÕES DE IDs
  const ID_PLANILHA_USUARIOS = "1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA";
  const ID_PLANILHA_SANEAMENTO = "1TxyCWwg9IBZpEh9g6E_PgNUx5ucR_CwlTCaS_eXihTs"; // Nova planilha

  // 3. OBTER LISTA DE EXCLUSÃO (SANEADORES)
  // O script acessa a nova planilha apenas para saber quem IGNORAR aqui.
  let listaSaneadores = [];
  try {
    const ssSan = SpreadsheetApp.openById(ID_PLANILHA_SANEAMENTO);
    const guiaConfigSan = ssSan.getSheetByName("Config_Saneamento");
    if (guiaConfigSan) {
      // Pega logins da coluna A
      listaSaneadores = guiaConfigSan.getRange("A2:A").getValues().flat().map(String);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Aviso: Não consegui ler a lista de Saneadores para filtrar. Verifique o ID da planilha nova.");
    return;
  }

  // 4. MAPEAR NOMES DE USUÁRIOS
  let ssUsers;
  try {
    ssUsers = SpreadsheetApp.openById(ID_PLANILHA_USUARIOS);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao abrir planilha de Usuários.");
    return;
  }
  const guiaUsers = ssUsers.getSheetByName("User_SEI");
  const dadosUsers = guiaUsers.getDataRange().getValues();
  const mapaUsuarios = {};
  
  for (let i = 1; i < dadosUsers.length; i++) {
    const nome = dadosUsers[i][0]; 
    const login = dadosUsers[i][1];
    if (login && nome) {
      const pNome = nome.split(" ")[0].trim();
      const formatado = pNome.charAt(0).toUpperCase() + pNome.slice(1).toLowerCase();
      mapaUsuarios[login] = formatado;
    }
  }

  // 5. PROCESSAR DISTRIBUIÇÃO (COM FILTRO)
  let criados = 0;
  let distribuidos = 0;
  let ignoradosSaneamento = 0;

  linhasToFor.forEach(linha => {
    const processo = linha[0];    
    const usuarioLogin = String(linha[1]).trim();   
    const marcador = linha[2];       
    const especificacao = linha[3];  

    // VERIFICAÇÃO IMPORTANTE:
    // Se o login estiver na lista de Saneadores, PULA este registro.
    if (listaSaneadores.includes(usuarioLogin)) {
      ignoradosSaneamento++;
      return; // Sai desta iteração e vai para a próxima linha
    }

    if (processo && usuarioLogin) {
      const nomeGuia = mapaUsuarios[usuarioLogin];

      if (nomeGuia) {
        let abaComprador = ssControle.getSheetByName(nomeGuia);

        // Se a aba não existe, cria (Estrutura Padrão de Comprador)
        if (!abaComprador) {
          // Usa a função do arquivo 02_Guias.gs
          // Note que aqui só criamos "PADRAO", pois os saneadores já foram filtrados
          abaComprador = criarGuiaComprador(nomeGuia, "PADRAO");
          criados++;
        }

        // Monta a linha Padrão (A=Processo, B=Marcador, C=Espec)
        const novaLinha = [
          processo,   
          marcador,      
          especificacao  
        ];

        abaComprador.appendRow(novaLinha);
        distribuidos++;
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    `Distribuição Concluída (Controle de Processos)!\n\n` +
    `🆕 Guias Criadas: ${criados}\n` +
    `📝 Processos Distribuídos: ${distribuidos}\n` +
    `🚫 Ignorados (São Saneadores): ${ignoradosSaneamento}`
  );
}