/**
 * SCRIPT DE DISTRIBUIÇÃO AUTOMÁTICA (DE/PARA)
 * Deve ser executado da Planilha de CONTROLE DE PROCESSOS.
 * Lógica atualizada: Define Saneamento pelo LOGIN do usuário, não apenas pelo marcador.
 */

function executarDistribuicaoToFor() {
  const ssControle = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ACESSAR DADOS DA 'TOFOR' LOCALMENTE
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

  const linhasToFor = dadosToFor.slice(1); 

  // 2. CONEXÃO COM A PLANILHA DE USUÁRIOS E CONFIGURAÇÕES
  const idPlanilhaUsuarios = "1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA";
  let ssUsers;
  try {
    ssUsers = SpreadsheetApp.openById(idPlanilhaUsuarios);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erro ao abrir planilha de Usuários/Configuração.");
    return;
  }

  // A) Carregar Nomes (User_SEI)
  const guiaUsers = ssUsers.getSheetByName("User_SEI");
  const dadosUsers = guiaUsers.getDataRange().getValues();
  const mapaUsuarios = {}; // Login -> Nome Formatado
  
  for (let i = 1; i < dadosUsers.length; i++) {
    const nome = dadosUsers[i][0]; 
    const login = dadosUsers[i][1];
    if (login && nome) {
      const pNome = nome.split(" ")[0].trim();
      const formatado = pNome.charAt(0).toUpperCase() + pNome.slice(1).toLowerCase();
      mapaUsuarios[login] = formatado;
    }
  }

  // B) Carregar Lista de Saneadores (Config_Saneamento)
  const guiaSaneadores = ssUsers.getSheetByName("Config_Saneamento");
  let listaSaneadores = [];
  
  if (guiaSaneadores) {
    const dadosSan = guiaSaneadores.getDataRange().getValues();
    // Assume que os logins estão na Coluna A
    for (let i = 1; i < dadosSan.length; i++) {
      const loginSan = dadosSan[i][0]; // Coluna A
      if (loginSan) {
        listaSaneadores.push(String(loginSan).trim());
      }
    }
  } else {
    SpreadsheetApp.getUi().alert("Aviso: Guia 'Config_Saneamento' não encontrada na planilha de usuários.");
  }

  // 3. PROCESSAR DISTRIBUIÇÃO
  let criados = 0;
  let distribuidos = 0;
  let saneamentoCount = 0;

  linhasToFor.forEach(linha => {
    const processo = linha[0];    
    const usuarioLogin = linha[1];   
    const marcador = linha[2]; // Mantemos apenas para registro se for comprador
    const especificacao = linha[3];  

    if (processo && usuarioLogin) {
      const nomeBase = mapaUsuarios[usuarioLogin];

      if (nomeBase) {
        let nomeAbaFinal = nomeBase;
        
        // --- NOVA LÓGICA DE DECISÃO ---
        // Verifica se o LOGIN está na lista de saneadores
        let isSaneamento = listaSaneadores.includes(String(usuarioLogin).trim());
        
        let novaLinha = [];

        if (isSaneamento) {
          // É Saneador: Força estrutura de Saneamento
          nomeAbaFinal = nomeBase + " (Saneamento)";
          saneamentoCount++;
          
          novaLinha = [
            processo,           // A: Processo
            new Date(),         // B: Data Chegada (Hoje)
            "",                 // C: Protocolo
            "",                 // D: Parecer
            especificacao,      // E: Objeto (Vem da Especificação)
            "",                 // F: Célula
            "",                 // G: Modalidade
            "",                 // H: Data Status
            "NÃO",              // I: Encerrado?
            "",                 // J: Localização
            "A Iniciar"         // K: Status
          ];
        } else {
          // Não é Saneador: Estrutura Padrão de Comprador
          novaLinha = [
            processo,      // A
            marcador,      // B (Usa o que vier no SEI, ou vazio)
            especificacao  // C
          ];
        }

        // --- CRIAÇÃO/OBTENÇÃO DA ABA ---
        let abaDestino = ssControle.getSheetByName(nomeAbaFinal);
        
        if (!abaDestino) {
          // Cria usando a função do 02_Guias
          abaDestino = criarGuiaComprador(nomeAbaFinal, isSaneamento ? "SANEAMENTO" : "PADRAO");
          criados++;
        }

        // --- INSERÇÃO ---
        abaDestino.appendRow(novaLinha);
        distribuidos++;
      }
    }
  });

  SpreadsheetApp.getUi().alert(
    `Distribuição Concluída!\n\n` + 
    `🆕 Abas Criadas: ${criados}\n` + 
    `📝 Total Processos: ${distribuidos}\n` + 
    `🛠️ Identificados como Saneamento: ${saneamentoCount}`
  );
}