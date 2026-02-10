// --- CONFIGURAÇÕES GERAIS ---

// ID da Planilha Mestra (Onde estão os dados originais - DB)
const ID_PLANILHA_MESTRA = "1n6l2ofxEvQTrZ49IY7b30U_dcUqb-MuAbVaW890S6ng";

// LISTA DE MARCADORES (Para validação da Coluna B e cores)
const LISTA_MARCADORES = [
  "Assuntos Pessoais e Transitórios",
  "Consumo",
  "Consumo - Emergencial",
  "Medicamento",
  "Medicamento - Emergencial",
  "Permanente",
  "Permanente Emenda Parlamentar",
  "SANEAMENTO",
  "Secretaria",
  "Serviços",
  "Secretaria- Prioridade Estratégicas COAGE - DESATIVADO"
];

// LISTA DE MODALIDADES (Para validação da Coluna E)
// Adicione ou remova itens conforme a necessidade do INCA
const LISTA_MODALIDADES = [
  "Pregão Eletrônico",
  "Dispensa de Licitação (Art. 75)",
  "Inexigibilidade (Art. 74)",
  "Concorrência",
  "Adesão à ARP",
  "Chamamento Público",
  "Credenciamento",
  "Não se aplica"
];

// CORES DOS MARCADORES
const CORES_MARCADORES = {
  "Assuntos Pessoais e Transitórios": "#FFDAB9", // Laranja pastel
  "Consumo": "#FFFF00", // Amarelo
  "Consumo - Emergencial": "#FF0000", // Vermelho
  "Medicamento": "#00FF00", // Verde
  "Medicamento - Emergencial": "#000000", // Preto
  "Permanente": "#0000FF", // Azul
  "Permanente Emenda Parlamentar": "#FF4500", // Laranja brilhante
  "SANEAMENTO": "#808080", // Cinza
  "Secretaria": "#FFFFFF", // Branco
  "Serviços": "#FF69B4", // Rosa Choque
  "Secretaria- Prioridade Estratégicas COAGE - DESATIVADO": "#D3D3D3"
};

// MAPEAMENTO DAS COLUNAS DA PLANILHA MESTRA
const COLUNAS_MESTRA = {
  PROCESSO: 1,           // A
  MARCADOR: 2,           // B
  ESPECIFICACAO: 3,      // C
  OBJETO: 4,             // D
  MODALIDADE: 5,         // E
  QTD_ITENS: 6,          // F
  VALOR_ESTIMADO: 7,     // G
  DATA_JUSTIFICATIVA: 14, // N
  JUSTIFICATIVA: 15,      // O
  ACAO: 16,               // P
  ENVIO_REQ: 18,          // R
  RETORNO_REQ: 19         // S
};