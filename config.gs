/**
 * ============================================================================
 * Arquivo: Config.gs
 * Descrição:
 *   Centraliza as configurações do projeto (nomes de abas, colunas, etc.)
 *   Assim, se algo mudar na planilha, você só ajusta aqui.
 * ============================================================================
 */

/**
 * Objeto de configuração principal.
 * Cada linha abaixo tem um comentário explicando o que faz na prática.
 */
const CONFIG = {
  // Nome da aba onde está o cadastro de produtos para a sidebar.
  // Ex.: aba com os produtos em estoque.
  NOME_ABA_ESTOQUE: 'ESTOQUE',

  // Número da coluna (índice numérico) onde estão os nomes dos produtos na aba ESTOQUE.
  // 1 = Coluna A, 2 = Coluna B, 3 = Coluna C...
  // Aqui: produtos estão na coluna B da aba ESTOQUE.
  COLUNA_PRODUTOS_ESTOQUE: 2,

  // Primeira linha da aba ESTOQUE que contém produtos.
  // Se a linha 1 for cabeçalho, geralmente os dados começam na linha 2.
  PRIMEIRA_LINHA_PRODUTOS_ESTOQUE: 2,

  // Nome do arquivo HTML que será usado como sidebar.
  // Deve ser igual ao nome do arquivo HTML criado no projeto (ex.: "SidebarProdutos.html").
  NOME_ARQUIVO_HTML_SIDEBAR: 'SidebarProdutos',

  // Número da coluna de destino onde a composição de produtos será escrita.
  // 5 = Coluna E, então a sidebar SEMPRE vai preencher a coluna E.
  COLUNA_DESTINO_COMPOSICAO: 5,

  // Primeira linha permitida para preenchimento na coluna de destino.
  // Ex.: se a linha 1 e 2 são cabeçalhos, começar a partir da linha 3.
  LINHA_INICIAL_DESTINO: 3
};
