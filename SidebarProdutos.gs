/**
 * ============================================================================
 * Arquivo: SidebarProdutos.gs
 * Descrição:
 *   - Cria menu no Google Sheets
 *   - Abre sidebar com lista de produtos
 *   - Recebe itens selecionados da sidebar e preenche a coluna E (linha ativa)
 * ============================================================================
 */

/**
 * Cria o menu personalizado ao abrir a planilha.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ferramentas de Produtos')
    .addItem('Selecionar Produtos (Sidebar)', 'mostrarSidebarProdutos')
    .addToUi();
}

/**
 * Exibe a barra lateral de seleção de produtos.
 * Carrega os produtos da aba definida em CONFIG.NOME_ABA_ESTOQUE.
 */
function mostrarSidebarProdutos() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaEstoque = ss.getSheetByName(CONFIG.NOME_ABA_ESTOQUE);

  if (!abaEstoque) {
    ui.alert(`Aba "${CONFIG.NOME_ABA_ESTOQUE}" não encontrada. Verifique o nome na Config.`);
    return;
  }

  const ultimaLinha = abaEstoque.getLastRow();
  if (ultimaLinha < CONFIG.PRIMEIRA_LINHA_PRODUTOS_ESTOQUE) {
    ui.alert('Nenhum produto encontrado na aba de estoque.');
    return;
  }

  // Pega a coluna dos produtos (ex.: B) da primeira linha de dados até a última.
  const rangeProdutos = abaEstoque.getRange(
    CONFIG.PRIMEIRA_LINHA_PRODUTOS_ESTOQUE,
    CONFIG.COLUNA_PRODUTOS_ESTOQUE,
    ultimaLinha - CONFIG.PRIMEIRA_LINHA_PRODUTOS_ESTOQUE + 1,
    1
  );

  const dadosEstoque = rangeProdutos.getValues();
  const produtos = dadosEstoque
    .map(row => row[0])
    .filter(valor => valor && valor.toString().trim() !== '');

  if (produtos.length === 0) {
    ui.alert('Não há produtos preenchidos na coluna configurada da aba ESTOQUE.');
    return;
  }

  // Passa a lista de produtos para o template HTML.
  const htmlTemplate = HtmlService.createTemplateFromFile(CONFIG.NOME_ARQUIVO_HTML_SIDEBAR);
  htmlTemplate.produtos = produtos;

  const htmlOutput = htmlTemplate.evaluate()
    .setTitle('Seleção de Produtos')
    .setWidth(320)
    .setHeight(520);

  ui.showSidebar(htmlOutput);
}

/**
 * Recebe os itens selecionados na sidebar e preenche a COLUNA E
 * na linha atualmente ativa (desde que seja >= linha 3).
 *
 * Estrutura esperada de "itens" vindo do HTML:
 * [
 *   { produto: "Produto A", quantidade: 2 },
 *   { produto: "Produto B", quantidade: 1 }
 * ]
 */
function preencherColunaComProdutos(itens) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaAtiva = ss.getActiveSheet();
  const celulaAtiva = abaAtiva.getActiveCell();
  const linhaAtiva = celulaAtiva.getRow();

  // Garante que a linha ativa é permitida (>= linha inicial configurada).
  if (linhaAtiva < CONFIG.LINHA_INICIAL_DESTINO) {
    ui.alert(
      `Selecione uma célula na linha ${CONFIG.LINHA_INICIAL_DESTINO} ou abaixo ` +
      `para preencher a composição de produtos.`
    );
    return;
  }

  if (!itens || itens.length === 0) {
    ui.alert('Nenhum item selecionado na sidebar. Nenhuma alteração foi feita.');
    return;
  }

  // Monta a string no formato: "Produto A + Produto A + Produto B"
  let stringParaPreencher = '';

  itens.forEach(item => {
    const produto = item.produto;
    const quantidade = Number(item.quantidade) || 0;

    for (let j = 0; j < quantidade; j++) {
      if (stringParaPreencher.length > 0) {
        stringParaPreencher += ' + ';
      }
      stringParaPreencher += produto;
    }
  });

  const celulaDestino = abaAtiva.getRange(linhaAtiva, CONFIG.COLUNA_DESTINO_COMPOSICAO);

  if (stringParaPreencher) {
    celulaDestino.setValue(stringParaPreencher);
    ui.alert(
      `A célula ${celulaDestino.getA1Notation()} foi preenchida com:\n\n` +
      stringParaPreencher
    );
  } else {
    celulaDestino.setValue('');
    ui.alert('Nenhum produto válido foi informado. A célula foi limpa.');
  }
}

/**
 * Helper para incluir arquivos HTML no Apps Script.
 * Permite usar <?!= include('Arquivo') ?> dentro de templates HTML.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
