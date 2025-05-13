/**
 * Função para atualizar a aba updatedVariantsPriceTable
 * conforme o código original fornecido inicialmente.
 */

function button0_updatePriceTableSheet_V3() {
  if (typeof DEBUG === 'undefined') { var DEBUG = true; }
  if (DEBUG) {Logger.log('Iniciando o processo de atualização da aba updatedVariantsPriceTable'); console.log('Iniciando o processo de atualização da aba updatedVariantsPriceTable');}

  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDestino = ss.getSheetByName('updatedVariantsPriceTable');
  var logs = [];
  logs.push('Iniciando o processo de atualização.');

  if (!sheetDestino) {
    ui.alert('A aba "updatedVariantsPriceTable" não foi encontrada.');
    logs.push('A aba de destino não foi encontrada e o processo foi interrompido.');
    if (DEBUG) {Logger.log('A aba de destino não foi encontrada e o processo foi interrompido.'); console.log('A aba de destino não foi encontrada e o processo foi interrompido.');}
    return;
  }
  logs.push('Aba de destino encontrada: updatedVariantsPriceTable');
  if (DEBUG) {Logger.log('Aba de destino encontrada: updatedVariantsPriceTable'); console.log('Aba de destino encontrada: updatedVariantsPriceTable');}

  var responseFacil = ui.alert(
    'Atualização de Preços',
    'Você deseja atualizar os preços para a FACIL PERSIANAS?',
    ui.ButtonSet.YES_NO
  );
  logs.push('Resposta sobre FACIL PERSIANAS: ' + (responseFacil == ui.Button.YES ? 'Sim' : 'Não'));
  if (DEBUG) {Logger.log('Resposta sobre FACIL PERSIANAS: ' + (responseFacil == ui.Button.YES ? 'Sim' : 'Não')); console.log('Resposta sobre FACIL PERSIANAS: ' + (responseFacil == ui.Button.YES ? 'Sim' : 'Não'));}

  var DEFAULT_SPREADSHEET_ID = '13m0q4ryo-LACgxY4XRqVz9mOs7-0x3Eq3pE-KXRYBjE';
  var sourceSpreadsheetId = DEFAULT_SPREADSHEET_ID;
  var sourceSheetName = '';

  if (responseFacil == ui.Button.YES) {
    sourceSheetName = 'Facil Persianas';
    if (DEBUG) {Logger.log('Usuário selecionou FACIL PERSIANAS.'); console.log('Usuário selecionou FACIL PERSIANAS.');}
  } else {
    var responseP2GO = ui.alert(
      'Atualização de Preços',
      'Você gostaria de atualizar o preço da P2GO usando a tabela da P2GO?',
      ui.ButtonSet.YES_NO
    );
    logs.push('Resposta sobre P2GO: ' + (responseP2GO == ui.Button.YES ? 'Sim' : 'Não'));
    if (DEBUG) {Logger.log('Resposta sobre P2GO: ' + (responseP2GO == ui.Button.YES ? 'Sim' : 'Não')); console.log('Resposta sobre P2GO: ' + (responseP2GO == ui.Button.YES ? 'Sim' : 'Não'));}

    if (responseP2GO == ui.Button.YES) {
      sourceSheetName = 'P2GO';
      if (DEBUG) {Logger.log('Usuário selecionou P2GO.'); console.log('Usuário selecionou P2GO.');}
    } else {
      var sheetResponse = ui.prompt('Insira o nome da aba que deseja usar:', ui.ButtonSet.OK_CANCEL);
      if (sheetResponse.getSelectedButton() != ui.Button.OK) {
        ui.alert('Ação cancelada.');
        logs.push('Usuário cancelou a ação ao inserir o nome da aba.');
        if (DEBUG) {Logger.log('Usuário cancelou a ação ao inserir o nome da aba.'); console.log('Usuário cancelou a ação ao inserir o nome da aba.');}
        return;
      }
      sourceSheetName = sheetResponse.getResponseText().trim();
      if (!sourceSheetName) {
        ui.alert('Nome da aba não fornecido. Ação cancelada.');
        logs.push('Nome da aba não fornecido. Processo interrompido.');
        if (DEBUG) {Logger.log('Nome da aba não fornecido. Processo interrompido.'); console.log('Nome da aba não fornecido. Processo interrompido.');}
        return;
      }
      logs.push('Nome da aba fornecido: ' + sourceSheetName);
      if (DEBUG) {Logger.log('Nome da aba fornecido: ' + sourceSheetName); console.log('Nome da aba fornecido: ' + sourceSheetName);}

      var responseDefaultSheet = ui.alert(
        'Confirmação de Planilha',
        'A aba "' + sourceSheetName + '" está na planilha padrão?',
        ui.ButtonSet.YES_NO
      );
      logs.push('A aba está na planilha padrão? ' + (responseDefaultSheet == ui.Button.YES ? 'Sim' : 'Não'));
      if (DEBUG) {Logger.log('A aba está na planilha padrão? ' + (responseDefaultSheet == ui.Button.YES ? 'Sim' : 'Não')); console.log('A aba está na planilha padrão? ' + (responseDefaultSheet == ui.Button.YES ? 'Sim' : 'Não'));}

      if (responseDefaultSheet == ui.Button.NO) {
        var docResponse = ui.prompt(
          'Insira o ID do documento de origem:',
          'Exemplo de ID: 1abcD_EfgHiJkLmNoPqRsTuVwXyZ',
          ui.ButtonSet.OK_CANCEL
        );
        if (docResponse.getSelectedButton() != ui.Button.OK) {
          ui.alert('Ação cancelada.');
          logs.push('Usuário cancelou a ação ao inserir o ID do documento.');
          if (DEBUG) {Logger.log('Usuário cancelou a ação ao inserir o ID do documento.'); console.log('Usuário cancelou a ação ao inserir o ID do documento.');}
          return;
        }
        sourceSpreadsheetId = docResponse.getResponseText().trim();
        if (!sourceSpreadsheetId) {
          ui.alert('ID do documento não fornecido. Ação cancelada.');
          logs.push('ID do documento não fornecido. Processo interrompido.');
          if (DEBUG) {Logger.log('ID do documento não fornecido. Processo interrompido.'); console.log('ID do documento não fornecido. Processo interrompido.');}
          return;
        }
        logs.push('ID do documento fornecido: ' + sourceSpreadsheetId);
        if (DEBUG) {Logger.log('ID do documento fornecido: ' + sourceSpreadsheetId); console.log('ID do documento fornecido: ' + sourceSpreadsheetId);}
      } else {
        sourceSpreadsheetId = DEFAULT_SPREADSHEET_ID;
        logs.push('Usando a planilha padrão.');
        if (DEBUG) {Logger.log('Usando a planilha padrão.'); console.log('Usando a planilha padrão.');}
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Processando os dados...', 'Atualização em andamento', 5);
  logs.push('Iniciando a atualização com os parâmetros definidos.');
  if (DEBUG) {Logger.log('Iniciando a atualização com os parâmetros definidos.'); console.log('Iniciando a atualização com os parâmetros definidos.');}

  atualizarPlanilhaComFormatacao(
    sourceSpreadsheetId,
    sourceSheetName,
    sheetDestino,
    logs
  );
}


function atualizarPlanilhaComFormatacao(
  sourceSpreadsheetId,
  sourceSheetName,
  sheetDestino,
  logs
) {
  var ui = SpreadsheetApp.getUi();
  if (typeof DEBUG === 'undefined') { var DEBUG = true; }
  if (DEBUG) {Logger.log('Obtida interface do usuário na função de atualização.'); console.log('Obtida interface do usuário na função de atualização.');}
  try {
    logs.push('Iniciando atualização da planilha.');
    if (DEBUG) {Logger.log('Iniciando atualização da planilha.'); console.log('Iniciando atualização da planilha.');}

    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    logs.push('Documento de origem acessado com sucesso: ' + sourceSpreadsheetId);
    if (DEBUG) {Logger.log('Documento de origem acessado com sucesso: ' + sourceSpreadsheetId); console.log('Documento de origem acessado com sucesso: ' + sourceSpreadsheetId);}

    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      ui.alert('A aba "' + sourceSheetName + '" não foi encontrada no documento de origem.');
      logs.push('A aba de origem não foi encontrada. Processo interrompido.');
      if (DEBUG) {Logger.log('A aba de origem não foi encontrada. Processo interrompido.'); console.log('Aba de origem não encontrada.');}
      return;
    }
    logs.push('Aba de origem encontrada: ' + sourceSheetName);
    if (DEBUG) {Logger.log('Aba de origem encontrada: ' + sourceSheetName); console.log('Aba de origem encontrada: ' + sourceSheetName);}

    var COLUMN_START = 'A';
    var COLUMN_END = 'E';
    var lastRow = findLastRowWithData(sourceSheet, COLUMN_START, COLUMN_END);
    logs.push('Última linha preenchida encontrada: ' + lastRow);
    if (DEBUG) {Logger.log('Última linha preenchida encontrada: ' + lastRow); console.log('Última linha preenchida: ', lastRow);}

    if (lastRow < 2) {
      ui.alert('A aba de origem não contém dados suficientes para copiar.');
      logs.push('Dados insuficientes na aba de origem. Processo interrompido.');
      if (DEBUG) {Logger.log('Dados insuficientes na aba de origem. Processo interrompido.'); console.log('Dados insuficientes na aba de origem.');}
      return;
    }

    var lastColumn = sourceSheet.getLastColumn();
    var lastColumnLetter = columnToLetter(lastColumn);
    var sourceRangeNotation = 'A1:' + lastColumnLetter + lastRow;
    logs.push('Intervalo a ser copiado: ' + sourceRangeNotation);
    if (DEBUG) {Logger.log('Intervalo a ser copiado: ' + sourceRangeNotation); console.log('Intervalo a ser copiado: ', sourceRangeNotation);}

    var cellA1 = sheetDestino.getRange('A1');
    var formulaA1 = cellA1.getFormula();
    var valueA1 = cellA1.getValue();
    logs.push('Conteúdo da célula A1 armazenado.');
    if (DEBUG) {Logger.log('Conteúdo da célula A1 armazenado.'); console.log('Conteúdo da célula A1 armazenado.');}

    sheetDestino.clearContents();
    sheetDestino.clearFormats();
    logs.push('Aba de destino limpa.');
    if (DEBUG) {Logger.log('Aba de destino limpa.'); console.log('Aba de destino limpa.');}

    if (formulaA1) {
      cellA1.setFormula(formulaA1);
    } else {
      cellA1.setValue(valueA1);
    }
    logs.push('Conteúdo da célula A1 restaurado.');
    if (DEBUG) {Logger.log('Conteúdo da célula A1 restaurado.'); console.log('Conteúdo da célula A1 restaurado.');}

    var sourceRange = sourceSheet.getRange(sourceRangeNotation);
    var sourceValues = sourceRange.getValues();
    var sourceBackgrounds = sourceRange.getBackgrounds();
    logs.push('Dados e formatações obtidos da aba de origem.');
    if (DEBUG) {Logger.log('Dados e formatações obtidos da aba de origem.'); console.log('Dados e formatações obtidos da aba de origem.');}

    var destinoRange = sheetDestino.getRange('A1').offset(0, 0, sourceValues.length, sourceValues[0].length);
    destinoRange.setValues(sourceValues);
    logs.push('Dados aplicados na aba de destino.');
    if (DEBUG) {Logger.log('Dados aplicados na aba de destino.'); console.log('Dados aplicados na aba de destino.');}

    destinoRange.setBackgrounds(sourceBackgrounds);
    logs.push('Formatações aplicadas na aba de destino.');
    if (DEBUG) {Logger.log('Formatações aplicadas na aba de destino.'); console.log('Formatações aplicadas na aba de destino.');}

    sheetDestino.autoResizeColumns(1, sourceValues[0].length);
    logs.push('Largura das colunas ajustada.');
    if (DEBUG) {Logger.log('Largura das colunas ajustada.'); console.log('Largura das colunas ajustada.');}

    SpreadsheetApp.getActiveSpreadsheet().toast('A aba "updatedVariantsPriceTable" foi atualizada com sucesso!', 'Atualização Concluída', 5);
    logs.push('Processo concluído com sucesso.');
    if (DEBUG) {Logger.log('Processo concluído com sucesso.'); console.log('Processo concluído com sucesso.');}
  } catch (error) {
    ui.alert('Ocorreu um erro durante a atualização: ' + error);
    logs.push('Erro: ' + error);
    if (DEBUG) {Logger.log('Erro: ' + error); console.log('Erro: ', error);}
  } finally {
    logs.forEach(function(log) {
      if (DEBUG) {console.log(log); Logger.log(log);}
    });
  }
}

function findLastRowWithData(sheet, startCol, endCol) {
  if (typeof DEBUG === 'undefined') { var DEBUG = true; }
  var lastRow = sheet.getLastRow();
  if (DEBUG) {Logger.log('Última linha da planilha: ' + lastRow); console.log('Última linha da planilha: ', lastRow);}
  var dataRange = sheet.getRange(startCol + '1:' + endCol + lastRow);
  if (DEBUG) {Logger.log('Intervalo de dados definido: ' + startCol + '1:' + endCol + lastRow); console.log('Intervalo de dados definido: ', startCol + '1:' + endCol + lastRow);}
  var data = dataRange.getValues();
  if (DEBUG) {Logger.log('Valores do intervalo obtidos.'); console.log('Valores do intervalo obtidos.');}

  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    var isRowFilled = row.every(function(cell) {
      return cell !== '';
    });
    if (isRowFilled) {
      if (DEBUG) {Logger.log('Última linha preenchida encontrada na linha: ' + (i + 1)); console.log('Última linha preenchida encontrada na linha: ', (i+1));}
      return i + 1;
    }
  }
  if (DEBUG) {Logger.log('Nenhuma linha preenchida encontrada.'); console.log('Nenhuma linha preenchida encontrada.');}
  return 0;
}

function columnToLetter(column) {
  if (typeof DEBUG === 'undefined') { var DEBUG = true; }
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  if (DEBUG) {Logger.log('Número da coluna convertido para letra: ' + letter); console.log('Número da coluna convertido para letra: ', letter);}
  return letter;
}
















































/********************************************************************************************
 * SUMÁRIO E OPÇÕES DE CONFIGURAÇÃO
 *
 * Neste código, adicionamos opções booleanas no início para configurar como o custo e o preço
 * devem ser escolhidos entre múltiplas linhas que atendem aos critérios.
 * Apenas uma opção por tipo (custo e preço) deve estar definida como true, as demais como false.
 *
 * OPÇÕES PARA SELEÇÃO DO CUSTO (apenas uma true):
 *   useMaxCost   = false; // usar o maior custo possível
 *   useMinCost   = true;  // usar o menor custo possível
 *   useAvgCost   = false; // usar o custo médio
 *
 * OPÇÕES PARA SELEÇÃO DO PREÇO (apenas uma true):
 *   useMaxPrice  = false; // usar o maior preço possível
 *   useMinPrice  = true;  // usar o menor preço possível
 *   useAvgPrice  = false; // usar o preço médio
 *
 * Também ajustamos a geração para vários produtos:
 * Se no intervalo informado várias linhas tiverem o mesmo itemNumber, apenas uma planilha será criada para aquele itemNumber.
 * Ou seja, pegamos o conjunto único de itemNumbers do intervalo, e para cada itemNumber único geramos uma planilha.
 *
 * O script sempre cria novas planilhas separadas, nunca abas na mesma planilha.
 * Cada planilha gerada para um produto contém o PriceGrid e o CostGrid.
 *
 * Ao final da criação do CostGrid, registramos o link na aba "PriceGrids".
 *
 * Este código é autossuficiente, basta colar todo o conteúdo no editor de scripts.
 * Caso queira ajustar logs, defina DEBUG = false.
 *
 ********************************************************************************************/

var DEBUG = true; // Se true, exibe logs. Se false, não exibe.

// OPÇÕES DE CUSTO (apenas uma deve ser true)
var useMaxCost = false;
var useMinCost = true;
var useAvgCost = false;

// OPÇÕES DE PREÇO (apenas uma deve ser true)
var useMaxPrice = false;
var useMinPrice = true;
var useAvgPrice = false;

var DESTINATION_SHEET_NAME = 'updatedVariantsPriceTable';

// Função para criar o menu
function onOpen() {
  const ui = SpreadsheetApp.getUi(); 
  const menu = ui.createMenu('DIY/HD - Automated Functions');
  menu.addItem('Gerar PriceGrid/CostGrid de 1 produto', 'generateGridsForSingleProductPrompt');
  menu.addItem('Gerar PriceGrid/CostGrid de vários produtos (intervalo)', 'generateGridsForMultipleProductsPrompt');
  menu.addItem('ATUALIZA - aba updatedVariantsPriceTable', 'button0_updatePriceTableSheet_V3');
  menu.addToUi();
}

/**
 * Prompt para um produto.
 * Solicita itemNumber e chama generateGridsForSingleProduct(itemNumber).
 */
function generateGridsForSingleProductPrompt() {
  if (DEBUG) {Logger.log("Iniciando generateGridsForSingleProductPrompt"); console.log("Iniciando generateGridsForSingleProductPrompt");}
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Digite o itemNumber do produto para gerar PriceGrid e CostGrid:");
  if (response.getSelectedButton() == ui.Button.CANCEL || response.getResponseText().trim() === "") {
    if (DEBUG) {Logger.log("ItemNumber não fornecido. Encerrando."); console.log("ItemNumber não fornecido. Encerrando.");}
    return;
  }
  var itemNumber = response.getResponseText().trim();
  if (DEBUG) {Logger.log("ItemNumber: " + itemNumber); console.log("ItemNumber: ", itemNumber);}
  generateGridsForSingleProduct(itemNumber);
}

/**
 * Prompt para múltiplos produtos.
 * Solicita um intervalo de linhas (ex: "10-20") na aba updatedVariantsPriceTable.
 * Obtém todos os itemNumbers únicos nesse intervalo e gera uma planilha para cada itemNumber único.
 */
function generateGridsForMultipleProductsPrompt() {
  if (DEBUG) {Logger.log("Iniciando generateGridsForMultipleProductsPrompt"); console.log("Iniciando generateGridsForMultipleProductsPrompt");}
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("Digite o intervalo de linhas (ex: 10-20) na aba '"+DESTINATION_SHEET_NAME+"' para gerar PriceGrid e CostGrid:");
  if (response.getSelectedButton() == ui.Button.CANCEL || response.getResponseText().trim() === "") {
    if (DEBUG) {Logger.log("Intervalo não fornecido. Encerrando."); console.log("Intervalo não fornecido. Encerrando.");}
    return;
  }
  var rangeText = response.getResponseText().trim();
  var match = rangeText.match(/^(\d+)-(\d+)$/);
  if (!match) {
    ui.alert("Formato inválido. Use algo como '10-20' para linhas 10 até 20.");
    return;
  }
  var startRow = parseInt(match[1],10);
  var endRow = parseInt(match[2],10);
  if (startRow > endRow) {
    ui.alert("O valor inicial deve ser menor ou igual ao valor final.");
    return;
  }

  if (DEBUG) {Logger.log("Obtendo itemNumbers do intervalo: " + startRow + "-" + endRow); console.log("Obtendo itemNumbers do intervalo: ", startRow, endRow);}

  var itemNumbers = getItemNumbersFromRange(startRow, endRow);
  if (itemNumbers.length == 0) {
    SpreadsheetApp.getUi().alert("Nenhum itemNumber encontrado no intervalo fornecido.");
    return;
  }

  // Extrair apenas itemNumbers únicos
  var uniqueItemNumbers = Array.from(new Set(itemNumbers));

  // Gerar para cada itemNumber único
  for (var i=0; i<uniqueItemNumbers.length; i++) {
    generateGridsForSingleProduct(uniqueItemNumbers[i]);
  }
}

/**
 * Obtém todos os itemNumbers no intervalo dado da aba updatedVariantsPriceTable.
 */
function getItemNumbersFromRange(startRow, endRow) {
  if (DEBUG) {Logger.log("Iniciando getItemNumbersFromRange: " + startRow + "-" + endRow); console.log("Iniciando getItemNumbersFromRange:", startRow, endRow);}

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DESTINATION_SHEET_NAME);
  if (!sheet) {
    if (DEBUG) {Logger.log("Aba " + DESTINATION_SHEET_NAME + " não encontrada."); console.log("Aba não encontrada.");}
    return [];
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  var itemNumberIndex = headers.indexOf('itemNumber');
  if (itemNumberIndex == -1) {
    if (DEBUG) {Logger.log("Coluna itemNumber não encontrada."); console.log("Coluna itemNumber não encontrada.");}
    return [];
  }

  var lastRow = sheet.getLastRow();
  if (endRow > lastRow) endRow = lastRow;

  var itemNumbers = [];
  for (var i = startRow; i <= endRow; i++) {
    if (i <= 1) continue; 
    var row = data[i-1];
    var val = row[itemNumberIndex];
    if (val && val.toString().trim() !== "") {
      itemNumbers.push(val.toString().trim());
    }
  }
  return itemNumbers;
}

/**
 * Gera PriceGrid e CostGrid para um único produto (itemNumber).
 * Sempre cria uma nova planilha.
 */
function generateGridsForSingleProduct(itemNumber) {
  if (DEBUG) {Logger.log("Iniciando generateGridsForSingleProduct: " + itemNumber); console.log("Iniciando generateGridsForSingleProduct:", itemNumber);}

  var productData = getProductDataByItemNumber(itemNumber);
  var mainSku = productData.mainSku;

  // Criar nova planilha para este produto
  var newSS = SpreadsheetApp.create("Grids_" + mainSku);
  if (DEBUG) {Logger.log("Nova planilha criada: " + newSS.getUrl()); console.log("Nova planilha criada:", newSS.getUrl());}

  // Gerar PriceGrid
  generatePriceGridForProduct(itemNumber, newSS);

  // Gerar CostGrid
  generateCostGridForProduct(itemNumber, productData, newSS);

  if (DEBUG) {Logger.log("generateGridsForSingleProduct concluída."); console.log("generateGridsForSingleProduct concluída.");}
}

/**
 * Gera a aba PriceGrid_mainSku na nova planilha.
 */
function generatePriceGridForProduct(itemNumber, newSS) {
  if (DEBUG) {Logger.log("Iniciando generatePriceGridForProduct para itemNumber: " + itemNumber); console.log("generatePriceGridForProduct:", itemNumber);}
  var productData = getProductDataByItemNumber(itemNumber);
  var mainSku = productData.mainSku;

  var sheetName = "PriceGrid_" + mainSku;
  var existingSheet = newSS.getSheetByName(sheetName);
  if (existingSheet) newSS.deleteSheet(existingSheet);
  var priceSheet = newSS.insertSheet(sheetName);

  var productInfo = productData.model + " " + productData.fabricLaminate + " " + productData.color + " " + productData.support + " " + productData.mainSku;
  priceSheet.getRange("A1").setValue(productInfo);

  var minW = productData.minimumWidthMM;
  var maxW = productData.maximumWidthMM;
  var minH = productData.minimumHeightMM;
  var maxH = productData.maximumHeightMM;

  var widths = [];
  for (var w = minW; w <= maxW; w += 100) { widths.push(w); }

  var heights = [];
  for (var h = minH; h <= maxH; h += 100) { heights.push(h); }

  priceSheet.getRange(1, 2, 1, widths.length).setValues([widths]);
  var heightValues = heights.map(function(val){return [val]});
  priceSheet.getRange(2, 1, heights.length, 1).setValues(heightValues);

  var priceTableData = getPriceTableForMainSku(mainSku);

  var result = [];
  for (var i=0; i<heights.length; i++) {
    var rowPrices = [];
    for (var j=0; j<widths.length; j++) {
      var prices = getAllPricesForDimensions(mainSku, widths[j], heights[i], priceTableData);
      var finalPrice = choosePrice(prices); // escolhe conforme as opções (max, min ou avg)
      var priceString = finalPrice.toFixed(2).replace(',', '.');
      rowPrices.push(priceString);
    }
    result.push(rowPrices);
  }

  priceSheet.getRange(2, 2, heights.length, widths.length).setValues(result);

  if (DEBUG) {Logger.log("PriceGrid gerada com sucesso."); console.log("PriceGrid gerada com sucesso.");}
}

/**
 * Gera a aba CostGrid_mainSku na nova planilha e registra na aba PriceGrids.
 */
function generateCostGridForProduct(itemNumber, productData, newSS) {
  if (DEBUG) {Logger.log("Iniciando generateCostGridForProduct: " + itemNumber); console.log("generateCostGridForProduct:", itemNumber);}
  var mainSku = productData.mainSku;

  var sheetName = "CostGrid_" + mainSku;
  var existingSheet = newSS.getSheetByName(sheetName);
  if (existingSheet) newSS.deleteSheet(existingSheet);
  var costSheet = newSS.insertSheet(sheetName);

  var productInfo = productData.model + " " + productData.fabricLaminate + " " + productData.color + " " + productData.support + " " + mainSku;
  costSheet.getRange("A1").setValue(productInfo);

  var minW = productData.minimumWidthMM;
  var maxW = productData.maximumWidthMM;
  var minH = productData.minimumHeightMM;
  var maxH = productData.maximumHeightMM;

  var widths = [];
  for (var w = minW; w <= maxW; w += 100) { widths.push(w); }

  var heights = [];
  for (var h = minH; h <= maxH; h += 100) { heights.push(h); }

  costSheet.getRange(1, 2, 1, widths.length).setValues([widths]);
  var heightValues = heights.map(function(val){return [val]});
  costSheet.getRange(2, 1, heights.length, 1).setValues(heightValues);

  var priceTableData = getPriceTableForMainSku(mainSku);

  var result = [];
  for (var i=0; i<heights.length; i++) {
    var rowCosts = [];
    for (var j=0; j<widths.length; j++) {
      var costs = getAllCostsForDimensions(mainSku, widths[j], heights[i], priceTableData);
      var finalCost = chooseCost(costs); // escolhe conforme as opções (max, min ou avg)
      var costString = finalCost.toFixed(2).replace(',', '.');
      rowCosts.push(costString);
    }
    result.push(rowCosts);
  }

  costSheet.getRange(2, 2, heights.length, widths.length).setValues(result);

  if (DEBUG) {Logger.log("CostGrid gerada com sucesso."); console.log("CostGrid gerada com sucesso.");}

  // Registrar na aba PriceGrids
  var originalSS = SpreadsheetApp.getActiveSpreadsheet();
  var priceGridsSheet = originalSS.getSheetByName("PriceGrids");
  if (!priceGridsSheet) {
    if (DEBUG) {Logger.log("Aba PriceGrids não encontrada."); console.log("Aba PriceGrids não encontrada.");}
    return;
  }

  var priceGridsData = priceGridsSheet.getDataRange().getValues();
  var headers = priceGridsData[0];

  var colN = headers.indexOf('nº');
  var colModel = headers.indexOf('model');
  var colFabricLaminate = headers.indexOf('fabricLaminate');
  var colColor = headers.indexOf('color');
  var colSupport = headers.indexOf('support');
  var colCodeInAxErp = headers.indexOf('codeInAxErp');
  var colProductCatalogId = headers.indexOf('productCatalogId');
  var colDateCreation = headers.indexOf('date_creation');
  var colCostGrid = headers.indexOf('file_Grid');
  var colPriceGridFullPrice = headers.indexOf('priceGrid_fullPrice');

  var newRowArray = new Array(headers.length).fill("");

  var lastRow = priceGridsSheet.getLastRow();
  var numberSeq = lastRow;

  var codeInAxErp = mainSku;
  var productCatalogId = "";
  var date_creation = new Date();
  var costGridLink = newSS.getUrl();
  var priceGrid_fullPrice = "";

  if (colN > -1) newRowArray[colN] = numberSeq;
  if (colModel > -1) newRowArray[colModel] = productData.model || "";
  if (colFabricLaminate > -1) newRowArray[colFabricLaminate] = productData.fabricLaminate || "";
  if (colColor > -1) newRowArray[colColor] = productData.color || "";
  if (colSupport > -1) newRowArray[colSupport] = productData.support || "";
  if (colCodeInAxErp > -1) newRowArray[colCodeInAxErp] = codeInAxErp || "";
  if (colProductCatalogId > -1) newRowArray[colProductCatalogId] = productCatalogId;
  if (colDateCreation > -1) newRowArray[colDateCreation] = date_creation;
  if (colCostGrid > -1) newRowArray[colCostGrid] = costGridLink;
  if (colPriceGridFullPrice > -1) newRowArray[colPriceGridFullPrice] = priceGrid_fullPrice;

  priceGridsSheet.appendRow(newRowArray);

  if (DEBUG) {Logger.log("Linha adicionada na aba PriceGrids."); console.log("Linha adicionada na aba PriceGrids.");}
}

/**
 * Obtém dados do produto a partir do itemNumber.
 */
function getProductDataByItemNumber(itemNumber) {
  if (DEBUG) {Logger.log("getProductDataByItemNumber: " + itemNumber); console.log("getProductDataByItemNumber:", itemNumber);}

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DESTINATION_SHEET_NAME);
  if (!sheet) {
    throw new Error("Aba não encontrada.");
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var itemNumberIndex = headers.indexOf('itemNumber');
  var modelIndex = headers.indexOf('model');
  var fabricLaminateIndex = headers.indexOf('fabricLaminate');
  var colorIndex = headers.indexOf('color');
  var supportIndex = headers.indexOf('support');
  var mainSkuIndex = headers.indexOf('mainSku');
  var minWidthIndex = headers.indexOf('minimumWidthMM');
  var maxWidthIndex = headers.indexOf('maximumWidthMM');
  var minHeightIndex = headers.indexOf('minimumHeightMM');
  var maxHeightIndex = headers.indexOf('maximumHeightMM');

  var requiredColumns = [itemNumberIndex, modelIndex, fabricLaminateIndex, colorIndex, supportIndex, mainSkuIndex, 
                         minWidthIndex, maxWidthIndex, minHeightIndex, maxHeightIndex];
  if (requiredColumns.includes(-1)) {
    throw new Error("Colunas necessárias não encontradas.");
  }

  var rowsForItem = [];
  for (var i=1; i<data.length; i++) {
    if (data[i][itemNumberIndex].toString().trim() == itemNumber) {
      rowsForItem.push(data[i]);
    }
  }

  if (rowsForItem.length == 0) {
    throw new Error("Nenhuma linha encontrada para o itemNumber informado.");
  }

  var productRow = rowsForItem[0];
  var modelValue = productRow[modelIndex];
  var fabricValue = productRow[fabricLaminateIndex];
  var colorValue = productRow[colorIndex];
  var supportValue = productRow[supportIndex];
  var mainSkuValue = productRow[mainSkuIndex];

  var minWidth = Infinity;
  var maxWidth = -Infinity;
  var minHeight = Infinity;
  var maxHeight = -Infinity;

  for (var r = 0; r < rowsForItem.length; r++) {
    var row = rowsForItem[r];
    var cMinW = parseNumber(row[minWidthIndex]);
    var cMaxW = parseNumber(row[maxWidthIndex]);
    var cMinH = parseNumber(row[minHeightIndex]);
    var cMaxH = parseNumber(row[maxHeightIndex]);

    if (cMinW < minWidth) minWidth = cMinW;
    if (cMaxW > maxWidth) maxWidth = cMaxW;
    if (cMinH < minHeight) minHeight = cMinH;
    if (cMaxH > maxHeight) maxHeight = cMaxH;
  }

  var productData = {
    model: modelValue,
    fabricLaminate: fabricValue,
    color: colorValue,
    support: supportValue,
    mainSku: mainSkuValue,
    minimumWidthMM: minWidth,
    maximumWidthMM: maxWidth,
    minimumHeightMM: minHeight,
    maximumHeightMM: maxHeight,
    headers: headers,
    allData: data,
    itemNumber: itemNumber
  };

  return productData;
}

/**
 * Obtém todas as linhas da updatedVariantsPriceTable que correspondem ao mainSku.
 */
function getPriceTableForMainSku(mainSku) {
  if (DEBUG) {Logger.log("getPriceTableForMainSku: " + mainSku); console.log("getPriceTableForMainSku:", mainSku);}
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DESTINATION_SHEET_NAME);
  if (!sheet) {
    throw new Error("Aba de price table não encontrada.");
  }

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var mainSkuIndex = headers.indexOf('mainSku');
  if (mainSkuIndex == -1) {
    throw new Error("Coluna mainSku não encontrada.");
  }

  var filteredRows = [];
  for (var i=1; i<data.length; i++) {
    if (data[i][mainSkuIndex] == mainSku) {
      filteredRows.push(data[i]);
    }
  }

  return {
    headers: headers,
    rows: filteredRows
  };
}

/**
 * Ao invés de calcular diretamente o menor preço, obtém todos os preços possíveis 
 * e depois escolhe conforme a opção (max, min, avg).
 */
function getAllPricesForDimensions(mainSku, widthMM, heightMM, priceTableData) {
  var headers = priceTableData.headers;
  var rows = priceTableData.rows;

  var idxFields = ['fabricCostPerM2','hardwareCostPerWidthML','hardwareCostPerHeightML','hardwareCostPerUnit',
                   'laborCostPerUnit','wasteAndPackagingPerUnit','taxVatOverCost','markUp','discount',
                   'minimumWidthMM','maximumWidthMM','minimumHeightMM','maximumHeightMM',
                   'minimumChargedWidthMM','minimumChargedHeightMM','minimumChargedAreaM2','maximumAreaM2','correlationHeightByLength'];

  var idx = {};
  idxFields.forEach(function(field){ idx[field] = headers.indexOf(field); });

  var candidates = rows.filter(function(row){
    var minW = parseNumber(row[idx['minimumWidthMM']]);
    var maxW = parseNumber(row[idx['maximumWidthMM']]);
    var minH = parseNumber(row[idx['minimumHeightMM']]);
    var maxH = parseNumber(row[idx['maximumHeightMM']]);
    var maxArea = parseNumber(row[idx['maximumAreaM2']]);
    var corr = parseNumber(row[idx['correlationHeightByLength']]);

    var realArea = (widthMM/1000)*(heightMM/1000);
    var realHByL = heightMM/widthMM;

    return (widthMM>=minW && widthMM<=maxW && heightMM>=minH && heightMM<=maxH && realArea<=maxArea && realHByL<=corr);
  });

  if (candidates.length==0) return [19999.00];

  var prices = candidates.map(function(row){
    var fabric = parseNumber(row[idx['fabricCostPerM2']]);
    var hwW = parseNumber(row[idx['hardwareCostPerWidthML']]);
    var hwH = parseNumber(row[idx['hardwareCostPerHeightML']]);
    var hwU = parseNumber(row[idx['hardwareCostPerUnit']]);
    var labor = parseNumber(row[idx['laborCostPerUnit']]);
    var waste = parseNumber(row[idx['wasteAndPackagingPerUnit']]);
    var tax = normalizePercentage(row[idx['taxVatOverCost']]);
    var mkUp = parseNumber(row[idx['markUp']]);
    var disc = normalizePercentage(row[idx['discount']]);

    var minCW = parseNumber(row[headers.indexOf('minimumChargedWidthMM')]);
    var minCH = parseNumber(row[headers.indexOf('minimumChargedHeightMM')]);
    var minCA = parseNumber(row[headers.indexOf('minimumChargedAreaM2')]);

    var lengthToUse = Math.max(widthMM,minCW);
    var heightToUse = Math.max(heightMM,minCH);
    var areaToUse = Math.max((lengthToUse/1000)*(heightToUse/1000), minCA);

    var totalCost = (fabric*areaToUse)+(hwW*(lengthToUse/1000))+(hwH*(heightToUse/1000))+(hwU+labor+waste);
    totalCost *= (1+tax);
    totalCost = parseFloat(totalCost.toFixed(2));

    var finalPrice = totalCost * mkUp;
    finalPrice = parseFloat(finalPrice.toFixed(2));
    // discount poderia ser aplicado se necessário, mas no código original não estava sendo aplicado explicitamente.
    // mantemos o mesmo comportamento original (usando finalPrice sem discount adicional).

    return finalPrice;
  });

  return prices;
}

/**
 * Ao invés de calcular diretamente o menor custo, obtém todos os custos possíveis 
 * e depois escolhe conforme a opção (max, min, avg).
 */
function getAllCostsForDimensions(mainSku, widthMM, heightMM, priceTableData) {
  var headers = priceTableData.headers;
  var rows = priceTableData.rows;

  var idxFields = ['fabricCostPerM2','hardwareCostPerWidthML','hardwareCostPerHeightML','hardwareCostPerUnit',
                   'laborCostPerUnit','wasteAndPackagingPerUnit','taxVatOverCost',
                   'minimumWidthMM','maximumWidthMM','minimumHeightMM','maximumHeightMM',
                   'minimumChargedWidthMM','minimumChargedHeightMM','minimumChargedAreaM2','maximumAreaM2','correlationHeightByLength'];

  var idx = {};
  idxFields.forEach(function(field){ idx[field] = headers.indexOf(field); });

  var candidates = rows.filter(function(row){
    var minW = parseNumber(row[idx['minimumWidthMM']]);
    var maxW = parseNumber(row[idx['maximumWidthMM']]);
    var minH = parseNumber(row[idx['minimumHeightMM']]);
    var maxH = parseNumber(row[idx['maximumHeightMM']]);
    var maxA = parseNumber(row[idx['maximumAreaM2']]);
    var corr = parseNumber(row[idx['correlationHeightByLength']]);

    var realArea = (widthMM/1000)*(heightMM/1000);
    var realHByL = heightMM/widthMM;

    return (widthMM>=minW && widthMM<=maxW && heightMM>=minH && heightMM<=maxH && realArea<=maxA && realHByL<=corr);
  });

  if (candidates.length==0) return [19999.00];

  var costs = candidates.map(function(row){
    var fabric = parseNumber(row[idx['fabricCostPerM2']]);
    var hwW = parseNumber(row[idx['hardwareCostPerWidthML']]);
    var hwH = parseNumber(row[idx['hardwareCostPerHeightML']]);
    var hwU = parseNumber(row[idx['hardwareCostPerUnit']]);
    var labor = parseNumber(row[idx['laborCostPerUnit']]);
    var waste = parseNumber(row[idx['wasteAndPackagingPerUnit']]);
    var tax = normalizePercentage(row[idx['taxVatOverCost']]);

    var minCW = parseNumber(row[headers.indexOf('minimumChargedWidthMM')]);
    var minCH = parseNumber(row[headers.indexOf('minimumChargedHeightMM')]);
    var minCA = parseNumber(row[headers.indexOf('minimumChargedAreaM2')]);

    var lengthToUse = Math.max(widthMM,minCW);
    var heightToUse = Math.max(heightMM,minCH);
    var areaToUse = Math.max((lengthToUse/1000)*(heightToUse/1000), minCA);

    var totalCost = (fabric*areaToUse)+(hwW*(lengthToUse/1000))+(hwH*(heightToUse/1000))+(hwU+labor+waste);
    totalCost *= (1+tax);
    totalCost = parseFloat(totalCost.toFixed(2));

    return totalCost;
  });

  return costs;
}

/**
 * Escolhe o preço final (max, min ou avg) conforme as booleans definidas.
 */
function choosePrice(prices) {
  if (prices.length == 0) return 19999.00;
  if (useMinPrice) {
    return Math.min.apply(null, prices);
  } else if (useMaxPrice) {
    return Math.max.apply(null, prices);
  } else if (useAvgPrice) {
    var sum = prices.reduce(function(a,b){return a+b;},0);
    return sum / prices.length;
  }
  // Se nenhum estiver true, fallback min
  return Math.min.apply(null, prices);
}

/**
 * Escolhe o custo final (max, min ou avg) conforme as booleans definidas.
 */
function chooseCost(costs) {
  if (costs.length == 0) return 19999.00;
  if (useMinCost) {
    return Math.min.apply(null, costs);
  } else if (useMaxCost) {
    return Math.max.apply(null, costs);
  } else if (useAvgCost) {
    var sum = costs.reduce(function(a,b){return a+b;},0);
    return sum / costs.length;
  }
  // fallback min
  return Math.min.apply(null, costs);
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    value = value.replace(',', '.');
  }
  var number = parseFloat(value);
  if (isNaN(number)) number = 0;
  return number;
}

function normalizePercentage(value) {
  var number = parseNumber(value);
  if (number > 1) {
    number = number / 100;
  }
  return number;
}

/** Funções para atualizar a aba updatedVariantsPriceTable */
function button0_updatePriceTableSheet_V3() {
  if (typeof DEBUG === 'undefined') { var DEBUG = true; }
  if (DEBUG) {Logger.log('Iniciando o processo de atualização da aba updatedVariantsPriceTable'); console.log('Iniciando o processo de atualização da aba updatedVariantsPriceTable');}

  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDestino = ss.getSheetByName('updatedVariantsPriceTable');
  var logs = [];
  logs.push('Iniciando o processo de atualização.');

  if (!sheetDestino) {
    ui.alert('A aba "updatedVariantsPriceTable" não foi encontrada.');
    logs.push('A aba de destino não foi encontrada e o processo foi interrompido.');
    if (DEBUG) {Logger.log('A aba de destino não foi encontrada e o processo foi interrompido.'); console.log('A aba de destino não foi encontrada e o processo foi interrompido.');}
    return;
  }
  logs.push('Aba de destino encontrada: updatedVariantsPriceTable');
  if (DEBUG) {Logger.log('Aba de destino encontrada: updatedVariantsPriceTable'); console.log('Aba de destino encontrada: updatedVariantsPriceTable');}

  var responseFacil = ui.alert(
    'Atualização de Preços',
    'Você deseja atualizar os preços para a FACIL PERSIANAS?',
    ui.ButtonSet.YES_NO
  );
  logs.push('Resposta sobre FACIL PERSIANAS: ' + (responseFacil == ui.Button.YES ? 'Sim' : 'Não'));
  if (DEBUG) {Logger.log('Resposta sobre FACIL PERSIANAS: ' + (responseFacil == ui.Button.YES ? 'Sim' : 'Não')); console.log('Resposta sobre FACIL PERSIANAS: ' + (responseFacil == ui.Button.YES ? 'Sim' : 'Não'));}

  var DEFAULT_SPREADSHEET_ID = '13m0q4ryo-LACgxY4XRqVz9mOs7-0x3Eq3pE-KXRYBjE';
  var sourceSpreadsheetId = DEFAULT_SPREADSHEET_ID;
  var sourceSheetName = '';

  if (responseFacil == ui.Button.YES) {
    sourceSheetName = 'Facil Persianas';
    if (DEBUG) {Logger.log('Usuário selecionou FACIL PERSIANAS.'); console.log('Usuário selecionou FACIL PERSIANAS.');}
  } else {
    var responseP2GO = ui.alert(
      'Atualização de Preços',
      'Você gostaria de atualizar o preço da P2GO usando a tabela da P2GO?',
      ui.ButtonSet.YES_NO
    );
    logs.push('Resposta sobre P2GO: ' + (responseP2GO == ui.Button.YES ? 'Sim' : 'Não'));
    if (DEBUG) {Logger.log('Resposta sobre P2GO: ' + (responseP2GO == ui.Button.YES ? 'Sim' : 'Não')); console.log('Resposta sobre P2GO: ' + (responseP2GO == ui.Button.YES ? 'Sim' : 'Não'));}

    if (responseP2GO == ui.Button.YES) {
      sourceSheetName = 'P2GO';
      if (DEBUG) {Logger.log('Usuário selecionou P2GO.'); console.log('Usuário selecionou P2GO.');}
    } else {
      var sheetResponse = ui.prompt('Insira o nome da aba que deseja usar:', ui.ButtonSet.OK_CANCEL);
      if (sheetResponse.getSelectedButton() != ui.Button.OK) {
        ui.alert('Ação cancelada.');
        logs.push('Usuário cancelou a ação ao inserir o nome da aba.');
        if (DEBUG) {Logger.log('Usuário cancelou a ação ao inserir o nome da aba.'); console.log('Usuário cancelou a ação ao inserir o nome da aba.');}
        return;
      }
      sourceSheetName = sheetResponse.getResponseText().trim();
      if (!sourceSheetName) {
        ui.alert('Nome da aba não fornecido. Ação cancelada.');
        logs.push('Nome da aba não fornecido. Processo interrompido.');
        if (DEBUG) {Logger.log('Nome da aba não fornecido. Processo interrompido.'); console.log('Nome da aba não fornecido. Processo interrompido.');}
        return;
      }
      logs.push('Nome da aba fornecido: ' + sourceSheetName);
      if (DEBUG) {Logger.log('Nome da aba fornecido: ' + sourceSheetName); console.log('Nome da aba fornecido: ' + sourceSheetName);}

      var responseDefaultSheet = ui.alert(
        'Confirmação de Planilha',
        'A aba "' + sourceSheetName + '" está na planilha padrão?',
        ui.ButtonSet.YES_NO
      );
      logs.push('A aba está na planilha padrão? ' + (responseDefaultSheet == ui.Button.YES ? 'Sim' : 'Não'));
      if (DEBUG) {Logger.log('A aba está na planilha padrão? ' + (responseDefaultSheet == ui.Button.YES ? 'Sim' : 'Não')); console.log('A aba está na planilha padrão? ' + (responseDefaultSheet == ui.Button.YES ? 'Sim' : 'Não'));}

      if (responseDefaultSheet == ui.Button.NO) {
        var docResponse = ui.prompt(
          'Insira o ID do documento de origem:',
          'Exemplo de ID: 1abcD_EfgHiJkLmNoPqRsTuVwXyZ',
          ui.ButtonSet.OK_CANCEL
        );
        if (docResponse.getSelectedButton() != ui.Button.OK) {
          ui.alert('Ação cancelada.');
          logs.push('Usuário cancelou a ação ao inserir o ID do documento.');
          if (DEBUG) {Logger.log('Usuário cancelou a ação ao inserir o ID do documento.'); console.log('Usuário cancelou a ação ao inserir o ID do documento.');}
          return;
        }
        sourceSpreadsheetId = docResponse.getResponseText().trim();
        if (!sourceSpreadsheetId) {
          ui.alert('ID do documento não fornecido. Ação cancelada.');
          logs.push('ID do documento não fornecido. Processo interrompido.');
          if (DEBUG) {Logger.log('ID do documento não fornecido. Processo interrompido.'); console.log('ID do documento não fornecido. Processo interrompido.');}
          return;
        }
        logs.push('ID do documento fornecido: ' + sourceSpreadsheetId);
        if (DEBUG) {Logger.log('ID do documento fornecido: ' + sourceSpreadsheetId); console.log('ID do documento fornecido: ' + sourceSpreadsheetId);}
      } else {
        sourceSpreadsheetId = DEFAULT_SPREADSHEET_ID;
        logs.push('Usando a planilha padrão.');
        if (DEBUG) {Logger.log('Usando a planilha padrão.'); console.log('Usando a planilha padrão.');}
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Processando os dados...', 'Atualização em andamento', 5);
  logs.push('Iniciando a atualização com os parâmetros definidos.');
  if (DEBUG) {Logger.log('Iniciando a atualização com os parâmetros definidos.'); console.log('Iniciando a atualização com os parâmetros definidos.');}

  atualizarPlanilhaComFormatacao(
    sourceSpreadsheetId,
    sourceSheetName,
    sheetDestino,
    logs
  );
}

function atualizarPlanilhaComFormatacao(
  sourceSpreadsheetId,
  sourceSheetName,
  sheetDestino,
  logs
) {
  var ui = SpreadsheetApp.getUi();
  if (typeof DEBUG === 'undefined') { var DEBUG = true; }
  if (DEBUG) {Logger.log('Obtida interface do usuário na função de atualização.'); console.log('Obtida interface do usuário na função de atualização.');}
  try {
    logs.push('Iniciando atualização da planilha.');
    if (DEBUG) {Logger.log('Iniciando atualização da planilha.'); console.log('Iniciando atualização da planilha.');}

    var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    logs.push('Documento de origem acessado com sucesso: ' + sourceSpreadsheetId);
    if (DEBUG) {Logger.log('Documento de origem acessado com sucesso: ' + sourceSpreadsheetId); console.log('Documento de origem acessado com sucesso: ' + sourceSpreadsheetId);}

    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      ui.alert('A aba "' + sourceSheetName + '" não foi encontrada no documento de origem.');
      logs.push('A aba de origem não foi encontrada. Processo interrompido.');
      if (DEBUG) {Logger.log('A aba de origem não foi encontrada. Processo interrompido.'); console.log('Aba de origem não encontrada.');}
      return;
    }
    logs.push('Aba de origem encontrada: ' + sourceSheetName);
    if (DEBUG) {Logger.log('Aba de origem encontrada: ' + sourceSheetName); console.log('Aba de origem encontrada: ' + sourceSheetName);}

    var COLUMN_START = 'A';
    var COLUMN_END = 'E';
    var lastRow = findLastRowWithData(sourceSheet, COLUMN_START, COLUMN_END);
    logs.push('Última linha preenchida encontrada: ' + lastRow);
    if (DEBUG) {Logger.log('Última linha preenchida encontrada: ' + lastRow); console.log('Última linha preenchida: ', lastRow);}

    if (lastRow < 2) {
      ui.alert('A aba de origem não contém dados suficientes para copiar.');
      logs.push('Dados insuficientes na aba de origem. Processo interrompido.');
      if (DEBUG) {Logger.log('Dados insuficientes na aba de origem. Processo interrompido.'); console.log('Dados insuficientes na aba de origem.');}
      return;
    }

    var lastColumn = sourceSheet.getLastColumn();
    var lastColumnLetter = columnToLetter(lastColumn);
    var sourceRangeNotation = 'A1:' + lastColumnLetter + lastRow;
    logs.push('Intervalo a ser copiado: ' + sourceRangeNotation);
    if (DEBUG) {Logger.log('Intervalo a ser copiado: ' + sourceRangeNotation); console.log('Intervalo a ser copiado: ', sourceRangeNotation);}

    var cellA1 = sheetDestino.getRange('A1');
    var formulaA1 = cellA1.getFormula();
    var valueA1 = cellA1.getValue();
    logs.push('Conteúdo da célula A1 armazenado.');
    if (DEBUG) {Logger.log('Conteúdo da célula A1 armazenado.'); console.log('Conteúdo da célula A1 armazenado.');}

    sheetDestino.clearContents();
    sheetDestino.clearFormats();
    logs.push('Aba de destino limpa.');
    if (DEBUG) {Logger.log('Aba de destino limpa.'); console.log('Aba de destino limpa.');}

    if (formulaA1) {
      cellA1.setFormula(formulaA1);
    } else {
      cellA1.setValue(valueA1);
    }
    logs.push('Conteúdo da célula A1 restaurado.');
    if (DEBUG) {Logger.log('Conteúdo da célula A1 restaurado.'); console.log('Conteúdo da célula A1 restaurado.');}

    var sourceRange = sourceSheet.getRange(sourceRangeNotation);
    var sourceValues = sourceRange.getValues();
    var sourceBackgrounds = sourceRange.getBackgrounds();
    logs.push('Dados e formatações obtidos da aba de origem.');
    if (DEBUG) {Logger.log('Dados e formatações obtidos da aba de origem.'); console.log('Dados e formatações obtidos da aba de origem.');}

    var destinoRange = sheetDestino.getRange('A1').offset(0, 0, sourceValues.length, sourceValues[0].length);
    destinoRange.setValues(sourceValues);
    logs.push('Dados aplicados na aba de destino.');
    if (DEBUG) {Logger.log('Dados aplicados na aba de destino.'); console.log('Dados aplicados na aba de destino.');}

    destinoRange.setBackgrounds(sourceBackgrounds);
    logs.push('Formatações aplicadas na aba de destino.');
    if (DEBUG) {Logger.log('Formatações aplicadas na aba de destino.'); console.log('Formatações aplicadas na aba de destino.');}

    sheetDestino.autoResizeColumns(1, sourceValues[0].length);
    logs.push('Largura das colunas ajustada.');
    if (DEBUG) {Logger.log('Largura das colunas ajustada.'); console.log('Largura das colunas ajustada.');}

    SpreadsheetApp.getActiveSpreadsheet().toast('A aba "updatedVariantsPriceTable" foi atualizada com sucesso!', 'Atualização Concluída', 5);
    logs.push('Processo concluído com sucesso.');
    if (DEBUG) {Logger.log('Processo concluído com sucesso.'); console.log('Processo concluído com sucesso.');}
  } catch (error) {
    ui.alert('Ocorreu um erro durante a atualização: ' + error);
    logs.push('Erro: ' + error);
    if (DEBUG) {Logger.log('Erro: ' + error); console.log('Erro: ', error);}
  } finally {
    logs.forEach(function(log) {
      if (DEBUG) {console.log(log); Logger.log(log);}
    });
  }
}

function findLastRowWithData(sheet, startCol, endCol) {
  if (typeof DEBUG === 'undefined') { var DEBUG = true; }
  var lastRow = sheet.getLastRow();
  if (DEBUG) {Logger.log('Última linha da planilha: ' + lastRow); console.log('Última linha da planilha: ', lastRow);}
  var dataRange = sheet.getRange(startCol + '1:' + endCol + lastRow);
  if (DEBUG) {Logger.log('Intervalo de dados definido: ' + startCol + '1:' + endCol + lastRow); console.log('Intervalo de dados definido: ', startCol + '1:' + endCol + lastRow);}
  var data = dataRange.getValues();
  if (DEBUG) {Logger.log('Valores do intervalo obtidos.'); console.log('Valores do intervalo obtidos.');}

  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    var isRowFilled = row.every(function(cell) {
      return cell !== '';
    });
    if (isRowFilled) {
      if (DEBUG) {Logger.log('Última linha preenchida encontrada na linha: ' + (i + 1)); console.log('Última linha preenchida encontrada na linha: ', (i+1));}
      return i + 1;
    }
  }
  if (DEBUG) {Logger.log('Nenhuma linha preenchida encontrada.'); console.log('Nenhuma linha preenchida encontrada.');}
  return 0;
}

function columnToLetter(column) {
  if (typeof DEBUG === 'undefined') { var DEBUG = true; }
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  if (DEBUG) {Logger.log('Número da coluna convertido para letra: ' + letter); console.log('Número da coluna convertido para letra: ', letter);}
  return letter;
}

