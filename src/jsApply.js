function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Simulador de Precificação')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function simulatePricing(sku, tag) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var priceSheet = ss.getSheetByName('Facil Persianas');
  
  var priceHeaders = priceSheet.getRange(2, 1, 1, priceSheet.getLastColumn()).getValues()[0];
  var indices = {
    larguraMin: priceHeaders.indexOf('Largura Minima (mm)'),
    larguraMax: priceHeaders.indexOf('Largura Maxima (mm)'),
    alturaMax: priceHeaders.indexOf('Altura Maxima (mm)'),
    areaMin: priceHeaders.indexOf('Area Minima Cobrada (m2)'),
    areaMax: priceHeaders.indexOf('Area Maxima (m2)'),
    tagIdentificacao: priceHeaders.indexOf('Tag Identificacao'),
    precoM2: priceHeaders.indexOf('Preco final POR: (m2)'),
    precoDe: priceHeaders.indexOf('Preco DE: (m2)')
  };

  var priceData = priceSheet.getDataRange().getValues();
  var length = parseInt(sku.slice(-7, -4)) * 10;
  var height = parseInt(sku.slice(-3)) * 10;
  var area = (length * height) / 1000000;
  var areaAjustada = Math.max(area, 1.4);

  var autoPriceTag = tag.split(',').find(t => t.trim().startsWith("auto-price:"));
  if (!autoPriceTag) {
    return { error: "Tag auto-price não encontrada" };
  }

  var filteredRows = [];
  var filteringSteps = [];

  for (var i = 1; i < priceData.length; i++) {
    var row = priceData[i];
    var rowTags = row[indices.tagIdentificacao].split(',');
    var rowAutoPrice = rowTags.find(t => t.trim().startsWith("auto-price:"));
    
    if (rowAutoPrice && rowAutoPrice.trim() === autoPriceTag.trim()) {
      var stepResult = {
        row: i + 1,
        criteria: {
          tag: true,
          largura: false,
          altura: false,
          area: false
        },
        passed: false
      };

      if (length >= row[indices.larguraMin] && length <= row[indices.larguraMax]) {
        stepResult.criteria.largura = true;
      }
      if (height <= row[indices.alturaMax]) {
        stepResult.criteria.altura = true;
      }
      if (areaAjustada >= row[indices.areaMin] && areaAjustada <= row[indices.areaMax]) {
        stepResult.criteria.area = true;
      }

      if (stepResult.criteria.largura && stepResult.criteria.altura && stepResult.criteria.area) {
        stepResult.passed = true;
        filteredRows.push({
          row: row,
          precoM2: parseFloat(row[indices.precoM2]),
          tagGuiaLateral: rowTags.find(t => t.trim().startsWith("kbox-autoprice2GL:")),
          tagGuiaInferior: rowTags.find(t => t.trim().startsWith("kbox-autoprice1GI:")),
          tagBandoSuperior: rowTags.find(t => t.trim().startsWith("kbox-autoprice1B:"))
        });
      }

      filteringSteps.push(stepResult);
    }
  }

  if (filteredRows.length === 0) {
    return { error: "Nenhuma linha correspondente encontrada", filteringSteps: filteringSteps };
  }

  filteredRows.sort((a, b) => a.precoM2 - b.precoM2);
  var bestMatch = filteredRows[0];

  var precoM2 = bestMatch.precoM2;
  var precoDe = parseFloat(bestMatch.row[indices.precoDe]);
  var precoPeca = (areaAjustada * precoM2).toFixed(2);
  var precoDePeca = (areaAjustada * precoDe).toFixed(2);

  if (area <= 0.4) {
    precoPeca = precoM2.toFixed(2);
    precoDePeca = precoDe.toFixed(2);
  }

  return {
    sku: sku,
    tag: tag,
    length: length,
    height: height,
    area: area.toFixed(2),
    areaAjustada: areaAjustada.toFixed(2),
    precoPeca: precoPeca,
    precoDePeca: precoDePeca,
    matchedRow: bestMatch.row,
    tagGuiaLateral: bestMatch.tagGuiaLateral,
    tagGuiaInferior: bestMatch.tagGuiaInferior,
    tagBandoSuperior: bestMatch.tagBandoSuperior,
    filteringSteps: filteringSteps
  };
}