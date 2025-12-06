//////////////////////////////////////////////////////////////////////////////////////////////////////////////\
// Helper Functions //

function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

function getServiceMap(sheet) {
  const servicesList   = getColumnValuesUntilEmpty(sheet, 1, 2);
  const pricesList     = getColumnValuesUntilEmpty(sheet, 2, 2);
  const datesList      = getColumnValuesUntilEmpty(sheet, 3, 2);

  adjustListLengths(servicesList, pricesList, datesList);

  const serviceMap = new Map();
  for (let i = 0; i < servicesList.length; i++) {
    const service   = servicesList[i];
    const price     = pricesList[i];
    const date      = datesList[i];

    if (!serviceMap.has(service)) {
      serviceMap.set(service, []);
    }

    const serviceDetails = serviceMap.get(service);
    serviceDetails.push({Price: price, Date: date});
  }
  return serviceMap;
}

function getSpentsMap(sheet) {
  const spentServicesList  = getColumnValuesUntilEmpty(sheet, 2, 21);
  const spentPricesList   = getColumnValuesUntilEmpty(sheet, 3, 21);

  adjustListLengths(spentServicesList, spentPricesList);
  
  const spentServiceMap = new Map();
  for (let i = 0; i < spentServicesList.length; i++) {
    const spentService  = spentServicesList[i];
    const spentPrice    = spentPricesList[i];

    if (!spentServiceMap.has(spentService)) {
      spentServiceMap.set(spentService, []);
    }

    const spentServiceDetails = spentServiceMap.get(spentService);
    spentServiceDetails.push({Price: spentPrice});
  }
  return spentServiceMap;
}

function adjustListLengths(servicesList, pricesList, datesList) {
  while (servicesList.length > pricesList.length) {
    pricesList.push(0);
  }
  if (!datesList === null){
    while (servicesList.length > datesList.length) {
      datesList.push(0);
    }
  }
}

function getColumnValuesUntilEmpty(sheet, column, startRow) {
  const values = [];
  let row = startRow;

  while (true) {
    const cellValue = sheet.getRange(row, column).getValue();
    if (cellValue === "") {
      break;
    }
    values.push(cellValue);
    row++;
  }

  return values;
}

function locateCell(sheet, valueToFind, startRow, column, isInput, input) {
  if (isInput === null) {
    isInput = false;
  }

  var lastRow = sheet.getLastRow();
  for (var row = startRow; row <= lastRow; row++) { 
    let cellValue = sheet.getRange(row, column).getValue();
    if (celLValue === valueToFind) {
      if (isInput === true) {
        sheet.getRange(row, column).setValue(input);
        return sheet.getRange(row, column);
      }
      else {
        return sheet.getRange(row, column).getValue();
      }
    }
  }

  return null;
}

function locateAdjacentCell(sheet, valueToFindAdjacent, startRow, column, isFindAdjacent, input) {
  if (isFindAdjacent === null) {
    isFindAdjacent = false;
  }

  var lastRow = sheet.getLastRow();
  for (var row = startRow; row <= lastRow; row++) {
    let cellValue = sheet.getRange(row, column).getValue();
    if (cellValue === valueToFindAdjacent) {
      if (isFindAdjacent === true) {
        sheet.getRange(row, column+1).setValue(input);
        return sheet.getRange(row, column+1);
      }
      else {
        return sheet.getRange(row, column+1).getValue();
      }
    }
  }

  return null;
}

function getFloatFromCell(cell) {
  const cellValue = cell.getValue();
  const floatValue = parseFloat(cellValue);

  if (isNaN(floatValue)) {
    throw new Error('The cell value is not a valid float.');
  }

  return floatValue;
}

function getCurrentMonth() {
  return new Date().getMonth();
}

function calculatetotalProjectedCosts(serviceMap) {
  let cc = 0;
  for (let service of serviceMap.values()) {
    for (let detail of service) {
      cc += parseFloat(detail.Price);
    }
  }

  return cc;
}

function calculateSpentCosts(spentServiceMap) {
  let totalSpentCosts = 0;
  for (let services of spentServiceMap.values()) {
    for (let detail of services) {
      totalSpentCosts += detail.Price; // Prices are already floats
    }
  }
  return totalSpentCosts;
}
