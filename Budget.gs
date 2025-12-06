// Base Functions //

function onOpen() {
  // Calculate budget on open to ensure up-to-date values
  calculateBudget();
}

function onEdit(e) {
  // When the sheet is edited, update the values
  calculateBudget();
}

function calculateBudget() {
  const sheet = getSheet('Sheet1');

  var currentSavings = updateSavingsCell(sheet);

  const totalSavings = locateAdjacentCell(sheet, "Total Savings", 8, 1, true, currentSavings);
  const spendingBudget = locateAdjacentCell(sheet, "Spending Budget", 11, 1, false);

  const serviceMap = getServiceMap(sheet);
  const totalProjectedCosts = calculatetotalProjectedCosts(serviceMap);

  const spentServiceMap = getSpentsMap(sheet);
  const totalSpentCosts = calculateSpentCosts(spentServiceMap);

  updateSavingsProjections(sheet, currentSavings, totalProjectedCosts, totalSpentCosts);
  updateRealSavingsProjections(sheet, currentSavings, totalProjectedCosts, totalSpentCosts);
  updateOverAndUnderProjections(sheet, currentSavings, totalSpentCosts, spendingBudget);
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////\
// Primary Functions //

function updateSavingsCell(sheet) {
  let checkings   = locateAdjacentCell(sheet, "Checkings", 7, 1, false);
  let savings     = locateAdjacentCell(sheet, "Savings", 8, 1, false);
  return checkings + savings;
}

function updateSavingsProjections(sheet, currentSavings, costs, spent) {
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  let currentColumn = 3;
  let monthIndex = getCurrentMonth();
  let savings = currentSavings
  let endSavings = currentSavings - spent

  for (let i = 0; i < 12; i++) {
    monthIndex = monthIndex % 12;

    sheet.getRange(16, currentColumn).setValue(months[monthIndex]);
    sheet.getRange(17, currentColumn).setValue(savings);
    sheet.getRange(18, currentColumn).setValue(endSavings);
    
    currentColumn++;
    monthIndex++;
    savings -= costs;
    endSavings -= costs;
  }
}

function updateRealSavingsProjections(sheet, currentSavings, costs, spent) {
  locateAdjacentCell(sheet, "Projected Costs", 23, 2, true, costs);
  locateAdjacentCell(sheet, "Total Spent Costs", 24, 2, true, spent);

  let disparity = costs - spent;
  if (disparity < 0) {
    locateAdjacentCell(sheet, "Costs Disparity", 26, 2, true, (disparity * -1));
  }
  else {
    locateAdjacentCell(sheet, "Costs Disparity", 26, 2, true, disparity);
  }
}


function updateOverAndUnderProjections(sheet, currentSavings, spent, spendingBudget) {
  let projection = spendingBudget - spent;
  if (projection < spendingBudget && projection > 0) {
    locateAdjacentCell(sheet, "Over Projection", 29, 2, true, 0);
    locateAdjacentCell(sheet, "Under Projection", 30, 2, true, projection);
  }
  else if (projection < 0) {
    locateAdjacentCell(sheet, "Over Projection", 29, 2, true, (projection * (-1)));
    locateAdjacentCell(sheet, "Over Projection", 30, 2, true, 0);
  }
}

