const { createWorkbook, createWorksheetTemplate } = require('./createTemplate');
const pbcJsonParser = require('./jsonParser');

// Fill the workbook main function
const fillWorkbook = async (workbook, worksheet) => {
  const parsedData = pbcJsonParser();
  for (var index = 0; index < parsedData.length; index++) {
    transactionSort(parsedData, index);
    const propertiesArray = getPropertiesArray(parsedData, index);
    worksheet.eachRow((row, rowNumber) => {
      if (monthAndYearCompare(row, parsedData, index)) {
        fillMonth0(row, propertiesArray);
        const acquisitionCost = row.getCell('acquisitionCost').value;
        const totalPaymentsAndPercentage = fillPayments(row, propertiesArray, acquisitionCost);
        fillRemainingCells(row, totalPaymentsAndPercentage);
      }
    })
  }
  return await workbook.xlsx.writeFile('properties_by_cohort.xlsx');
}

// Sort the properties' transaction in ascending date order
const transactionSort = (parsedData, index) => {
  parsedData.at(index).properties.sort((a, b) => {
    return new Date(a['Payments.transactionDate']) - new Date(b['Payments.transactionDate']);
  })
}

// Fetch the properties array into a more readable object name
const getPropertiesArray = (parsedData, index) => {
  return parsedData.at(index).properties;
}

// Compare rows' month and year and parsed datas' month and year
// If they match then move on to iterate the cell in each column of the row
const monthAndYearCompare = (row, parsedData, index) => {
  if (row.getCell(1).value === null || row.getCell(1).value === 'Cohort') {
    return false;
  } else if (row.getCell(1).value.getMonth() + 1 === parsedData.at(index).month && (row.getCell(1).value.getYear() + 1900) === parsedData.at(index).year) {
    console.log((row.getCell(1).value.getMonth() + 1) + ' - ' + row.getCell(1).value.getYear());
    return true;
  } else {
    return false;
  }
}

// Fill the cell in the month 0 column (#, %, and IDR)
const fillMonth0 = (row, propertiesArray) => {
  let acquisitionCost = propertiesArray.at(0).acquisitionCost;
  let id = [propertiesArray.at(0).id];
  for (var propertiesIndex = 0; propertiesIndex < propertiesArray.length; propertiesIndex++) {
    let check = false;
    for (var idIndex = 0; idIndex < id.length; idIndex++) {
      if (id.at(idIndex) === propertiesArray.at(propertiesIndex).id) {
        check = true;
        break;
      }
    }
    if (check === false) {
      id.push(propertiesArray.at(propertiesIndex).id);
      acquisitionCost += propertiesArray.at(propertiesIndex).acquisitionCost;
    }
  }
  row.getCell('amount').value = id.length;
  row.getCell('original').value = 100;
  row.getCell('acquisitionCost').value = acquisitionCost;
}

// Fill the rest of the months columns
const fillPayments = (row, propertiesArray, acquisitionCost) => {
  const acquisitionDate = new Date(propertiesArray.at(0).earliestAcquisitionDate);
  var totalAmount = 0;
  var percentage = 0;
  for (var propertiesIndex = 0; propertiesIndex < propertiesArray.length; propertiesIndex++) {
    const property = propertiesArray.at(propertiesIndex);
    const date = new Date(property['Payments.transactionDate']);
    if (monthDiff(acquisitionDate, date) === 0 || monthDiff(acquisitionDate, date) === 1) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase1', 'rupiah1', totalAmount, percentage);
      continue;
    } else if (row.getCell('rupiah1').value === null) {
      fillCell(row, 'increase1', 'rupiah1', totalAmount, percentage);
    } 
    if (monthDiff(acquisitionDate, date) === 2) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase2', 'rupiah2', totalAmount, percentage);
      continue;
    } else if (row.getCell('rupiah2').value === null) {
      if (monthDiff(acquisitionDate, date) < 2) {
        continue;
      }
      fillCell(row, 'increase2', 'rupiah2', totalAmount, percentage);
    } 
    if (monthDiff(acquisitionDate, date) === 3) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase3', 'rupiah3', totalAmount, percentage);
      continue;
    } else if (row.getCell('rupiah3').value === null) {
      if (monthDiff(acquisitionDate, date) < 3) {
        continue;
      }
      fillCell(row, 'increase3', 'rupiah3', totalAmount, percentage);
    }
    if (monthDiff(acquisitionDate, date) === 4) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase4', 'rupiah4', totalAmount, percentage);
      continue;
    } else if (row.getCell('rupiah4').value === null) {
      if (monthDiff(acquisitionDate, date) < 4) {
        continue;
      }
      fillCell(row, 'increase4', 'rupiah4', totalAmount, percentage);
    }
    if (monthDiff(acquisitionDate, date) === 5) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase5', 'rupiah5', totalAmount, percentage);
      continue;
    } else if (row.getCell('rupiah5').value === null) {
      if (monthDiff(acquisitionDate, date) < 5) {
        continue;
      }
      fillCell(row, 'increase5', 'rupiah5', totalAmount, percentage);
    }
    if (monthDiff(acquisitionDate, date) === 6) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase6', 'rupiah6', totalAmount, percentage);
      continue;
    } else if (row.getCell('rupiah6').value === null) {
      if (monthDiff(acquisitionDate, date) < 6) {
        continue;
      }
      fillCell(row, 'increase6', 'rupiah6', totalAmount, percentage);
    }
    if (monthDiff(acquisitionDate, date) === 7) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase7', 'rupiah7', totalAmount, percentage);
      continue;
    } else if (row.getCell('rupiah7').value === null) {
      if (monthDiff(acquisitionDate, date) < 7) {
        continue;
      }
      fillCell(row, 'increase7', 'rupiah7', totalAmount, percentage);
    }
    if (monthDiff(acquisitionDate, date) === 8) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase8', 'rupiah8', totalAmount, percentage);
      continue;
    } else if (row.getCell('rupiah8').value === null) {
      if (monthDiff(acquisitionDate, date) < 8) {
        continue;
      }
      fillCell(row, 'increase8', 'rupiah8', totalAmount, percentage);
    }
    if (monthDiff(acquisitionDate, date) === 9) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase9', 'rupiah9', totalAmount, percentage);
      continue;
    } else if (row.getCell('rupiah9').value === null) {
      if (monthDiff(acquisitionDate, date) < 9) {
        continue;
      }
      fillCell(row, 'increase9', 'rupiah9', totalAmount, percentage);
    }
    if (monthDiff(acquisitionDate, date) > 9) {
      totalAmount += property['Payments.originalAmount'];
      percentage = (totalAmount / acquisitionCost) * 100;
      fillCell(row, 'increase9plus', 'rupiah9plus', totalAmount, percentage);
    }
  }
  // return an object for later usage
  return { totalAmount, percentage };
}

// Fill the cell base on the key called inside the argument
const fillCell = (row, keyIncrease, keyRupiah, totalAmount, percentage) => {
  row.getCell(keyIncrease).value = percentage;
  row.getCell(keyRupiah).value = totalAmount
}

// Fill remaining empty cells that are necessary to be input (to create a diagonal shape)
const fillRemainingCells = (row, totalPaymentsAndPercentage) => {
  const lastDate = new Date('2022-07-05');
  const date = new Date(row.getCell(1).value);
  const getMonthDiff = monthDiff(date, lastDate);
  if (getMonthDiff > 9) {
    for (var cellIndex = 5; cellIndex <= 24; cellIndex++) {
      fillCellByIndex(row, cellIndex, totalPaymentsAndPercentage);
    }
  } else {
    const lastCell = 22 - ((9 - getMonthDiff) * 2);
    for (var cellIndex = 5; cellIndex <= lastCell; cellIndex++) {
      fillCellByIndex(row, cellIndex, totalPaymentsAndPercentage);
    }
  }
}

// Fill the empty cells base on the cellIndex (odd = % column and even = IDR column)
const fillCellByIndex = (row, cellIndex, totalPaymentsAndPercentage) => {
  if (cellIndex % 2 != 0 && row.getCell(cellIndex).value === null) {
    row.getCell(cellIndex).value = totalPaymentsAndPercentage.percentage;
  } else if (cellIndex % 2 === 0 && row.getCell(cellIndex).value === null) {
    row.getCell(cellIndex).value = totalPaymentsAndPercentage.totalAmount;
  }
}

// Calculate the month difference between the earliest acquisition date and the Payments.transactionDate
const monthDiff = (d1, d2) => {
  var months;
  months = (d2.getFullYear() - d1.getFullYear()) * 12;
  months -= d1.getMonth();
  months += d2.getMonth();
  return months <= 0 ? 0 : months;
}

// Run this code for generation becasuse the createWorksheetTemplate must be awaited
const runWork = async() => {
  const workbook = createWorkbook();
  const worksheet = await createWorksheetTemplate(workbook);
  fillWorkbook(workbook, worksheet);
}

// Run the fillWorkbook.js file
runWork();