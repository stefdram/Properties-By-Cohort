// Template creation file

const exceljs = require('exceljs');

// Create workbook
const createWorkbook = () => {
  const workbook = new exceljs.Workbook();
  return workbook;
}

// Worksheet template creation
const createWorksheetTemplate = async (workbook, name) => {
  const worksheet = workbook.addWorksheet(name, {
		pageSetup: { fitToPage: true, fitToHeight: 5, fitToWidth: 7 }
	});

  worksheet.columns = [
    { header: 'Cohort', key: 'cohort', width: 10 },
    { header: '#', key: 'amount', width: 7 },
    { header: '%', key: 'original', width: 7 },
    { header: 'IDR', key: 'acquisitionCost', width: 15 },
    { header: '%', key: 'increase1', width: 7 },
    { header: 'IDR', key: 'rupiah1', width: 15 },
    { header: '%', key: 'increase2', width: 7 },
    { header: 'IDR', key: 'rupiah2', width: 15 },
    { header: '%', key: 'increase3', width: 7 },
    { header: 'IDR', key: 'rupiah3', width: 15 },
    { header: '%', key: 'increase4', width: 7 },
    { header: 'IDR', key: 'rupiah4', width: 15 },
    { header: '%', key: 'increase5', width: 7 },
    { header: 'IDR', key: 'rupiah5', width: 15 },
    { header: '%', key: 'increase6', width: 7 },
    { header: 'IDR', key: 'rupiah6', width: 15 },
    { header: '%', key: 'increase7', width: 7 },
    { header: 'IDR', key: 'rupiah7', width: 15 },
    { header: '%', key: 'increase8', width: 7 },
    { header: 'IDR', key: 'rupiah8', width: 15 },
    { header: '%', key: 'increase9', width: 7 },
    { header: 'IDR', key: 'rupiah9', width: 15 },
    { header: '%', key: 'increase9plus', width: 7 },
    { header: 'IDR', key: 'rupiah9plus', width: 15 },
  ];

  for (var d = new Date('2019-10-01'); d <= new Date('2022-08-01'); d.setMonth(d.getMonth() + 1)) {
    worksheet.addRow({ cohort: new Date(d) });
  }

  worksheet.getColumn('cohort').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = 'mmm-yy';
  });

  currencyFormatter(worksheet);

  addMonthsHeader(worksheet);

  worksheet.eachRow({ includeEmpty: true }, formatCell);

  await workbook.xlsx.writeFile('properties_by_cohort.xlsx');

  return worksheet;
}

// Format the IDR columns to be currency numbers and the % columns to have 2 decimal places
const currencyFormatter = (worksheet) => {
  worksheet.getColumn('acquisitionCost').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '#,##;[Red]\-#,##';
  });
  worksheet.getColumn('increase1').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah1').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
  worksheet.getColumn('increase2').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah2').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
  worksheet.getColumn('increase3').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah3').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
  worksheet.getColumn('increase4').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah4').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
  worksheet.getColumn('increase5').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah5').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
  worksheet.getColumn('increase6').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah6').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
  worksheet.getColumn('increase7').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah7').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
  worksheet.getColumn('increase8').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah8').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
  worksheet.getColumn('increase9').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah9').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
  worksheet.getColumn('increase9plus').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0.00'
  })
  worksheet.getColumn('rupiah9plus').eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    cell.numFmt = '0,##';
  });
}

// Add the months header on top
const addMonthsHeader = (worksheet) => {
  const insertRow = [];
  worksheet.insertRow(1, insertRow);

  worksheet.getCell('B1').value = 'Month 0'
  worksheet.getCell('E1').value = 'Month 1';
  worksheet.getCell('G1').value = 'Month 2';
  worksheet.getCell('I1').value = 'Month 3';
  worksheet.getCell('K1').value = 'Month 4';
  worksheet.getCell('M1').value = 'Month 5';
  worksheet.getCell('O1').value = 'Month 6';
  worksheet.getCell('Q1').value = 'Month 7';
  worksheet.getCell('S1').value = 'Month 8';
  worksheet.getCell('U1').value = 'Month 9';
  worksheet.getCell('W1').value = 'Month 9+';

  worksheet.mergeCells('B1:D1');
  worksheet.mergeCells('E1:F1');
  worksheet.mergeCells('G1:H1');
  worksheet.mergeCells('I1:J1');
  worksheet.mergeCells('K1:L1');
  worksheet.mergeCells('M1:N1');
  worksheet.mergeCells('O1:P1');
  worksheet.mergeCells('Q1:R1');
  worksheet.mergeCells('S1:T1');
  worksheet.mergeCells('U1:V1');
  worksheet.mergeCells('W1:X1');
}

// Formatting fonts and alignments base on the given template
const formatCell = (row, rowNumber) => {
  row.font = { size: 12 };
  row.alignment = {
    vertical: 'middle',
    horizontal: 'center',
    wrapText: true,
  };
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    if (colNumber === 1) {
      if (rowNumber === 2) {
        cell.font = { bold: true, size: 12 };
      } else {
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'right',
          wrapText: true,
        };
      }  
    } else {
      if (rowNumber === 1) {
        cell.font = { bold: true, size: 12 };
      }
    }
  });
};

module.exports = {createWorkbook, createWorksheetTemplate};