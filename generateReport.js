const Excel = require('exceljs')
const fs = require('fs');

import _ from 'lodash'
import { db } from './database';

const logoPath = './vendor/logo.png';

export const generateReport = (batch, outputDirectory) => {
  return new Promise((resolve, reject) => {
    generateReportTemplate(batch, outputDirectory).then((reportFilePath) => {
      fillData(batch, 'All').then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'B');
      }).then(() => {
        return fillData(batch, 'ByBatch');
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'C');
      }).then(() => {
        resolve(reportFilePath);
      })
    })
  })
}

const generateReportTemplate = (batch, outputDirectory) => {
  return new Promise((resolve, reject) => {
    let dir = outputDirectory + '/' + batch;

    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir)
    }

    let reportFilePath = dir + '/' + batch + '_report.xlsx';

    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet('Abs', {});

    worksheet.getColumn('A').width = 85;
    worksheet.getRow(1).height = 50;

    // Add Logo
    let logo = workbook.addImage({
      filename: logoPath,
      extension: 'png'
    });

    worksheet.addImage(logo, {
      tl: { col: 0, row: 0 },
      br: { col: 1, row: 1 }
    });

    worksheet.getColumn('B').width = 25;
    worksheet.getColumn('C').width = 25;
    // A1

    worksheet.getCell('B1').value = 'OPPO PROJECT';

    worksheet.getCell('B1').font = {
      bold: true, size: 27, name: 'Calibri', family: 2,
      color: { argb: 'FFFF0000' }
    }

    worksheet.getCell('B1').alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('B1:D1')

    // A2
    worksheet.getCell('B2').font = {
      bold: true, size: 14, name: 'Calibri', family: 2,
      underline: true,
      color: { argb: 'FFFF0000' }
    }

    worksheet.getCell('B2').alignment = { horizontal: 'center', vertical: 'middle' };

    worksheet.getCell('B2').value = 'Step 1: Database Clean';

    // A4
    worksheet.getCell('A4').border = {
      left: { style: 'thin' },
      right: { style: 'thin' },
      top: { style: 'thin' },
      bottom: { style: 'thin' }
    }

    worksheet.getCell('A4').font = {
      bold: true, size: 14, name: 'Calibri', family: 2
    }

    worksheet.getCell('A4').alignment = { horizontal: 'center', vertical: 'middle' };

    worksheet.getCell('A4').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFABF8F' },
      bgColor: { indexed: 64 }
    };

    worksheet.getCell('A4').value = batch;

    // A6, A27
    buildReportFirstColumnType3(worksheet, 5, 'Raw data received from Agency');
    buildReportFirstColumnType3(worksheet, 13, 'Valid database (value) - base all');

    // A7, A14, A21
    buildReportFirstColumnType2(worksheet, 6, 'Phone number missing (column C=blank)');
    buildReportFirstColumnType2(worksheet, 7, 'Duplicated Phone (Checking vs. total database all projects, column C)');
    buildReportFirstColumnType2(worksheet, 10, 'Illogical Phone (column C)');

    // A8 - A13, A15 - A20, A22-A25
    buildReportFirstColumnType1(worksheet, 8, "Duplicated within same model/ product (same data in column F)");
    buildReportFirstColumnType1(worksheet, 9, "Duplicated with previous model/ product (different data in column F)");
    buildReportFirstColumnType1(worksheet, 11, "Illogical phone number format (smaller/ higher than 10 digits)");
    buildReportFirstColumnType1(worksheet, 12, "Illogical phone providers (10 digits but not 03x, 05x, 07x, 08x, 09x)");
    // Done 1st Col

    // Row 5, B4-T4
    buildReportRow5(worksheet, 'B', 'Total Project');
    buildReportRow5(worksheet, 'C', 'Total ' + batch);

    // Data
    let colArr = ['B', 'C'];
    let rowArr = [5, 6, 7, 8, 9, 10, 11, 12, 13];

    for (let rowArrIndex = 0; rowArrIndex < rowArr.length; rowArrIndex += 1) {
      for (let colArrIndex = 0; colArrIndex < colArr.length; colArrIndex += 1 ) {
        buildDataRow(worksheet, rowArr[rowArrIndex], colArr[colArrIndex]);
      }
    }
    // End Data

    // Write to File
    workbook.xlsx.writeFile(reportFilePath).then((res) => {
      resolve(reportFilePath);
    });
  });
}

function buildReportRow5(worksheet, cellIndex, text) {
  let row = worksheet.getRow(4);
  let fgColor = { theme: 0, tint: -0.1499984740745262 };

  if (cellIndex == 'B') {
    fgColor = { theme: 2, tint: -0.249977111117893 };
  }

  if (cellIndex == 'C') {
    fgColor = { theme: 5, tint: 0.5999938962981048 };
  }

  if (cellIndex == 'D' || cellIndex == 'E' || cellIndex == 'F' || cellIndex == 'G' ||
    cellIndex == 'H' || cellIndex == 'I'
  ) {
    fgColor = { argb: 'FFFFFF00' };
  }

  if (cellIndex == 'J' || cellIndex == 'K' || cellIndex == 'L') {
    fgColor = { theme: 6, tint: 0.3999755851924192 };
  }

  row.getCell(cellIndex).border = {
    left: { style: 'thin' },
    right: { style: 'thin' },
    top: { style: 'thin' },
    bottom: { style: 'thin' }
  }

  row.getCell(cellIndex).font = {
    bold: true, size: 14, name: 'Calibri', family: 2
  }

  row.getCell(cellIndex).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: fgColor,
    bgColor: { indexed: 64 }
  };

  row.getCell(cellIndex).alignment = { horizontal: 'center', vertical: 'middle' };

  row.getCell(cellIndex).value = text;
}

function buildReportRow4(worksheet, cellIndex, mergeRange, text) {
  let row = worksheet.getRow(4);

  row.getCell(cellIndex).border = {
    left: { style: 'thin' },
    right: { style: 'thin' },
    top: { style: 'thin' },
    bottom: { style: 'thin' }
  }

  let fgColor = '';

  switch (cellIndex) {
    case 'D':
      fgColor = { argb: 'FFFFFF00' };
      break;
    case 'J':
      fgColor = { theme: 6, tint: 0.3999755851924192 };
      break;
  }

  row.getCell(cellIndex).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: fgColor,
    bgColor: { indexed: 64 }
  }

  row.getCell(cellIndex).font = {
    bold: true,
    size: 12,
    color: { argb: 'FF0070C0' },
    name: 'Calibri',
    family: 2
  }

  row.getCell(cellIndex).alignment = { horizontal: 'center', vertical: 'middle' };

  row.getCell(cellIndex).value = text;

  worksheet.mergeCells(mergeRange);
}

function buildReportFirstColumnType3(worksheet, rowIndex, text) {
  let row = worksheet.getRow(rowIndex);

  row.getCell('A').border = {
    left: { style: 'thin' },
    right: { style: 'thin' },
    top: { style: 'thin' },
    bottom: { style: 'thin' }
  }

  if (rowIndex == 20) {
    row.getCell('A').font = {
      bold: true, size: 14, name: 'Calibri', family: 2,
      color: { theme: 0 }
    }
  } else {
    row.getCell('A').font = {
      bold: true, size: 14, name: 'Calibri', family: 2,
      color: { argb: 'FFFF0000' }
    }
  }

  row.getCell('A').alignment = { vertical: 'middle' };

  if (rowIndex == 20) {
    row.getCell('A').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF0000' },
      bgColor: { indexed: 64 }
    };
  }

  row.getCell('A').value = text;
}


function buildReportFirstColumnType2(worksheet, rowIndex, text) {
  let row = worksheet.getRow(rowIndex);

  row.getCell('A').border = {
    left: { style: 'thin' },
    right: { style: 'thin' },
    top: { style: 'thin' },
    bottom: { style: 'thin' }
  }

  row.getCell('A').font = {
    bold: true, size: 14, name: 'Calibri', family: 2,
    color: { theme: 0 }
  }

  row.getCell('A').alignment = { vertical: 'middle' };

  row.getCell('A').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF00B0F0' },
    bgColor: { indexed: 64 }
  };

  row.getCell('A').value = text;
}

function buildReportFirstColumnType1(worksheet, rowIndex, text) {
  let row = worksheet.getRow(rowIndex);

  row.getCell('A').border = {
    left: { style: 'thin' },
    right: { style: 'thin' },
    top: { style: 'thin' },
    bottom: { style: 'thin' }
  }

  row.getCell('A').font = {
    size: 14, name: 'Calibri', family: 2
  }

  row.getCell('A').alignment = { horizontal: 'right', vertical: 'middle' };

  row.getCell('A').value = text;
}

function buildDataRow(worksheet, rowIndex, cellIndex) {
  let row = worksheet.getRow(rowIndex);
  let bold = false;
  let color = { argb: 'FF000000' };

  if (rowIndex == 5 || rowIndex == 6 || rowIndex == 7 || rowIndex == 10 || rowIndex == 13) {
    bold = true;
  }

  if (rowIndex == 6 || rowIndex == 7 || rowIndex == 10) {
    row.getCell(cellIndex).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B0F0' },
      bgColor: { indexed: 64 }
    };
  }

  if (rowIndex == 13) {
    row.getCell(cellIndex).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF0000' },
      bgColor: { indexed: 64 }
    };
  }

  if (rowIndex == 5) {
    color = { argb: 'FFFF0000' };
  }

  if (rowIndex == 6 || rowIndex == 7 || rowIndex == 10 || rowIndex == 13) {
    color = { theme: 0 };
  }

  row.getCell(cellIndex).border = {
    left: { style: 'thin' },
    right: { style: 'thin' },
    top: { style: 'thin' },
    bottom: { style: 'thin' }
  }

  row.getCell(cellIndex).font = {
    italic: true, bold: bold, size: 14, name: 'Calibri', family: 2,
    color: color
  }

  row.getCell(cellIndex).numFmt = '#,##0';

  row.getCell(cellIndex).value = 0;
}

function fillData(batch, filterType) {
  return new Promise((resolve, reject) => {
    let baseQuery = 'SELECT COUNT(*) AS TotalBase, coalesce(SUM(hasError),0) AS HasError,\
    coalesce(SUM(missingData),0) AS MissingData,\
    coalesce(SUM(missingName),0) AS MissingName,\
    coalesce(SUM(missingPhoneNumber),0) AS MissingPhoneNumber, \
    coalesce(SUM(missingAddress),0) As MissingAddress, \
    coalesce(SUM(missingModel),0) AS MissingModel, \
    coalesce(SUM(illogicalPhone),0) AS IllogicalPhone, \
    coalesce(SUM(illogicalPhoneFormat),0) AS IllogicalPhoneFormat, \
    coalesce(SUM(illogicalPhoneProvider),0) AS IllogicalPhoneProvider, \
    coalesce(SUM(duplicatedPhone),0) As DuplicatedPhone, \
    coalesce(SUM(duplicatedPhoneSameModel),0) As DuplicatedPhoneSameModel, \
    coalesce(SUM(duplicatedPhoneDiffModel),0) As DuplicatedPhoneDiffModel \
    FROM customers'

    let whereCondition = '';
    let joinTable = '';
    let params = {};

    if (batch !== '' && filterType !== 'All') {
      params = _.merge(params, {
        $batch: batch
      });
      if (whereCondition === '') {
        whereCondition = 'WHERE customers.batch = $batch'
      } else {
        whereCondition += " AND customers.batch = $batch";
      }
    }

    let query = baseQuery + ' ' + joinTable + ' ' + whereCondition + ';';

    db.get(query, params, (err, row) => {
      if (err) {
        return reject(err);
      }
      resolve(row);
    });
  });
}

function writeToTemplate(reportFilePath, rowData, cellIndex) {
  return new Promise((resolve, reject) => {
    let workbook = new Excel.Workbook();
    workbook.xlsx.readFile(reportFilePath).then((response) => {

      let worksheet = workbook.getWorksheet(1);
      let row;

      row = worksheet.getRow(5);
      row.getCell(cellIndex).value = rowData.TotalBase;

      row = worksheet.getRow(6);
      row.getCell(cellIndex).value = rowData.MissingPhoneNumber;

      row = worksheet.getRow(7);
      row.getCell(cellIndex).value = rowData.DuplicatedPhone;

      row = worksheet.getRow(8);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneSameModel;

      row = worksheet.getRow(9);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneDiffModel;

      row = worksheet.getRow(10);
      row.getCell(cellIndex).value = rowData.IllogicalPhone;

      row = worksheet.getRow(11);
      row.getCell(cellIndex).value = rowData.IllogicalPhoneFormat;

      row = worksheet.getRow(12);
      row.getCell(cellIndex).value = rowData.IllogicalPhoneProvider;

      row = worksheet.getRow(13);
      row.getCell(cellIndex).value = rowData.TotalBase - rowData.HasError;

      resolve(workbook.xlsx.writeFile(reportFilePath));
    });
  });
}
