const Excel = require('exceljs')
const fs = require('fs')

import _ from 'lodash'

const validTitle = 'DATA CLEANING RESULT - VALID LIST'
const invalidTitle = 'DATA CLEANING RESULT - INVALID LIST'
const invalidPhoneFormatTitle = 'DATA CLEANING RESULT - INVALID LIST - Phone Format'
const invalidPhoneProviderTitle = 'DATA CLEANING RESULT - INVALID LIST - Phone Provider'
const dupSameProductTitle = 'DATA CLEANING RESULT - DUPLICATION LIST - Same Model'
const dupDiffProductTitle = 'DATA CLEANING RESULT - DUPLICATION LIST - Different Model'
const logoPath = './vendor/logo.png';

export const buildExcelTemplate = (outputPath) => {
  let workbook = new Excel.Workbook();

  return new Promise((resolve, reject) => {
    if (!fs.existsSync(outputPath)) {
      return resolve(writeTemplate(outputPath, workbook));
    } else {
      workbook.xlsx.readFile(outputPath).then(() => {
        let sheetName = 'Valid';
        let worksheet = workbook.getWorksheet(sheetName);
        if (worksheet === undefined) {
          return resolve(writeTemplate(outputPath, workbook));
        }
        resolve(workbook);
      });
    }
  });
}

function writeTemplate(outputPath, workbook) {
  return new Promise((resolve, reject) => {
    let sheetName = 'Valid';
    let worksheet = workbook.addWorksheet(sheetName, {});
    writeBaseTemplate(workbook, worksheet, validTitle);

    sheetName = 'Invalid';
    worksheet = workbook.addWorksheet(sheetName, {});
    writeBaseTemplate(workbook, worksheet, invalidTitle);

    sheetName = 'Invalid - Phone Format';
    worksheet = workbook.addWorksheet(sheetName, {});
    writeBaseTemplate(workbook, worksheet, invalidPhoneFormatTitle);
    sheetName = 'Invalid - Phone Provider';
    worksheet = workbook.addWorksheet(sheetName, {});
    writeBaseTemplate(workbook, worksheet, invalidPhoneProviderTitle);

    sheetName = 'Duplicated - Same Model';
    worksheet = workbook.addWorksheet(sheetName, {});
    writeBaseTemplate(workbook, worksheet, dupSameProductTitle);
    sheetName = 'Duplicated - Different Model';
    worksheet = workbook.addWorksheet(sheetName, {});
    writeBaseTemplate(workbook, worksheet, dupDiffProductTitle);
    // Write to File
    workbook.xlsx.writeFile(outputPath).then(() => {
      resolve(workbook);
    });
  });
}

function writeBaseTemplate(workbook, worksheet, title) {
  worksheet.getColumn('A').width = 6;
  worksheet.getColumn('B').width = 30;
  worksheet.getColumn('C').width = 24;
  worksheet.getColumn('D').width = 40;
  worksheet.getColumn('E').width = 16;
  worksheet.getColumn('F').width = 20;
  worksheet.getColumn('G').width = 20;

  worksheet.getRow('5').height = 30;

  worksheet.getCell('E1').font = {
    bold: true, size: 14, name: 'Arial', family: 2,
    color: { argb: 'FFFF0000' }
  };

  worksheet.getCell('E1').alignment = { vertical: 'middle' };

  worksheet.getCell('E1').value = title;

  // Table Headers
  worksheet.mergeCells('A5:A6');

  worksheet.getCell('A5').font = {
    bold: true,
    size: 10,
    color: { theme: 1 },
    name: 'Arial',
    family: 2
  }

  worksheet.getCell('A5').fill =  {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF00' },
    bgColor: { indexed: 64 }
  }

  worksheet.getCell('A5').alignment = {
    horizontal: 'center', vertical: 'middle', wrapText: true
  }

  worksheet.getCell('A5').border = {
    left: { style: 'thin' },
    right: { style: 'thin' },
    top: { style: 'thin' },
    bottom: { style: 'thin' }
  }

  worksheet.getCell('A5').value = 'No.'

  worksheet.mergeCells('B5:B6');
  worksheet.getCell('B5').font = worksheet.getCell('A5').font;
  worksheet.getCell('B5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('B5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('B5').border = worksheet.getCell('A5').border;
  worksheet.getCell('B5').value = 'Name';

  worksheet.mergeCells('C5:C6');
  worksheet.getCell('C5').font = worksheet.getCell('A5').font;
  worksheet.getCell('C5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('C5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('C5').border = worksheet.getCell('A5').border;
  worksheet.getCell('C5').value = 'Phone';

  worksheet.mergeCells('D5:D6');
  worksheet.getCell('D5').font = worksheet.getCell('A5').font;
  worksheet.getCell('D5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('D5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('D5').border = worksheet.getCell('A5').border;
  worksheet.getCell('D5').value = 'Address';

  worksheet.mergeCells('E5:E6');
  worksheet.getCell('E5').font = worksheet.getCell('A5').font;
  worksheet.getCell('E5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('E5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('E5').border = worksheet.getCell('A5').border;
  worksheet.getCell('E5').value = 'T';

  worksheet.mergeCells('F5:F6');
  worksheet.getCell('F5').font = worksheet.getCell('A5').font;
  worksheet.getCell('F5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('F5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('F5').border = worksheet.getCell('A5').border;
  worksheet.getCell('F5').value = 'Phiên bản';

  if (worksheet.name.endsWith('Duplicated - Same Model') || worksheet.name.endsWith('Duplicated - Different Model')) {
    worksheet.mergeCells('G5:G6');
    worksheet.getCell('G5').font = worksheet.getCell('A5').font;
    worksheet.getCell('G5').fill = worksheet.getCell('A5').fill;
    worksheet.getCell('G5').alignment = worksheet.getCell('A5').alignment;
    worksheet.getCell('G5').border = worksheet.getCell('A5').border;
    worksheet.getCell('G5').value = 'Tuần';
  }

  // Add Logo
  let logo = workbook.addImage({
    filename: logoPath,
    extension: 'png'
  });

  worksheet.addImage(logo, 'A1:B3');
}
