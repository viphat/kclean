const Excel = require('exceljs')
const fs = require('fs')
const _ = require('lodash')

const validTitle = 'DATA CLEANING RESULT - VALID LIST'
const invalidTitle = 'DATA CLEANING RESULT - INVALID LIST'
const duplicationTitle = 'DATA CLEANING RESULT - DUPLICATION LIST'
// const duplicationWithAnotherAgencyTitle = 'DATA CLEANING RESULT - DUPLICATION WITH ANOTHER AGENCY LIST';
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
    sheetName = 'Duplication';
    worksheet = workbook.addWorksheet(sheetName, {});
    writeBaseTemplate(workbook, worksheet, duplicationTitle);
    // sheetName = 'Duplication With Another Agency';
    // worksheet = workbook.addWorksheet(sheetName, {});
    // writeBaseTemplate(workbook, worksheet, duplicationWithAnotherAgencyTitle);

    // Write to File
    workbook.xlsx.writeFile(outputPath).then(() => {
      resolve(workbook);
    });
  });
}

function writeBaseTemplate(workbook, worksheet, title) {
  worksheet.getColumn('A').width = 6;
  worksheet.getColumn('B').width = 16;
  worksheet.getColumn('C').width = 16;
  worksheet.getColumn('D').width = 16;
  worksheet.getColumn('E').width = 16;
  worksheet.getColumn('F').width = 16;
  worksheet.getColumn('G').width = 24;
  worksheet.getColumn('H').width = 24;
  worksheet.getColumn('I').width = 18;
  worksheet.getColumn('J').width = 18;
  worksheet.getColumn('K').width = 16;
  worksheet.getColumn('L').width = 13.8;
  worksheet.getColumn('M').width = 13.8;
  worksheet.getColumn('N').width = 24;
  worksheet.getColumn('O').width = 16;
  worksheet.getColumn('P').width = 16;
  worksheet.getColumn('Q').width = 16;
  worksheet.getColumn('R').width = 16;
  worksheet.getColumn('S').width = 16;
  worksheet.getColumn('T').width = 16;
  worksheet.getColumn('U').width = 16;
  worksheet.getColumn('V').width = 16;
  worksheet.getColumn('W').width = 16;
  worksheet.getColumn('X').width = 16;

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

  worksheet.getCell('A5').value = 'STT.'

  worksheet.mergeCells('B5:B6');
  worksheet.getCell('B5').font = worksheet.getCell('A5').font;
  worksheet.getCell('B5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('B5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('B5').border = worksheet.getCell('A5').border;
  worksheet.getCell('B5').value = 'Address / School';

  worksheet.mergeCells('C5:C6');
  worksheet.getCell('C5').font = worksheet.getCell('A5').font;
  worksheet.getCell('C5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('C5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('C5').border = worksheet.getCell('A5').border;
  worksheet.getCell('C5').value = 'Province';

  worksheet.mergeCells('D5:D6');
  worksheet.getCell('D5').font = worksheet.getCell('A5').font;
  worksheet.getCell('D5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('D5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('D5').border = worksheet.getCell('A5').border;
  worksheet.getCell('D5').value = 'District';

  worksheet.mergeCells('E5:E6');
  worksheet.getCell('E5').font = worksheet.getCell('A5').font;
  worksheet.getCell('E5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('E5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('E5').border = worksheet.getCell('A5').border;
  worksheet.getCell('E5').value = 'Date';

  worksheet.mergeCells('F5:F6');
  worksheet.getCell('F5').font = worksheet.getCell('A5').font;
  worksheet.getCell('F5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('F5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('F5').border = worksheet.getCell('A5').border;
  worksheet.getCell('F5').value = 'hh:ss';

  worksheet.mergeCells('G5:G6');
  worksheet.getCell('G5').font = worksheet.getCell('A5').font;
  worksheet.getCell('G5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('G5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('G5').border = worksheet.getCell('A5').border;
  worksheet.getCell('G5').value = 'First Name';

  worksheet.mergeCells('H5:H6');
  worksheet.getCell('H5').font = worksheet.getCell('A5').font;
  worksheet.getCell('H5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('H5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('H5').border = worksheet.getCell('A5').border;
  worksheet.getCell('H5').value = 'Last Name'

  worksheet.mergeCells('I5:I6');
  worksheet.getCell('I5').font = worksheet.getCell('A5').font;
  worksheet.getCell('I5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('I5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('I5').border = worksheet.getCell('A5').border;
  worksheet.getCell('I5').value = "Student's mobile"

  worksheet.mergeCells('J5:J6');
  worksheet.getCell('J5').font = worksheet.getCell('A5').font;
  worksheet.getCell('J5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('J5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('J5').border = worksheet.getCell('A5').border;
  worksheet.getCell('J5').value = "Parent's mobile"

  worksheet.mergeCells('K5:K6');
  worksheet.getCell('K5').font = worksheet.getCell('A5').font;
  worksheet.getCell('K5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('K5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('K5').border = worksheet.getCell('A5').border;
  worksheet.getCell('K5').value = 'Date of Birth (mm/dd/yyyy)'

  worksheet.mergeCells('L5:L6');
  worksheet.getCell('L5').font = worksheet.getCell('A5').font;
  worksheet.getCell('L5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('L5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('L5').border = worksheet.getCell('A5').border;
  worksheet.getCell('L5').value = 'Brand'

  worksheet.mergeCells('M5:M6');
  worksheet.getCell('M5').font = worksheet.getCell('A5').font;
  worksheet.getCell('M5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('M5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('M5').border = worksheet.getCell('A5').border;
  worksheet.getCell('M5').value = 'Sub-brand'

  worksheet.mergeCells('N5:N6');
  worksheet.getCell('N5').font = worksheet.getCell('A5').font;
  worksheet.getCell('N5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('N5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('N5').border = worksheet.getCell('A5').border;
  worksheet.getCell('N5').value = 'Sampling Product'

  worksheet.mergeCells('O5:O6');
  worksheet.getCell('O5').font = worksheet.getCell('A5').font;
  worksheet.getCell('O5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('O5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('O5').border = worksheet.getCell('A5').border;
  worksheet.getCell('O5').value = 'District ID'

  worksheet.mergeCells('P5:P6');
  worksheet.getCell('P5').font = worksheet.getCell('A5').font;
  worksheet.getCell('P5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('P5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('P5').border = worksheet.getCell('A5').border;
  worksheet.getCell('P5').value = 'Province ID'

  worksheet.mergeCells('Q5:Q6');
  worksheet.getCell('Q5').font = worksheet.getCell('A5').font;
  worksheet.getCell('Q5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('Q5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('Q5').border = worksheet.getCell('A5').border;
  worksheet.getCell('Q5').value = 'Opt In'

  worksheet.mergeCells('R5:R6');
  worksheet.getCell('R5').font = worksheet.getCell('A5').font;
  worksheet.getCell('R5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('R5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('R5').border = worksheet.getCell('A5').border;
  worksheet.getCell('R5').value = 'FW'

  worksheet.mergeCells('S5:S6');
  worksheet.getCell('S5').font = worksheet.getCell('A5').font;
  worksheet.getCell('S5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('S5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('S5').border = worksheet.getCell('A5').border;
  worksheet.getCell('S5').value = 'Khối'

  worksheet.mergeCells('T5:T6');
  worksheet.getCell('T5').font = worksheet.getCell('A5').font;
  worksheet.getCell('T5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('T5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('T5').border = worksheet.getCell('A5').border;
  worksheet.getCell('T5').value = 'Đại diện'

  worksheet.mergeCells('U5:U6');
  worksheet.getCell('U5').font = worksheet.getCell('A5').font;
  worksheet.getCell('U5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('U5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('U5').border = worksheet.getCell('A5').border;
  worksheet.getCell('U5').value = 'PG'

  worksheet.mergeCells('V5:V6');
  worksheet.getCell('V5').font = worksheet.getCell('A5').font;
  worksheet.getCell('V5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('V5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('V5').border = worksheet.getCell('A5').border;
  worksheet.getCell('V5').value = 'Activation'

  worksheet.mergeCells('W5:W6');
  worksheet.getCell('W5').font = worksheet.getCell('A5').font;
  worksheet.getCell('W5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('W5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('W5').border = worksheet.getCell('A5').border;
  worksheet.getCell('W5').value = 'Target'
  // End Table Headers

  if (worksheet.name.endsWith('Duplication')) {
    worksheet.mergeCells('X5:X6');
    worksheet.getCell('X5').font = worksheet.getCell('A5').font;
    worksheet.getCell('X5').fill = worksheet.getCell('A5').fill;
    worksheet.getCell('X5').alignment = worksheet.getCell('A5').alignment;
    worksheet.getCell('X5').border = worksheet.getCell('A5').border;
    worksheet.getCell('X5').value = 'Tuần';
  }

  // Add Logo
  let logo = workbook.addImage({
    filename: logoPath,
    extension: 'png'
  });

  worksheet.addImage(logo, 'A1:B3');
}
