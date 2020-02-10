const Excel = require('exceljs')
const fs = require('fs')

import _ from 'lodash'

const validTitle = 'DATA CLEANING RESULT - VALID LIST'
const invalidTitle = 'DATA CLEANING RESULT - INVALID LIST'
const duplicationTitle = 'DATA CLEANING RESULT - DUPLICATION LIST'
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
  worksheet.getColumn('E').width = 24;
  worksheet.getColumn('F').width = 6;
  worksheet.getColumn('G').width = 12;
  worksheet.getColumn('H').width = 12;
  worksheet.getColumn('I').width = 12;
  worksheet.getColumn('J').width = 12;
  worksheet.getColumn('K').width = 13.8;
  worksheet.getColumn('L').width = 13.8;
  worksheet.getColumn('M').width = 13.8;
  worksheet.getColumn('N').width = 13.8;
  worksheet.getColumn('O').width = 13.8;
  worksheet.getColumn('P').width = 20;
  worksheet.getColumn('Q').width = 20;
  worksheet.getColumn('R').width = 10;
  worksheet.getColumn('S').width = 10;
  worksheet.getColumn('T').width = 10;

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
  worksheet.getCell('B5').value = 'Khu Vực';

  worksheet.mergeCells('C5:C6');
  worksheet.getCell('C5').font = worksheet.getCell('A5').font;
  worksheet.getCell('C5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('C5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('C5').border = worksheet.getCell('A5').border;
  worksheet.getCell('C5').value = 'Tỉnh Thành';

  worksheet.mergeCells('D5:D6');
  worksheet.getCell('D5').font = worksheet.getCell('A5').font;
  worksheet.getCell('D5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('D5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('D5').border = worksheet.getCell('A5').border;
  worksheet.getCell('D5').value = 'Địa Điểm';

  worksheet.mergeCells('E5:E6');
  worksheet.getCell('E5').font = worksheet.getCell('A5').font;
  worksheet.getCell('E5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('E5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('E5').border = worksheet.getCell('A5').border;
  worksheet.getCell('E5').value = 'Khách Hàng';

  worksheet.mergeCells('F5:F6');
  worksheet.getCell('F5').font = worksheet.getCell('A5').font;
  worksheet.getCell('F5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('F5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('F5').border = worksheet.getCell('A5').border;
  worksheet.getCell('F5').value = 'Năm sinh';

  worksheet.mergeCells('G5:G6');
  worksheet.getCell('G5').font = worksheet.getCell('A5').font;
  worksheet.getCell('G5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('G5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('G5').border = worksheet.getCell('A5').border;
  worksheet.getCell('G5').value = 'Số Điện Thoại';

  worksheet.mergeCells('H5:H6');
  worksheet.getCell('H5').font = worksheet.getCell('A5').font;
  worksheet.getCell('H5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('H5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('H5').border = worksheet.getCell('A5').border;
  worksheet.getCell('H5').value = 'Số Điện Thoại Phụ Huynh'

  worksheet.mergeCells('I5:I6');
  worksheet.getCell('I5').font = worksheet.getCell('A5').font;
  worksheet.getCell('I5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('I5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('I5').border = worksheet.getCell('A5').border;
  worksheet.getCell('I5').value = 'Facebook'

  worksheet.mergeCells('J5:J6');
  worksheet.getCell('J5').font = worksheet.getCell('A5').font;
  worksheet.getCell('J5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('J5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('J5').border = worksheet.getCell('A5').border;
  worksheet.getCell('J5').value = 'Email'

  worksheet.mergeCells('K5:O5');
  worksheet.getCell('K5').font = worksheet.getCell('A5').font;
  worksheet.getCell('K5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('K5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('K5').border = worksheet.getCell('A5').border;
  worksheet.getCell('K5').value = 'Sản phẩm bạn đang dùng'

  worksheet.getCell('K6').font = worksheet.getCell('A5').font;
  worksheet.getCell('K6').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('K6').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('K6').border = worksheet.getCell('A5').border;
  worksheet.getCell('K6').value = 'Kotex'

  worksheet.getCell('L6').font = worksheet.getCell('A5').font;
  worksheet.getCell('L6').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('L6').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('L6').border = worksheet.getCell('A5').border;
  worksheet.getCell('L6').value = 'Diana'

  worksheet.getCell('M6').font = worksheet.getCell('A5').font;
  worksheet.getCell('M6').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('M6').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('M6').border = worksheet.getCell('A5').border;
  worksheet.getCell('M6').value = 'Laurier';

  worksheet.getCell('N6').font = worksheet.getCell('A5').font;
  worksheet.getCell('N6').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('N6').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('N6').border = worksheet.getCell('A5').border;
  worksheet.getCell('N6').value = 'Whisper'

  worksheet.getCell('O6').font = worksheet.getCell('A5').font;
  worksheet.getCell('O6').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('O6').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('O6').border = worksheet.getCell('A5').border;
  worksheet.getCell('O6').value = 'Khác'

  worksheet.mergeCells('P5:P6');
  worksheet.getCell('P5').font = worksheet.getCell('A5').font;
  worksheet.getCell('P5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('P5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('P5').border = worksheet.getCell('A5').border;
  worksheet.getCell('P5').value = 'Ghi chú'

  worksheet.mergeCells('Q5:Q6');
  worksheet.getCell('Q5').font = worksheet.getCell('A5').font;
  worksheet.getCell('Q5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('Q5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('Q5').border = worksheet.getCell('A5').border;
  worksheet.getCell('Q5').value = 'Ngày Nhập'

  worksheet.mergeCells('R5:R6');
  worksheet.getCell('R5').font = worksheet.getCell('A5').font;
  worksheet.getCell('R5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('R5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('R5').border = worksheet.getCell('A5').border;
  worksheet.getCell('R5').value = 'Nhận'

  worksheet.mergeCells('S5:S6');
  worksheet.getCell('S5').font = worksheet.getCell('A5').font;
  worksheet.getCell('S5').fill = worksheet.getCell('A5').fill;
  worksheet.getCell('S5').alignment = worksheet.getCell('A5').alignment;
  worksheet.getCell('S5').border = worksheet.getCell('A5').border;
  worksheet.getCell('S5').value = 'Đối Tượng'
  // End Table Headers

  if (worksheet.name.endsWith('Duplication')) {
    worksheet.mergeCells('T5:T6');
    worksheet.getCell('T5').font = worksheet.getCell('A5').font;
    worksheet.getCell('T5').fill = worksheet.getCell('A5').fill;
    worksheet.getCell('T5').alignment = worksheet.getCell('A5').alignment;
    worksheet.getCell('T5').border = worksheet.getCell('A5').border;
    worksheet.getCell('T5').value = 'Tuần';
  }

  // Add Logo
  let logo = workbook.addImage({
    filename: logoPath,
    extension: 'png'
  });

  worksheet.addImage(logo, 'A1:B3');
}
