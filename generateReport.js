const Excel = require('exceljs')
const fs = require('fs');
const _ = require('lodash');

import { db } from './database';

const logoPath = './vendor/logo.png';

export const generateReport = (batch, source, outputDirectory) => {
  return new Promise((resolve, reject) => {
    generateReportTemplate(batch, source, outputDirectory).then((reportFilePath) => {
      fillData(batch, source, 'All').then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'B');
      }).then(() => {
        return fillData(batch, source, 'ByBatch');
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'C');
      }).then(() => {
        return fillData(batch, source, { provinceId: 23 }); // Ho Chi Minh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'D');
      }).then(() => {
        return fillData(batch, source, { provinceId: 21 }); // Ha Noi
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'E');
      }).then(() => {
        return fillData(batch, source, { provinceId: 28 }); // Hai Phong
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'F');
      }).then(() => {
        return fillData(batch, source, { provinceId: 16 }); //. Da Nang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'G');
      }).then(() => {
        return fillData(batch, source, { provinceId: 33 }); // Khanh Hoa
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'H');
      }).then(() => {
        return fillData(batch, source, { provinceId: 13 }); // Can Tho
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'I');
      }).then(() => {
        return fillData(batch, source, { provinceId: 17 }); // Dong Nai
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'J');
      }).then(() => {
        return fillData(batch, source, { provinceId: 3 }); // Binh Duong
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'K');
      }).then(() => {
        return fillData(batch, source, { provinceId: 1 }); // An Giang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'L');
      }).then(() => {
        resolve(reportFilePath);
      })
    })
  })
}

const generateReportTemplate = (batch, source, outputDirectory) => {
  return new Promise((resolve, reject) => {
    let dir = outputDirectory + '/' + batch;

    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir)
    }

    let reportFilePath = dir + '/' + batch + '_' + source + '_report.xlsx';

    let workbook = new Excel.Workbook();
    let worksheet = workbook.addWorksheet('Abs', {});

    worksheet.getColumn('A').width = 60;
    worksheet.getRow(1).height = 50;
    worksheet.getRow(4).height = 30;
    worksheet.getRow(5).height = 40;

    // Add Logo
    let logo = workbook.addImage({
      filename: logoPath,
      extension: 'png'
    });

    worksheet.addImage(logo, {
      tl: { col: 0, row: 0 },
      br: { col: 1, row: 1 }
    });

    worksheet.getColumn('B').width = 30;
    worksheet.getColumn('C').width = 30;
    worksheet.getColumn('D').width = 30;
    worksheet.getColumn('E').width = 30;
    worksheet.getColumn('F').width = 30;
    worksheet.getColumn('G').width = 30;
    worksheet.getColumn('H').width = 30;
    worksheet.getColumn('I').width = 30;
    worksheet.getColumn('J').width = 30;
    worksheet.getColumn('K').width = 30;
    worksheet.getColumn('L').width = 30;
    // A1

    worksheet.getCell('B1').value = 'KOTEX CALL CENTER 2020 PROJECT';

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
    worksheet.getCell('A5').border = {
      left: { style: 'thin' },
      right: { style: 'thin' },
      top: { style: 'thin' },
      bottom: { style: 'thin' }
    }

    worksheet.getCell('A5').font = {
      bold: true, size: 14, name: 'Calibri', family: 2
    }

    worksheet.getCell('A5').alignment = { horizontal: 'center', vertical: 'middle' };

    worksheet.getCell('A5').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFABF8F' },
      bgColor: { indexed: 64 }
    };

    worksheet.getCell('A5').value = batch;

    // A6, A27
    buildReportFirstColumnType3(worksheet, 6, 'Raw data received from ' + source);
    buildReportFirstColumnType3(worksheet, 21, 'Valid database (value) - base all');

    // A7, A14, A21
    buildReportFirstColumnType2(worksheet, 7, 'Data missing');
    buildReportFirstColumnType2(worksheet, 16, 'Duplicated Data (Checking vs. total database since 1st week)');
    buildReportFirstColumnType2(worksheet, 17, 'Illogical data');

    // A8 - A13, A15 - A20, A22-A25
    buildReportFirstColumnType1(worksheet, 8, "Respondent's name (column G+H)");
    buildReportFirstColumnType1(worksheet, 9, "Living city (column C+D)");
    buildReportFirstColumnType1(worksheet, 10, "Contact information\nBRAND MAX/ HIGH SCHOOL: must have either phone number (column I), parents number (column J)\nFOCUS MKT/ UNIVERSITY: must have phone number (column I)");
    buildReportFirstColumnType1(worksheet, 11, "Birth Year (column K)");
    buildReportFirstColumnType1(worksheet, 12, "Sampling Date (column E)");
    buildReportFirstColumnType1(worksheet, 13, "School Name (column B+C+D)");
    buildReportFirstColumnType1(worksheet, 14, "Brand Using (column L)");
    buildReportFirstColumnType1(worksheet, 15, "Sampling Type\nFOCUS MKT/ UNIVERSITY: column N\nBRAND MAX/ HIGH SCHOOL: ok to be blank");
    buildReportFirstColumnType1(worksheet, 18, "Illogical phone number format (not 03x, 05x, 07x, 08x, 09x)");
    buildReportFirstColumnType1(worksheet, 19, "Illogical Age - High School (not 2002 - 2006)");
    buildReportFirstColumnType1(worksheet, 20, "Illogical Age - University (not 1998 - 2003)");
    // Done 1st Col

    // Row 4 - D4, K4, P4, S4
    buildReportRow4(worksheet, 'D', 'D4:L4', 'Break-down by city');

    // // Row 5, B4-T4
    buildReportRow5(worksheet, 'B', 'Total Project');
    buildReportRow5(worksheet, 'C', 'Total ' + batch);
    buildReportRow5(worksheet, 'D', 'Hồ Chí Minh');
    buildReportRow5(worksheet, 'E', 'Hà Nội');
    buildReportRow5(worksheet, 'F', 'Hải Phòng');
    buildReportRow5(worksheet, 'G', 'Đà Nẵng');
    buildReportRow5(worksheet, 'H', 'Nha Trang');
    buildReportRow5(worksheet, 'I', 'Cần Thơ');
    buildReportRow5(worksheet, 'J', 'Biên Hòa');
    buildReportRow5(worksheet, 'K', 'Bình Dương');
    buildReportRow5(worksheet, 'L', 'An Giang');

    // Data
    let colArr = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'];
    let rowArr = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21];

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
  let row = worksheet.getRow(5);
  let fgColor = { theme: 0, tint: -0.1499984740745262 };

  if (cellIndex == 'B') {
    fgColor = { theme: 2, tint: -0.249977111117893 };
  }

  if (cellIndex == 'C') {
    fgColor = { theme: 5, tint: 0.5999938962981048 };
  }

  if (cellIndex == 'D' || cellIndex == 'E' || cellIndex == 'F' || cellIndex == 'G' ||
    cellIndex == 'H' || cellIndex == 'I' || cellIndex == 'J' || cellIndex == 'K' || cellIndex == 'L'
  ) {
    fgColor = { argb: 'FFFFFF00' };
  }

  // if (cellIndex == 'J' || cellIndex == 'K' || cellIndex == 'L') {
  //   fgColor = { theme: 6, tint: 0.3999755851924192 };
  // }

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
    // case 'J':
    //   fgColor = { theme: 6, tint: 0.3999755851924192 };
    //   break;
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

  row.getCell('A').alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };

  row.getCell('A').value = text;
}

function buildDataRow(worksheet, rowIndex, cellIndex) {
  let row = worksheet.getRow(rowIndex);
  let bold = false;
  let color = { argb: 'FF000000' };

  if (rowIndex == 6 || rowIndex == 7 || rowIndex == 16 || rowIndex == 17 || rowIndex == 21) {
    bold = true;
  }

  if (rowIndex == 7 || rowIndex == 16 || rowIndex == 17) {
    row.getCell(cellIndex).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B0F0' },
      bgColor: { indexed: 64 }
    };
  }

  if (rowIndex == 21) {
    row.getCell(cellIndex).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF0000' },
      bgColor: { indexed: 64 }
    };
  }

  if (rowIndex == 6) {
    color = { argb: 'FFFF0000' };
  }

  if (rowIndex == 7 || rowIndex == 16 || rowIndex == 17 || rowIndex == 21) {
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

function fillData(batch, source, filterType) {
  return new Promise((resolve, reject) => {
    let baseQuery = 'SELECT COUNT(*) AS TotalBase, coalesce(SUM(hasError),0) AS HasError,\
    coalesce(SUM(missingData),0) AS MissingData,\
    coalesce(SUM(missingName),0) AS MissingName, coalesce(SUM(missingLivingCity),0) AS MissingLivingCity,\
    coalesce(SUM(missingContactInformation),0) AS MissingContactInformation, \
    coalesce(SUM(missingAge),0) As MissingAge, \
    coalesce(SUM(missingSchoolName),0) AS MissingSchoolName, \
    coalesce(SUM(missingCollectedDate),0) AS MissingCollectedDate, \
    coalesce(SUM(missingBrandUsing),0) AS MissingBrandUsing, \
    coalesce(SUM(missingSamplingType),0) AS MissingSamplingType, \
    coalesce(SUM(illogicalData),0) As IllogicalData, \
    coalesce(SUM(illogicalPhone),0) AS IllogicalPhone,\
    coalesce(SUM(illogicalAge),0) AS IllogicalAge,\
    coalesce(SUM(illogicalAgePupil),0) AS IllogicalAgePupil,\
    coalesce(SUM(illogicalAgeStudent),0) AS IllogicalAgeStudent,\
    coalesce(SUM(duplicatedPhone),0) As DuplicatedPhone, \
    coalesce(SUM(duplicatedPhoneBetweenPupilAndStudent),0) As DuplicatedPhoneBetweenPupilAndStudent, \
    coalesce(SUM(duplicatedPhoneWithinPupil),0) AS DuplicatedPhoneWithinPupil,\
    coalesce(SUM(duplicatedPhoneWithinStudent),0) AS DuplicatedPhoneWithinStudent\
    FROM customers'

    let whereCondition = '';
    let joinTable = '';
    let params = {};

    if (filterType.provinceId !== undefined && filterType.provinceId !== null) {
      whereCondition = 'WHERE customers.provinceId = $provinceId'
      params = {
        $provinceId: filterType.provinceId
      }
    }

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

    if (source !== '') {
      params = _.merge(params, {
        $source: source
      });
      if (whereCondition === '') {
        whereCondition = 'WHERE customers.source = $source'
      } else {
        whereCondition += " AND customers.source = $source";
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

      row = worksheet.getRow(6);
      row.getCell(cellIndex).value = rowData.TotalBase;

      row = worksheet.getRow(7);
      row.getCell(cellIndex).value = rowData.MissingData;

      row = worksheet.getRow(8);
      row.getCell(cellIndex).value = rowData.MissingName;

      row = worksheet.getRow(9);
      row.getCell(cellIndex).value = rowData.MissingLivingCity;

      row = worksheet.getRow(10);
      row.getCell(cellIndex).value = rowData.MissingContactInformation;

      row = worksheet.getRow(11);
      row.getCell(cellIndex).value = rowData.MissingAge;

      row = worksheet.getRow(12);
      row.getCell(cellIndex).value = rowData.MissingCollectedDate;

      row = worksheet.getRow(13);
      row.getCell(cellIndex).value = rowData.MissingSchoolName;

      row = worksheet.getRow(14);
      row.getCell(cellIndex).value = rowData.MissingBrandUsing;

      row = worksheet.getRow(15);
      row.getCell(cellIndex).value = rowData.MissingSamplingType;

      row = worksheet.getRow(16);
      row.getCell(cellIndex).value = rowData.DuplicatedPhone;

      row = worksheet.getRow(17);
      row.getCell(cellIndex).value = rowData.IllogicalData;

      row = worksheet.getRow(18);
      row.getCell(cellIndex).value = rowData.IllogicalPhone;

      row = worksheet.getRow(19);
      row.getCell(cellIndex).value = rowData.IllogicalAgePupil;

      row = worksheet.getRow(20);
      row.getCell(cellIndex).value = rowData.IllogicalAgeStudent;

      row = worksheet.getRow(21);
      row.getCell(cellIndex).value = rowData.TotalBase - rowData.HasError;

      resolve(workbook.xlsx.writeFile(reportFilePath));
    });
  });
}
