const Excel = require('exceljs')
const fs = require('fs');
const _ = require('lodash');

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
        return fillData(batch, { provinceId: 1 });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'D');
      }).then(() => {
        return fillData(batch, { provinceId: 2 });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'E');
      }).then(() => {
        return fillData(batch, { provinceId: 3 });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'F');
      }).then(() => {
        return fillData(batch, { provinceId: 4 });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'G');
      }).then(() => {
        return fillData(batch, { provinceId: 5 });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'H');
      }).then(() => {
        return fillData(batch, { provinceId: 6 });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'I');
      }).then(() => {
        return fillData(batch, { groupId: 1 });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'J');
      }).then(() => {
        return fillData(batch, { groupId: 2 });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'K');
      }).then(() => {
        return fillData(batch, { groupId: 3 });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'L');
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
    buildReportFirstColumnType3(worksheet, 6, 'Raw data received from Agency');
    buildReportFirstColumnType3(worksheet, 27, 'Valid database (value) - base all');

    // A7, A14, A21
    buildReportFirstColumnType2(worksheet, 7, 'Data missing');
    buildReportFirstColumnType2(worksheet, 14, 'Duplicated Data (Checking vs. total database since 1st week)');
    buildReportFirstColumnType2(worksheet, 21, 'Illogical data');

    // A8 - A13, A15 - A20, A22-A25
    buildReportFirstColumnType1(worksheet, 8, "Respondent's name");
    buildReportFirstColumnType1(worksheet, 9, "Living city");
    buildReportFirstColumnType1(worksheet, 10, "Contact information");
    buildReportFirstColumnType1(worksheet, 11, "Age");
    buildReportFirstColumnType1(worksheet, 12, "School name");
    buildReportFirstColumnType1(worksheet, 13, "Brand using");
    buildReportFirstColumnType1(worksheet, 15, "Duplication Pupil/ Student");
    buildReportFirstColumnType1(worksheet, 16, "Duplication Pupil/ Others");
    buildReportFirstColumnType1(worksheet, 17, "Duplication Student/ Others");
    buildReportFirstColumnType1(worksheet, 18, "Duplication between Pupil");
    buildReportFirstColumnType1(worksheet, 19, "Duplication between Student");
    buildReportFirstColumnType1(worksheet, 20, "Duplication between Others");
    buildReportFirstColumnType1(worksheet, 22, "Illogical phone number format");
    buildReportFirstColumnType1(worksheet, 23, "Illogical age format (not 2 digit)");
    buildReportFirstColumnType1(worksheet, 24, "Pupil but <2001 (more than 20 years old)");
    buildReportFirstColumnType1(worksheet, 25, "Student but >2002 (<18 years old) or <1996 (>24 years old)");
    buildReportFirstColumnType1(worksheet, 26, "Illogical age of Others (<1960)");
    // Done 1st Col

    // Row 4 - D4, K4, P4, S4
    buildReportRow4(worksheet, 'D', 'D4:I4', 'Break-down by city');
    buildReportRow4(worksheet, 'J', 'J4:L4', 'Target');

    // // Row 5, B4-T4
    buildReportRow5(worksheet, 'B', 'Total Project');
    buildReportRow5(worksheet, 'C', 'Total ' + batch);
    buildReportRow5(worksheet, 'D', 'Hồ Chí Minh');
    buildReportRow5(worksheet, 'E', 'Cần  Thơ');
    buildReportRow5(worksheet, 'F', 'Vĩnh Long');
    buildReportRow5(worksheet, 'G', 'Đồng Nai');
    buildReportRow5(worksheet, 'H', 'Đà Nẵng');
    buildReportRow5(worksheet, 'I', 'Huế');
    buildReportRow5(worksheet, 'J', 'Pupil/ Học sinh');
    buildReportRow5(worksheet, 'K', 'Student/ Sinh viên');
    buildReportRow5(worksheet, 'L', 'Others/ Khác');

    // Data
    let colArr = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'];
    let rowArr = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27];

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

  if (rowIndex == 6 || rowIndex == 7 || rowIndex == 14 || rowIndex == 21 || rowIndex == 27) {
    bold = true;
  }

  if (rowIndex == 7 || rowIndex == 14 || rowIndex == 21) {
    row.getCell(cellIndex).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B0F0' },
      bgColor: { indexed: 64 }
    };
  }

  if (rowIndex == 27) {
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

  if (rowIndex == 7 || rowIndex == 14 || rowIndex == 21 || rowIndex == 27) {
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
    coalesce(SUM(missingName),0) AS MissingName, coalesce(SUM(missingLivingCity),0) AS MissingLivingCity,\
    coalesce(SUM(missingContactInformation),0) AS MissingContactInformation, \
    coalesce(SUM(missingAge),0) As MissingAge, \
    coalesce(SUM(missingSchoolName),0) AS MissingSchoolName, \
    coalesce(SUM(missingBrandUsing),0) AS MissingBrandUsing, \
    coalesce(SUM(illogicalData),0) As IllogicalData, \
    coalesce(SUM(illogicalPhone),0) AS IllogicalPhone,\
    coalesce(SUM(illogicalAge),0) AS IllogicalAge,\
    coalesce(SUM(illogicalAgePupil),0) AS IllogicalAgePupil,\
    coalesce(SUM(illogicalAgeStudent),0) AS IllogicalAgeStudent,\
    coalesce(SUM(illogicalAgeOthers),0) AS IllogicalAgeOthers,\
    coalesce(SUM(duplicatedPhone),0) As DuplicatedPhone, \
    coalesce(SUM(duplicatedPhoneBetweenPupilAndStudent),0) As DuplicatedPhoneBetweenPupilAndStudent, \
    coalesce(SUM(duplicatedPhoneBetweenPupilAndOthers),0) AS DuplicatedPhoneBetweenPupilAndOthers,\
    coalesce(SUM(duplicatedPhoneBetweenStudentAndOthers),0) AS DuplicatedPhoneBetweenStudentAndOthers,\
    coalesce(SUM(duplicatedPhoneWithinPupil),0) AS DuplicatedPhoneWithinPupil,\
    coalesce(SUM(duplicatedPhoneWithinStudent),0) AS DuplicatedPhoneWithinStudent,\
    coalesce(SUM(duplicatedPhoneWithinOthers),0) AS DuplicatedPhoneWithinOthers\
    FROM customers'

    let whereCondition = '';
    let joinTable = '';
    let params = {};

    if (filterType.groupId && filterType.groupId >= 1 && filterType.groupId <= 3){
      whereCondition = 'WHERE customers.groupId = $groupId';
      params = {
        $groupId: filterType.groupId
      }
    } else if (filterType.provinceId !== undefined && filterType.provinceId !== null) {
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
      row.getCell(cellIndex).value = rowData.MissingSchoolName;

      row = worksheet.getRow(13);
      row.getCell(cellIndex).value = rowData.MissingBrandUsing;

      row = worksheet.getRow(14);
      row.getCell(cellIndex).value = rowData.DuplicatedPhone;

      row = worksheet.getRow(15);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneBetweenPupilAndStudent;

      row = worksheet.getRow(16);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneBetweenPupilAndOthers;

      row = worksheet.getRow(17);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneBetweenStudentAndOthers;

      row = worksheet.getRow(18);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneWithinPupil;

      row = worksheet.getRow(19);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneWithinStudent;

      row = worksheet.getRow(20);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneWithinOthers;

      row = worksheet.getRow(21);
      row.getCell(cellIndex).value = rowData.IllogicalData;

      row = worksheet.getRow(22);
      row.getCell(cellIndex).value = rowData.IllogicalPhone;

      row = worksheet.getRow(23);
      row.getCell(cellIndex).value = rowData.IllogicalAge;

      row = worksheet.getRow(24);
      row.getCell(cellIndex).value = rowData.IllogicalAgePupil;

      row = worksheet.getRow(25);
      row.getCell(cellIndex).value = rowData.IllogicalAgeStudent;

      row = worksheet.getRow(26);
      row.getCell(cellIndex).value = rowData.IllogicalAgeOthers;

      row = worksheet.getRow(27);
      row.getCell(cellIndex).value = rowData.TotalBase - rowData.HasError;

      resolve(workbook.xlsx.writeFile(reportFilePath));
    });
  });
}
