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
        return fillData(batch, { target: "HIGH SCHOOL" });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'C');
      }).then(() => {
        return fillData(batch, { target: "UNIVERSITY" });
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'D');
      }).then(() => {
        return fillData(batch, { provinceId: 23, target: "HIGH SCHOOL" }); // Ho Chi Minh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'E');
      }).then(() => {
        return fillData(batch, { provinceId: 21, target: "HIGH SCHOOL" }); // Ha Noi
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'F');
      }).then(() => {
        return fillData(batch, { provinceId: 28, target: "HIGH SCHOOL" }); // Hai Phong
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'G');
      }).then(() => {
        return fillData(batch, { provinceId: 16, target: "HIGH SCHOOL" }); //. Da Nang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'H');
      }).then(() => {
        return fillData(batch, { provinceId: 33, target: "HIGH SCHOOL" }); // Nha Trang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'I');
      }).then(() => {
        return fillData(batch, { provinceId: 23, target: "UNIVERSITY" }); // Ho Chi Minh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'J');
      }).then(() => {
        return fillData(batch, { provinceId: 21, target: "UNIVERSITY" }); // Ha Noi
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'K');
      }).then(() => {
        return fillData(batch, { provinceId: 28, target: "UNIVERSITY" }); // Hai Phong
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'L');
      }).then(() => {
        return fillData(batch, { provinceId: 16, target: "UNIVERSITY" }); //. Da Nang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'M');
      }).then(() => {
        return fillData(batch, { provinceId: 33, target: "UNIVERSITY" }); // Nha Trang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'N');
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
    worksheet.getColumn('M').width = 30;
    worksheet.getColumn('N').width = 30;
    // A1

    worksheet.getCell('B1').value = 'KOTEX CALL CENTER 2023 PROJECT';

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
    buildReportFirstColumnType3(worksheet, 23, 'Valid database (value) - base all');

    buildReportFirstColumnType2(worksheet, 7, 'Data missing (=blank)');
    buildReportFirstColumnType2(worksheet, 15, 'Duplicated Data (Checking vs. total database since 1st week)');
    buildReportFirstColumnType2(worksheet, 19, 'Illogical data');

    buildReportFirstColumnType1(worksheet, 8, "Respondent's name");
    buildReportFirstColumnType1(worksheet, 9, "Living city");
    buildReportFirstColumnType1(worksheet, 10, "Contact information\nHIGH SCHOOL: must have either phone number, parents number\nUniversity: must have phone number");
    buildReportFirstColumnType1(worksheet, 11, "Birth Year");
    buildReportFirstColumnType1(worksheet, 12, "Sampling Date");
    buildReportFirstColumnType1(worksheet, 13, "School Name");
    buildReportFirstColumnType1(worksheet, 14, "Brand Using");

    buildReportFirstColumnType1(worksheet, 16, "Duplication within High School");
    buildReportFirstColumnType1(worksheet, 17, "Duplication within University");
    buildReportFirstColumnType1(worksheet, 18, "Duplication between High School & University");

    buildReportFirstColumnType1(worksheet, 20, "Illogical phone number format (not 03x, 05x, 07x, 08x, 09x)");
    buildReportFirstColumnType1(worksheet, 21, "Illogical Age - High School (not 2005 - 2008)");
    buildReportFirstColumnType1(worksheet, 22, "Illogical Age - University (not 2001 - 2005)");
    // Done 1st Col

    // Row 4 - D4, K4, P4, S4
    buildReportRow4(worksheet, 'E', 'E4:I4', 'Break-down by city (Highschool)');
    buildReportRow4(worksheet, 'J', 'J4:N4', 'Break-down by city (University)');

    // // Row 5, B4-T4
    buildReportRow5(worksheet, 'B', 'Total Project');
    buildReportRow5(worksheet, 'C', 'Total ' + batch + ' (High school)');
    buildReportRow5(worksheet, 'D', 'Total ' + batch + ' (University)');

    buildReportRow5(worksheet, 'E', 'Hồ Chí Minh');
    buildReportRow5(worksheet, 'F', 'Hà Nội');
    buildReportRow5(worksheet, 'G', 'Hải Phòng');
    buildReportRow5(worksheet, 'H', 'Đà Nẵng');
    buildReportRow5(worksheet, 'I', 'Nha Trang');

    buildReportRow5(worksheet, 'J', 'Hồ Chí Minh');
    buildReportRow5(worksheet, 'K', 'Hà Nội');
    buildReportRow5(worksheet, 'L', 'Hải Phòng');
    buildReportRow5(worksheet, 'M', 'Đà Nẵng');
    buildReportRow5(worksheet, 'N', 'Nha Trang');

    // Data
    let colArr = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N'];
    let rowArr = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23];

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
    case 'E':
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

  row.getCell('A').alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };

  row.getCell('A').value = text;
}

function buildDataRow(worksheet, rowIndex, cellIndex) {
  let row = worksheet.getRow(rowIndex);
  let bold = false;
  let color = { argb: 'FF000000' };

  if (rowIndex == 6 || rowIndex == 7 || rowIndex == 14 || rowIndex == 15 || rowIndex == 18) {
    bold = true;
  }

  if (rowIndex == 7 || rowIndex == 14 || rowIndex == 15) {
    row.getCell(cellIndex).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B0F0' },
      bgColor: { indexed: 64 }
    };
  }

  if (rowIndex == 18) {
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

  if (rowIndex == 7 || rowIndex == 14 || rowIndex == 15 || rowIndex == 18) {
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
    coalesce(SUM(missingContactInformation),0) AS MissingContactInformation, \
    coalesce(SUM(missingAge),0) As MissingAge, \
    coalesce(SUM(missingSchoolName),0) AS MissingSchoolName, \
    coalesce(SUM(missingCollectedDate),0) AS MissingCollectedDate, \
    coalesce(SUM(missingBrandUsing),0) AS MissingBrandUsing, \
    coalesce(SUM(illogicalData),0) As IllogicalData, \
    coalesce(SUM(illogicalPhone),0) AS IllogicalPhone,\
    coalesce(SUM(illogicalAge),0) AS IllogicalAge,\
    coalesce(SUM(illogicalAgePupil),0) AS IllogicalAgePupil,\
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

    if (filterType.target !== undefined && filterType.target !== null) {
      if (whereCondition === '') {
        whereCondition = 'WHERE customers.target = $target'
      } else {
        whereCondition += " AND customers.target = $target";
      }

      params = {
        $target: filterType.target
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
      row.getCell(cellIndex).value = rowData.MissingCollectedDate;

      row = worksheet.getRow(13);
      row.getCell(cellIndex).value = rowData.MissingSchoolName;

      row = worksheet.getRow(14);
      row.getCell(cellIndex).value = rowData.MissingBrandUsing;

      row = worksheet.getRow(15);
      row.getCell(cellIndex).value = rowData.DuplicatedPhone;

      row = worksheet.getRow(16);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneWithinPupil;

      row = worksheet.getRow(17);
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneWithinStudent;

      row = worksheet.getRow(18);
      row.getCell(cellIndex).value = rowData.duplicatedPhoneBetweenPupilAndStudent;

      row = worksheet.getRow(19);
      row.getCell(cellIndex).value = rowData.IllogicalData;

      row = worksheet.getRow(20);
      row.getCell(cellIndex).value = rowData.IllogicalPhone;

      row = worksheet.getRow(21);
      row.getCell(cellIndex).value = rowData.IllogicalAgePupil;

      row = worksheet.getRow(22);
      row.getCell(cellIndex).value = rowData.IllogicalAgeStudent;

      row = worksheet.getRow(23);
      row.getCell(cellIndex).value = rowData.TotalBase - rowData.HasError;

      resolve(workbook.xlsx.writeFile(reportFilePath));
    });
  });
}
