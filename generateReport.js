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
        return fillData(batch, { provinceId: 16, target: "HIGH SCHOOL" }); // Da Nang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'H');
      }).then(() => {
        return fillData(batch, { provinceId: 57, target: "HIGH SCHOOL" }); // Thái Nguyên
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'I');
      }).then(() => {
        return fillData(batch, { provinceId: 48, target: "HIGH SCHOOL" }); // Quảng Ninh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'J');
      }).then(() => {
        return fillData(batch, { provinceId: 24, target: "HIGH SCHOOL" }); // Hải Dương
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'K');
      }).then(() => {
        return fillData(batch, { provinceId: 31, target: "HIGH SCHOOL" }); // Hưng Yên
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'L');
      }).then(() => {
        return fillData(batch, { provinceId: 42, target: "HIGH SCHOOL" }); // Nam Định
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'M');
      }).then(() => {
        return fillData(batch, { provinceId: 53, target: "HIGH SCHOOL" }); // Thái Bình
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'N');
      }).then(() => {
        return fillData(batch, { provinceId: 41, target: "HIGH SCHOOL" }); // Ninh Bình
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'O');
      }).then(() => {
        return fillData(batch, { provinceId: 7, target: "HIGH SCHOOL" }); // Bắc Ninh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'P');
      }).then(() => {
        return fillData(batch, { provinceId: 57, target: "HIGH SCHOOL" }); // Bắc Giang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'Q');
      }).then(() => {
        return fillData(batch, { provinceId: 62, target: "HIGH SCHOOL" }); // Vĩnh Phúc
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'R');
      }).then(() => {
        return fillData(batch, { provinceId: 44, target: "HIGH SCHOOL" }); // Phú Thọ
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'S');
      }).then(() => {
        return fillData(batch, { provinceId: 55, target: "HIGH SCHOOL" }); // Thanh Hóa
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'T');
      }).then(() => {
        return fillData(batch, { provinceId: 40, target: "HIGH SCHOOL" }); // Nghệ An
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'U');
      }).then(() => {
        return fillData(batch, { provinceId: 29, target: "HIGH SCHOOL" }); // Hà Tĩnh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'V');
      }).then(() => {
        return fillData(batch, { provinceId: 59, target: "HIGH SCHOOL" }); // Huế
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'W');
      }).then(() => {
        return fillData(batch, { provinceId: 50, target: "HIGH SCHOOL" }); // Quảng Trị
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'X');
      }).then(() => {
        return fillData(batch, { provinceId: 46, target: "HIGH SCHOOL" }); // Quảng Bình
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'Y');
      }).then(() => {
        return fillData(batch, { provinceId: 49, target: "HIGH SCHOOL" }); // Quảng Nam
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'Z');
      }).then(() => {
        return fillData(batch, { provinceId: 47, target: "HIGH SCHOOL" }); // Quảng Ngãi
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AA');
      }).then(() => {
        return fillData(batch, { provinceId: 45, target: "HIGH SCHOOL" }); // Phú Yên
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AB');
      }).then(() => {
        return fillData(batch, { provinceId: 4, target: "HIGH SCHOOL" }); // Bình Định
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AC');
      }).then(() => {
        return fillData(batch, { provinceId: 43, target: "HIGH SCHOOL" }); // Ninh Thuận
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AD');
      }).then(() => {
        return fillData(batch, { provinceId: 38, target: "HIGH SCHOOL" }); // Lâm Đồng
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AE');
      }).then(() => {
        return fillData(batch, { provinceId: 15, target: "HIGH SCHOOL" }); // Dak Lak
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AF');
      }).then(() => {
        return fillData(batch, { provinceId: 20, target: "HIGH SCHOOL" }); // Gia Lai
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AG');
      }).then(() => {
        return fillData(batch, { provinceId: 63, target: "HIGH SCHOOL" }); // Bà Rịa - Vũng Tàu
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AH');
      }).then(() => {
        return fillData(batch, { provinceId: 6, target: "HIGH SCHOOL" }); // Bạc Liêu
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AI');
      }).then(() => {
        return fillData(batch, { provinceId: 52, target: "HIGH SCHOOL" }); // Sóc Trăng
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AJ');
      }).then(() => {
        return fillData(batch, { provinceId: 8, target: "HIGH SCHOOL" }); // Bình Phước
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AK');
      }).then(() => {
        return fillData(batch, { provinceId: 18, target: "HIGH SCHOOL" }); // Dak Nông
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AL');
      }).then(() => {
        return fillData(batch, { provinceId: 12, target: "HIGH SCHOOL" }); // Cà Mau
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AM');
      }).then(() => {
        return fillData(batch, { provinceId: 19, target: "HIGH SCHOOL" }); // Đồng Tháp
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AN');
      }).then(() => {
        return fillData(batch, { provinceId: 26, target: "HIGH SCHOOL" }); // Hậu Giang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AO');
      }).then(() => {
        return fillData(batch, { provinceId: 23, target: "UNIVERSITY" }); // Ho Chi Minh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AP');
      }).then(() => {
        return fillData(batch, { provinceId: 21, target: "UNIVERSITY" }); // Ha Noi
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AQ');
      }).then(() => {
        return fillData(batch, { provinceId: 28, target: "UNIVERSITY" }); // Hai Phong
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AR');
      }).then(() => {
        return fillData(batch, { provinceId: 16, target: "UNIVERSITY" }); // Da Nang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AS');
      }).then(() => {
        return fillData(batch, { provinceId: 57, target: "UNIVERSITY" }); // Thái Nguyên
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AT');
      }).then(() => {
        return fillData(batch, { provinceId: 48, target: "UNIVERSITY" }); // Quảng Ninh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AU');
      }).then(() => {
        return fillData(batch, { provinceId: 24, target: "UNIVERSITY" }); // Hải Dương
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AV');
      }).then(() => {
        return fillData(batch, { provinceId: 31, target: "UNIVERSITY" }); // Hưng Yên
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AW');
      }).then(() => {
        return fillData(batch, { provinceId: 42, target: "UNIVERSITY" }); // Nam Định
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AX');
      }).then(() => {
        return fillData(batch, { provinceId: 53, target: "UNIVERSITY" }); // Thái Bình
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AY');
      }).then(() => {
        return fillData(batch, { provinceId: 41, target: "UNIVERSITY" }); // Ninh Bình
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'AZ');
      }).then(() => {
        return fillData(batch, { provinceId: 7, target: "UNIVERSITY" }); // Bắc Ninh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BA');
      }).then(() => {
        return fillData(batch, { provinceId: 57, target: "UNIVERSITY" }); // Bắc Giang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BB');
      }).then(() => {
        return fillData(batch, { provinceId: 62, target: "UNIVERSITY" }); // Vĩnh Phúc
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BC');
      }).then(() => {
        return fillData(batch, { provinceId: 44, target: "UNIVERSITY" }); // Phú Thọ
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BD');
      }).then(() => {
        return fillData(batch, { provinceId: 55, target: "UNIVERSITY" }); // Thanh Hóa
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BE');
      }).then(() => {
        return fillData(batch, { provinceId: 40, target: "UNIVERSITY" }); // Nghệ An
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BF');
      }).then(() => {
        return fillData(batch, { provinceId: 29, target: "UNIVERSITY" }); // Hà Tĩnh
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BG');
      }).then(() => {
        return fillData(batch, { provinceId: 59, target: "UNIVERSITY" }); // Huế
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BH');
      }).then(() => {
        return fillData(batch, { provinceId: 50, target: "UNIVERSITY" }); // Quảng Trị
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BI');
      }).then(() => {
        return fillData(batch, { provinceId: 46, target: "UNIVERSITY" }); // Quảng Bình
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BJ');
      }).then(() => {
        return fillData(batch, { provinceId: 49, target: "UNIVERSITY" }); // Quảng Nam
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BK');
      }).then(() => {
        return fillData(batch, { provinceId: 47, target: "UNIVERSITY" }); // Quảng Ngãi
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BL');
      }).then(() => {
        return fillData(batch, { provinceId: 45, target: "UNIVERSITY" }); // Phú Yên
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BM');
      }).then(() => {
        return fillData(batch, { provinceId: 4, target: "UNIVERSITY" }); // Bình Định
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BN');
      }).then(() => {
        return fillData(batch, { provinceId: 43, target: "UNIVERSITY" }); // Ninh Thuận
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BO');
      }).then(() => {
        return fillData(batch, { provinceId: 38, target: "UNIVERSITY" }); // Lâm Đồng
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BP');
      }).then(() => {
        return fillData(batch, { provinceId: 15, target: "UNIVERSITY" }); // Dak Lak
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BQ');
      }).then(() => {
        return fillData(batch, { provinceId: 20, target: "UNIVERSITY" }); // Gia Lai
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BR');
      }).then(() => {
        return fillData(batch, { provinceId: 63, target: "UNIVERSITY" }); // Bà Rịa - Vũng Tàu
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BS');
      }).then(() => {
        return fillData(batch, { provinceId: 6, target: "UNIVERSITY" }); // Bạc Liêu
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BT');
      }).then(() => {
        return fillData(batch, { provinceId: 52, target: "UNIVERSITY" }); // Sóc Trăng
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BU');
      }).then(() => {
        return fillData(batch, { provinceId: 8, target: "UNIVERSITY" }); // Bình Phước
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BV');
      }).then(() => {
        return fillData(batch, { provinceId: 18, target: "UNIVERSITY" }); // Dak Nông
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BW');
      }).then(() => {
        return fillData(batch, { provinceId: 12, target: "UNIVERSITY" }); // Cà Mau
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BX');
      }).then(() => {
        return fillData(batch, { provinceId: 19, target: "UNIVERSITY" }); // Đồng Tháp
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BY');
      }).then(() => {
        return fillData(batch, { provinceId: 26, target: "UNIVERSITY" }); // Hậu Giang
      }).then((rowData) => {
        return writeToTemplate(reportFilePath, rowData, 'BZ');
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
    worksheet.getColumn('O').width = 30;
    worksheet.getColumn('P').width = 30;
    worksheet.getColumn('Q').width = 30;
    worksheet.getColumn('R').width = 30;
    worksheet.getColumn('S').width = 30;
    worksheet.getColumn('T').width = 30;
    worksheet.getColumn('U').width = 30;
    worksheet.getColumn('V').width = 30;
    worksheet.getColumn('W').width = 30;
    worksheet.getColumn('X').width = 30;
    worksheet.getColumn('Y').width = 30;
    worksheet.getColumn('Z').width = 30;
    worksheet.getColumn('AA').width = 30;
    worksheet.getColumn('AB').width = 30;
    worksheet.getColumn('AC').width = 30;
    worksheet.getColumn('AD').width = 30;
    worksheet.getColumn('AE').width = 30;
    worksheet.getColumn('AF').width = 30;
    worksheet.getColumn('AG').width = 30;
    worksheet.getColumn('AH').width = 30;
    worksheet.getColumn('AI').width = 30;
    worksheet.getColumn('AJ').width = 30;
    worksheet.getColumn('AK').width = 30;
    worksheet.getColumn('AL').width = 30;
    worksheet.getColumn('AM').width = 30;
    worksheet.getColumn('AN').width = 30;
    worksheet.getColumn('AO').width = 30;
    worksheet.getColumn('AP').width = 30;
    worksheet.getColumn('AQ').width = 30;
    worksheet.getColumn('AR').width = 30;
    worksheet.getColumn('AS').width = 30;
    worksheet.getColumn('AT').width = 30;
    worksheet.getColumn('AU').width = 30;
    worksheet.getColumn('AV').width = 30;
    worksheet.getColumn('AW').width = 30;
    worksheet.getColumn('AX').width = 30;
    worksheet.getColumn('AY').width = 30;
    worksheet.getColumn('AZ').width = 30;
    worksheet.getColumn('BA').width = 30;
    worksheet.getColumn('BB').width = 30;
    worksheet.getColumn('BC').width = 30;
    worksheet.getColumn('BD').width = 30;
    worksheet.getColumn('BE').width = 30;
    worksheet.getColumn('BF').width = 30;
    worksheet.getColumn('BG').width = 30;
    worksheet.getColumn('BH').width = 30;
    worksheet.getColumn('BI').width = 30;
    worksheet.getColumn('BJ').width = 30;
    worksheet.getColumn('BK').width = 30;
    worksheet.getColumn('BL').width = 30;
    worksheet.getColumn('BM').width = 30;
    worksheet.getColumn('BN').width = 30;
    worksheet.getColumn('BO').width = 30;
    worksheet.getColumn('BP').width = 30;
    worksheet.getColumn('BQ').width = 30;
    worksheet.getColumn('BR').width = 30;
    worksheet.getColumn('BS').width = 30;
    worksheet.getColumn('BT').width = 30;
    worksheet.getColumn('BU').width = 30;
    worksheet.getColumn('BV').width = 30;
    worksheet.getColumn('BW').width = 30;
    worksheet.getColumn('BX').width = 30;
    worksheet.getColumn('BY').width = 30;
    worksheet.getColumn('BZ').width = 30;

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
    buildReportRow4(worksheet, 'E', 'E4:AO4', 'Break-down by city (High School)');
    buildReportRow4(worksheet, 'AP', 'AP4:BZ4', 'Break-down by city (University)');

    // // Row 5, B4-T4
    buildReportRow5(worksheet, 'B', 'Total Project');
    buildReportRow5(worksheet, 'C', 'Total ' + batch + ' (High school)');
    buildReportRow5(worksheet, 'D', 'Total ' + batch + ' (University)');

    buildReportRow5(worksheet, 'E', 'Hồ Chí Minh');
    buildReportRow5(worksheet, 'F', 'Hà Nội');
    buildReportRow5(worksheet, 'G', 'Hải Phòng');
    buildReportRow5(worksheet, 'H', 'Đà Nẵng');
    buildReportRow5(worksheet, 'I', 'Thái Nguyên');
    buildReportRow5(worksheet, 'J', 'Quảng Ninh');
    buildReportRow5(worksheet, 'K', 'Hải Dương');
    buildReportRow5(worksheet, 'L', 'Hưng Yên');
    buildReportRow5(worksheet, 'M', 'Nam Định');
    buildReportRow5(worksheet, 'N', 'Thái Bình');
    buildReportRow5(worksheet, 'O', 'Ninh Bình');
    buildReportRow5(worksheet, 'P', 'Bắc Ninh');
    buildReportRow5(worksheet, 'Q', 'Bắc Giang');
    buildReportRow5(worksheet, 'R', 'Vĩnh Phúc');
    buildReportRow5(worksheet, 'S', 'Phú Thọ');
    buildReportRow5(worksheet, 'T', 'Thanh Hóa');
    buildReportRow5(worksheet, 'U', 'Nghệ An');
    buildReportRow5(worksheet, 'V', 'Hà Tĩnh');
    buildReportRow5(worksheet, 'W', 'Huế');
    buildReportRow5(worksheet, 'X', 'Quảng Trị');
    buildReportRow5(worksheet, 'Y', 'Quảng Bình');
    buildReportRow5(worksheet, 'Z', 'Quảng Nam');
    buildReportRow5(worksheet, 'AA', 'Quảng Ngãi');
    buildReportRow5(worksheet, 'AB', 'Phú Yên');
    buildReportRow5(worksheet, 'AC', 'Bình Định');
    buildReportRow5(worksheet, 'AD', 'Ninh Thuận');
    buildReportRow5(worksheet, 'AE', 'Lâm Đồng');
    buildReportRow5(worksheet, 'DF', 'Đắk Lắk');
    buildReportRow5(worksheet, 'AG', 'Gia Lai');
    buildReportRow5(worksheet, 'AH', 'Bà Rịa - Vũng Tàu');
    buildReportRow5(worksheet, 'AI', 'Bạc Liêu');
    buildReportRow5(worksheet, 'AJ', 'Sóc Trăng');
    buildReportRow5(worksheet, 'AK', 'Bình Phước');
    buildReportRow5(worksheet, 'AL', 'Đắk Nông');
    buildReportRow5(worksheet, 'AM', 'Cà Mau');
    buildReportRow5(worksheet, 'AN', 'Đồng Tháp');
    buildReportRow5(worksheet, 'AO', 'Hậu Giang');

    buildReportRow5(worksheet, 'AP', 'Hồ Chí Minh');
    buildReportRow5(worksheet, 'AQ', 'Hà Nội');
    buildReportRow5(worksheet, 'AR', 'Hải Phòng');
    buildReportRow5(worksheet, 'AS', 'Đà Nẵng');
    buildReportRow5(worksheet, 'AT', 'Thái Nguyên');
    buildReportRow5(worksheet, 'AU', 'Quảng Ninh');
    buildReportRow5(worksheet, 'AV', 'Hải Dương');
    buildReportRow5(worksheet, 'AW', 'Hưng Yên');
    buildReportRow5(worksheet, 'AX', 'Nam Định');
    buildReportRow5(worksheet, 'AY', 'Thái Bình');
    buildReportRow5(worksheet, 'AZ', 'Ninh Bình');
    buildReportRow5(worksheet, 'BA', 'Bắc Ninh');
    buildReportRow5(worksheet, 'BB', 'Bắc Giang');
    buildReportRow5(worksheet, 'BC', 'Vĩnh Phúc');
    buildReportRow5(worksheet, 'BD', 'Phú Thọ');
    buildReportRow5(worksheet, 'BE', 'Thanh Hóa');
    buildReportRow5(worksheet, 'BF', 'Nghệ An');
    buildReportRow5(worksheet, 'BG', 'Hà Tĩnh');
    buildReportRow5(worksheet, 'BH', 'Huế');
    buildReportRow5(worksheet, 'BI', 'Quảng Trị');
    buildReportRow5(worksheet, 'BJ', 'Quảng Bình');
    buildReportRow5(worksheet, 'BK', 'Quảng Nam');
    buildReportRow5(worksheet, 'BL', 'Quảng Ngãi');
    buildReportRow5(worksheet, 'BM', 'Phú Yên');
    buildReportRow5(worksheet, 'BN', 'Bình Định');
    buildReportRow5(worksheet, 'BO', 'Ninh Thuận');
    buildReportRow5(worksheet, 'BP', 'Lâm Đồng');
    buildReportRow5(worksheet, 'BQ', 'Đắk Lắk');
    buildReportRow5(worksheet, 'BR', 'Gia Lai');
    buildReportRow5(worksheet, 'BS', 'Bà Rịa - Vũng Tàu');
    buildReportRow5(worksheet, 'BT', 'Bạc Liêu');
    buildReportRow5(worksheet, 'BU', 'Sóc Trăng');
    buildReportRow5(worksheet, 'BV', 'Bình Phước');
    buildReportRow5(worksheet, 'BW', 'Đắk Nông');
    buildReportRow5(worksheet, 'BX', 'Cà Mau');
    buildReportRow5(worksheet, 'BY', 'Đồng Tháp');
    buildReportRow5(worksheet, 'BZ', 'Hậu Giang');

    // Data
    let colArr = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL','AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ'];

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

  if (cellIndex == 'C' || cellIndex == 'D') {
    fgColor = { theme: 5, tint: 0.5999938962981048 };
  }

  if (['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ'].includes(cellIndex)) {
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

  if (rowIndex == 6 || rowIndex == 7 || rowIndex == 15 || rowIndex == 19 || rowIndex == 23) {
    bold = true;
  }

  if (rowIndex == 7 || rowIndex == 15 || rowIndex == 19) {
    row.getCell(cellIndex).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00B0F0' },
      bgColor: { indexed: 64 }
    };
  }

  if (rowIndex == 23) {
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

  if (rowIndex == 7 || rowIndex == 15 || rowIndex == 19 || rowIndex == 23) {
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
    coalesce(SUM(MissingLivingCity),0) AS MissingLivingCity,\
    coalesce(SUM(missingContactInformation),0) AS MissingContactInformation, \
    coalesce(SUM(missingAge),0) As MissingAge, \
    coalesce(SUM(missingSchoolName),0) AS MissingSchoolName, \
    coalesce(SUM(missingCollectedDate),0) AS MissingCollectedDate, \
    coalesce(SUM(missingBrandUsing),0) AS MissingBrandUsing, \
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

    if (filterType.target !== undefined && filterType.target !== null) {
      if (whereCondition === '') {
        whereCondition = 'WHERE customers.target = $target'
      } else {
        whereCondition += " AND customers.target = $target";
      }

      params = _.merge(params, {
        $target: filterType.target
      });
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
      row.getCell(cellIndex).value = rowData.DuplicatedPhoneBetweenPupilAndStudent;

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
