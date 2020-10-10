const Excel = require('exceljs')
const fs = require('fs');

import _ from 'lodash'
import { db } from './database';
import { createCustomer } from './createCustomer'
import { buildExcelTemplate } from './buildExcelTemplate'

// BrandMax - High School
// Focus MKT - University

const dataBeginRow = 2
const indexCol = 1
const schoolNameCol = 2
const provinceNameCol = 3
const districtNameCol = 4
const collectedDateCol = 5
const collectedTimeCol = 6
const firstNameCol = 7
const lastNameCol = 8
const phoneNumberCol = 9
const parentPhoneNumberCol = 10
const dateOfBirthCol = 11
const brandCol = 12
const subBrandCol = 13
const samplingProductCol = 14
const genderCol = 15
const districtIdCol = 16
const provinceIdCol = 17
const optInCol = 18

const isEmptyRow = (row) => {
  if (row.getCell(indexCol).value === null     &&
      row.getCell(schoolNameCol).value === null      &&
      row.getCell(provinceNameCol).value === null      &&
      row.getCell(districtNameCol).value === null      &&
      row.getCell(phoneNumberCol).value === null         &&
      row.getCell(firstNameCol).value === null           &&
      row.getCell(lastNameCol).value === null           &&
      row.getCell(brandCol).value === null           &&
      row.getCell(dateOfBirthCol).value === null
    ) {
    // Empty Row
    return true
  }

  return false
}

export const importData = (excelFile, batch, source, outputDirectory) => {
  return new Promise((resolve, reject) => {
    if ( !_.endsWith(outputDirectory, '/') ) {
      outputDirectory += '/';
    }

    let dir = outputDirectory + batch;

    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir)
    }

    dir = dir + '/';

    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir)
    }

    resolve(readFile(excelFile, batch, source, dir));
  });
}

const readFile = (excelFile, batch, source, outputDirectory) => {
  return new Promise((resolve, reject) => {
    let workbook = new Excel.Workbook();
    workbook.xlsx.readFile(excelFile).then(() => {
      let worksheet = workbook.getWorksheet(1);
      let rowNumber = dataBeginRow;
      let outputPath = outputDirectory + '/' + batch + '_' + source.replace(/ /g, '_') + '_cleaned_data.xlsx';

      if (fs.existsSync(outputPath)) {
        fs.unlinkSync(outputPath);
      }

      buildExcelTemplate(outputPath).then((outputWorkbook) => {
        return readEachRow(excelFile, outputWorkbook, batch, source, worksheet, rowNumber);
      }).then((outputWorkbook) => {
        resolve(outputWorkbook.xlsx.writeFile(outputPath));
      });
    })
  })
}

const readEachRow = (excelFile, outputWorkbook, batch, source, worksheet, rowNumber) => {
  return new Promise((resolve, reject) => {
    let row = worksheet.getRow(rowNumber);
    console.log('Row: ' + rowNumber);

    if (isEmptyRow(row)) {
      return resolve(outputWorkbook);
    }

    let dateOfBirth = row.getCell(dateOfBirthCol).value
    let currentYear = new Date().getFullYear()
    let age;
    if (yearOfBirth) {
      age = currentYear - parseInt(yearOfBirth)
    }

    let customer = {
      areaName: row.getCell(areaCol).value,
      provinceName: row.getCell(provinceCol).value,
      schoolName: row.getCell(schoolNameCol).value,
      name: row.getCell(nameCol).value,
      yearOfBirth: yearOfBirth,
      age: age,
      phoneNumber: row.getCell(phoneNumberCol).value,
      parentPhoneNumber: row.getCell(parentPhoneNumberCol).value,
      facebook: row.getCell(facebookCol).value,
      email: row.getCell(emailCol).value,
      kotexData: row.getCell(kotexCol).value,
      dianaData: row.getCell(dianaCol).value,
      laurierData: row.getCell(laurierCol).value,
      whisperData: row.getCell(whisperCol).value,
      othersData: row.getCell(othersCol).value,
      notes: row.getCell(notesCol).value,
      createdAt: row.getCell(createdAtCol).value,
      receivedGift: row.getCell(receivedGiftCol).value,
      groupName: row.getCell(groupNameCol).value,
      batch: batch
    }

    customer

    createCustomer(customer).then((response) => {
      customer = response;
      let missingData = customer.missingData === 1;
      let illogicalData = customer.illogicalData === 1;
      let duplicateData = customer.duplicatedPhone === 1;

      let rowData = [
        customer.customerId,
        customer.areaName,
        customer.provinceName,
        customer.schoolName,
        customer.name,
        customer.yearOfBirth,
        customer.phoneNumber,
        customer.parentPhoneNumber,
        customer.facebook,
        customer.email,
        customer.kotexData,
        customer.dianaData,
        customer.laurierData,
        customer.whisperData,
        customer.othersData,
        customer.notes,
        customer.createdAt,
        customer.receivedGift,
        customer.groupName
      ];

      let outputSheetName = 'Valid';
      if (missingData || illogicalData) {
        outputSheetName = 'Invalid';
      } else if (duplicateData === true) {
        outputSheetName = 'Duplication';
      }


      if (duplicateData == true) {
        var duplicatedWith
        duplicatedWith = customer.duplicatedWith;

        var duplicatedRow = [
          duplicatedWith.customerId,
          duplicatedWith.areaName,
          duplicatedWith.provinceName,
          duplicatedWith.schoolName,
          duplicatedWith.name,
          duplicatedWith.yearOfBirth,
          duplicatedWith.phoneNumber,
          duplicatedWith.parentPhoneNumber,
          duplicatedWith.facebook,
          duplicatedWith.email,
          duplicatedWith.kotexData,
          duplicatedWith.dianaData,
          duplicatedWith.laurierData,
          duplicatedWith.whisperData,
          duplicatedWith.othersData,
          duplicatedWith.notes,
          duplicatedWith.createdAt,
          duplicatedWith.receivedGift,
          duplicatedWith.groupName,
          duplicatedWith.batch
        ]

        rowData.push(customer.batch);

        writeToFile(outputWorkbook, outputSheetName, duplicatedRow).then((workbook) => {
          writeToFile(outputWorkbook, outputSheetName, rowData).then((workbook) => {
            if (rowNumber % 1000 === 0) {
              setTimeout(function(){
                resolve(readEachRow(excelFile, workbook, batch, source, worksheet, rowNumber+1));
              }, 0);
            } else {
              resolve(readEachRow(excelFile, workbook, batch, source, worksheet, rowNumber+1));
            }
          });
        });
      } else {
        writeToFile(outputWorkbook, outputSheetName, rowData).then((workbook) => {
          if (rowNumber % 1000 === 0) {
              setTimeout(function(){
                resolve(readEachRow(excelFile, workbook, batch, source, worksheet, rowNumber+1));
              }, 0);
            } else {
              resolve(readEachRow(excelFile, workbook, batch, source, worksheet, rowNumber+1));
            }
        });
      }
    });
  });
}

export const writeToFile = (outputWorkbook, outputSheetName, rowData) => {
  return new Promise((resolve, reject) => {
    let workbook = outputWorkbook;
    let worksheet = workbook.getWorksheet(outputSheetName);
    let row = worksheet.addRow(rowData);

    row.getCell(1).font = {
      size: 10, color: { theme: 1 }, name: 'Arial', family: 2
    };

    row.getCell(1).border = worksheet.getCell('A5').border;
    row.getCell(1).alignment = worksheet.getCell('A5').alignment;

    row.getCell(2).font = row.getCell(1).font;
    row.getCell(2).border = row.getCell(1).border;
    row.getCell(2).alignment = row.getCell(1).alignment;

    row.getCell(3).font = row.getCell(1).font;
    row.getCell(3).border = row.getCell(1).border;
    row.getCell(3).alignment = row.getCell(1).alignment;

    row.getCell(4).font = row.getCell(1).font;
    row.getCell(4).border = row.getCell(1).border;
    row.getCell(4).alignment = row.getCell(1).alignment;

    row.getCell(5).font = row.getCell(1).font;
    row.getCell(5).border = row.getCell(1).border;
    row.getCell(5).alignment = row.getCell(1).alignment;

    row.getCell(6).font = row.getCell(1).font;
    row.getCell(6).border = row.getCell(1).border;
    row.getCell(6).alignment = row.getCell(1).alignment;

    row.getCell(7).font = row.getCell(1).font;
    row.getCell(7).border = row.getCell(1).border;
    row.getCell(7).alignment = row.getCell(1).alignment;

    row.getCell(8).font = row.getCell(1).font;
    row.getCell(8).border = row.getCell(1).border;
    row.getCell(8).alignment = row.getCell(1).alignment;

    row.getCell(9).font = row.getCell(1).font;
    row.getCell(9).border = row.getCell(1).border;
    row.getCell(9).alignment = row.getCell(1).alignment;

    row.getCell(10).font = row.getCell(1).font;
    row.getCell(10).border = row.getCell(1).border;
    row.getCell(10).alignment = row.getCell(1).alignment;

    row.getCell(11).font = row.getCell(1).font;
    row.getCell(11).border = row.getCell(1).border;
    row.getCell(11).alignment = row.getCell(1).alignment;

    row.getCell(12).font = row.getCell(1).font;
    row.getCell(12).border = row.getCell(1).border;
    row.getCell(12).alignment = row.getCell(1).alignment;

    row.getCell(13).font = row.getCell(1).font;
    row.getCell(13).border = row.getCell(1).border;
    row.getCell(13).alignment = row.getCell(1).alignment;

    row.getCell(14).font = row.getCell(1).font;
    row.getCell(14).border = row.getCell(1).border;
    row.getCell(14).alignment = row.getCell(1).alignment;

    row.getCell(15).font = row.getCell(1).font;
    row.getCell(15).border = row.getCell(1).border;
    row.getCell(15).alignment = row.getCell(1).alignment;

    row.getCell(16).font = row.getCell(1).font;
    row.getCell(16).border = row.getCell(1).border;
    row.getCell(16).alignment = row.getCell(1).alignment;

    row.getCell(17).font = row.getCell(1).font;
    row.getCell(17).border = row.getCell(1).border;
    row.getCell(17).alignment = row.getCell(1).alignment;

    row.getCell(18).font = row.getCell(1).font;
    row.getCell(18).border = row.getCell(1).border;
    row.getCell(18).alignment = row.getCell(1).alignment;

    row.getCell(19).font = row.getCell(1).font;
    row.getCell(19).border = row.getCell(1).border;
    row.getCell(19).alignment = row.getCell(1).alignment;

    if (outputSheetName.endsWith('Duplication')) {
      row.getCell(20).font = row.getCell(1).font;
      row.getCell(20).border = row.getCell(1).border;
      row.getCell(20).alignment = row.getCell(1).alignment;
    }

    resolve(workbook);
  });
}
