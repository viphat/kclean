const Excel = require('exceljs')
const fs = require('fs');
const _ = require('lodash');
const padStart = require('string.prototype.padstart');

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

    if (isEmptyRow(row)) {
      return resolve(outputWorkbook);
    }

    console.log('Row: ' + rowNumber);

    let dateOfBirth = row.getCell(dateOfBirthCol).value;
    let dayOfBirth, monthOfBirth, yearOfBirth, age;
    let arr;

    dateOfBirth = new Date(dateOfBirth)
    // Do Data gốc bị sai format, nên phải ép lại

    if (dateOfBirth !== null && dateOfBirth !== undefined) {
      if (dateOfBirth.toString() === 'Invalid Date') {
        // dd/mm/yyyy
        arr = row.getCell(dateOfBirthCol).value.toString().split('/')

        if (arr.length !== 3) {
          return reject('Lỗi ngày tháng DOB ở dòng ' + rowNumber)
        }

        dayOfBirth = padStart(arr[0], 2, 0);
        monthOfBirth = padStart(arr[1], 2, 0);
        yearOfBirth = arr[2].length === 2 ? '20' + arr[2] : arr[2];
      } else {
        monthOfBirth = dateOfBirth.getDate()
        dayOfBirth = dateOfBirth.getMonth() + 1
        yearOfBirth = dateOfBirth.getFullYear()

        if (monthOfBirth > 12) {
          return reject('Lỗi ngày tháng DOB ở dòng ' + rowNumber)
        }
      }

      dateOfBirth = new Date(yearOfBirth + '-' + monthOfBirth + '-' + dayOfBirth)

      let currentYear = new Date().getFullYear()

      if (yearOfBirth) {
        age = currentYear - parseInt(yearOfBirth)
      }
    }

    // dateOfBirth = new Date(dateOfBirth)
    // console.log(dateOfBirth)
    // if (dateOfBirth.toString() === 'Invalid Date') {
    //   return reject('Lỗi ngày tháng DOB ở dòng ' + rowNumber)
    // }

    let collectedDay, collectedMonth, collectedYear;
    let collectedDate = row.getCell(collectedDateCol).value

    if (collectedDate !== null && collectedDate !== undefined) {
      arr = collectedDate.toString().split('/')
      if (arr.length !== 3) {
        return reject('Lỗi ngày tháng cột E ở dòng ' + rowNumber)
      }

      collectedDay = padStart(arr[0], 2, 0);
      collectedMonth = padStart(arr[1], 2, 0);
      collectedYear = arr[2].length === 2 ? '20' + arr[2] : arr[2];
    }

    // collectedDate = new Date(collectedDate)
    // let collectedDay = collectedDate.getDate()
    // let collectedMonth = collectedDate.getMonth() + 1
    // let collectedYear = collectedDate.getFullYear()
    // if (collectedDate.toString() === 'Invalid Date') {
    //   return reject('Lỗi ngày tháng cột E ở dòng ' + rowNumber)
    // }

    let customer = {
      customerIndex: row.getCell(indexCol).value,
      firstName: row.getCell(firstNameCol).value,
      lastName: row.getCell(lastNameCol).value,
      provinceId: row.getCell(provinceIdCol).value,
      provinceName: row.getCell(provinceNameCol).value,
      districtId: row.getCell(districtIdCol).value,
      districtName: row.getCell(districtNameCol).value,
      schoolName: row.getCell(schoolNameCol).value,
      phoneNumber: row.getCell(phoneNumberCol).value,
      parentPhoneNumber: row.getCell(parentPhoneNumberCol).value,
      dateOfBirth: yearOfBirth ? (yearOfBirth + '-' + padStart(monthOfBirth, 2, 0) + '-' + padStart(dayOfBirth, 2, 0)) : null,
      yearOfBirth: yearOfBirth,
      age: age,
      collectedDate: collectedYear? (collectedYear + '-' + padStart(collectedMonth, 2, 0) + '-' + padStart(collectedDay, 2, 0)) : null,
      collectedTime: row.getCell(collectedTimeCol).value,
      brand: row.getCell(brandCol).value,
      subBrand: row.getCell(subBrandCol).value,
      samplingProduct: row.getCell(samplingProductCol).value,
      gender: row.getCell(genderCol).value,
      optIn: row.getCell(optInCol).value,
      source: source,
      batch: batch
    }

    createCustomer(customer).then((response) => {
      customer = response;
      let missingData = customer.missingData === 1;
      let illogicalData = customer.illogicalData === 1;
      let duplicateData = customer.duplicatedPhone === 1;

      let rowData = [
        customer.customerIndex,
        customer.schoolName,
        customer.provinceName,
        customer.districtName,
        customer.collectedDate,
        customer.collectedTime,
        customer.firstName,
        customer.lastName,
        customer.phoneNumber,
        customer.parentPhoneNumber,
        customer.dateOfBirth,
        customer.brand,
        customer.subBrand,
        customer.samplingProduct,
        customer.gender,
        customer.districtId,
        customer.provinceId,
        customer.optIn,
        customer.source
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
          duplicatedWith.customerIndex,
          duplicatedWith.schoolName,
          duplicatedWith.provinceName,
          duplicatedWith.districtName,
          duplicatedWith.collectedDate,
          duplicatedWith.collectedTime,
          duplicatedWith.firstName,
          duplicatedWith.lastName,
          duplicatedWith.phoneNumber,
          duplicatedWith.parentPhoneNumber,
          duplicatedWith.dateOfBirth,
          duplicatedWith.brand,
          duplicatedWith.subBrand,
          duplicatedWith.samplingProduct,
          duplicatedWith.gender,
          duplicatedWith.districtId,
          duplicatedWith.provinceId,
          duplicatedWith.optIn,
          duplicatedWith.source,
          duplicatedWith.batch,
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
