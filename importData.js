const Excel = require('exceljs')
const fs = require('fs');

import _ from 'lodash'
import { db } from './database';
import { createCustomer } from './createCustomer'
import { buildExcelTemplate } from './buildExcelTemplate'

const dataBeginRow = 2
const indexCol = 1
const nameCol = 2
const phoneNumberCol = 3
const addressCol = 4
const cityCol = 5
const modelCol = 6

const isEmptyRow = (row) => {
  if (row.getCell(indexCol).value === null     &&
      row.getCell(nameCol).value === null      &&
      row.getCell(phoneNumberCol).value === null         &&
      row.getCell(addressCol).value === null           &&
      row.getCell(cityCol).value === null            &&
      row.getCell(modelCol).value === null) {
    // Empty Row
    return true
  }

  return false
}

export const importData = (excelFile, batch, outputDirectory) => {
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

    resolve(readFile(excelFile, batch, dir));
  });
}

const readFile = (excelFile, batch, outputDirectory) => {
  return new Promise((resolve, reject) => {
    let workbook = new Excel.Workbook();
    workbook.xlsx.readFile(excelFile).then(() => {
      let worksheet = workbook.getWorksheet(1);
      let rowNumber = dataBeginRow;
      let outputPath = outputDirectory + '/' + batch + '_cleaned_data.xlsx';

      if (fs.existsSync(outputPath)) {
        fs.unlinkSync(outputPath);
      }

      buildExcelTemplate(outputPath).then((outputWorkbook) => {
        return readEachRow(excelFile, outputWorkbook, batch, worksheet, rowNumber);
      }).then((outputWorkbook) => {
        resolve(outputWorkbook.xlsx.writeFile(outputPath));
      });
    })
  })
}

const readEachRow = (excelFile, outputWorkbook, batch, worksheet, rowNumber) => {
  return new Promise((resolve, reject) => {
    let row = worksheet.getRow(rowNumber);
    console.log('Row: ' + rowNumber);

    if (isEmptyRow(row)) {
      return resolve(outputWorkbook);
    }

    let customer = {
      name: row.getCell(nameCol).value,
      phoneNumber: row.getCell(phoneNumberCol).value,
      address: row.getCell(addressCol).value,
      city: row.getCell(cityCol).value,
      model: row.getCell(modelCol).value,
      batch: batch
    }

    createCustomer(customer).then((response) => {
      customer = response;
      let missingData = customer.missingData === 1;
      let illogicalData = customer.illogicalPhone === 1;
      let duplicateData = customer.duplicatedPhone === 1;

      let rowData = [
        customer.customerId,
        customer.name,
        customer.phoneNumber,
        customer.address,
        customer.city,
        customer.model
      ];

      let outputSheetName = 'Valid';

      if (missingData || illogicalData) {
        if (customer.illogicalPhoneFormat === 1) {
          outputSheetName = 'Invalid - Phone Format';
        } else if (customer.illogicalPhoneProvider === 1) {
          outputSheetName = 'Invalid - Phone Provider';
        } else {
          outputSheetName = 'Invalid';
        }
      } else if (duplicateData === true) {
        if (customer.duplicatedPhoneSameModel === 1) {
          outputSheetName = 'Duplicated - Same Model';
        } else {
          outputSheetName = 'Duplicated - Different Model';
        }
      }

      if (duplicateData == true) {
        var duplicatedWith
        duplicatedWith = customer.duplicatedWith;

        var duplicatedRow = [
          duplicatedWith.customerId,
          duplicatedWith.name,
          duplicatedWith.phoneNumber,
          duplicatedWith.address,
          duplicatedWith.city,
          duplicatedWith.model,
          duplicatedWith.batch
        ]

        rowData.push(customer.batch);

        writeToFile(outputWorkbook, outputSheetName, duplicatedRow).then((workbook) => {
          writeToFile(outputWorkbook, outputSheetName, rowData).then((workbook) => {
            if (rowNumber % 1000 === 0) {
              setTimeout(function(){
                resolve(readEachRow(excelFile, workbook, batch, worksheet, rowNumber + 1));
              }, 0);
            } else {
              resolve(readEachRow(excelFile, workbook, batch, worksheet, rowNumber + 1));
            }
          });
        });
      } else {
        writeToFile(outputWorkbook, outputSheetName, rowData).then((workbook) => {
          if (rowNumber % 1000 === 0) {
              setTimeout(function(){
                resolve(readEachRow(excelFile, workbook, batch, worksheet, rowNumber+1));
              }, 0);
            } else {
              resolve(readEachRow(excelFile, workbook, batch, worksheet, rowNumber+1));
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

    if (outputSheetName.endsWith('Duplicated - Same Model') || outputSheetName.endsWith('Duplicated - Different Model')) {
      row.getCell(7).font = row.getCell(1).font;
      row.getCell(7).border = row.getCell(1).border;
      row.getCell(7).alignment = row.getCell(1).alignment;
    }

    resolve(workbook);
  });
}
