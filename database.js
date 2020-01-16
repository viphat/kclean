const electron = require('electron');
const dialog = electron.dialog;
const sqlite3 = require('sqlite3').verbose();
export const db = new sqlite3.Database('db.sqlite3');

export const setupDatabase = () => {
  db.serialize(()=>{
    createTableAreas();
    createTableProvinces();
    createTableCustomers();
  });

  dialog.showMessageBox({type: 'info', title: 'Thông báo', message: 'Đã khởi tạo Database thành công, bạn có thể tiếp tục sử dụng ứng dụng.'});
}

export const importData = () => {
  db.serialize(()=>{
    insertTableAreas();
    insertTableProvinces();
  });

  dialog.showMessageBox({type: 'info', title: 'Thông báo', message: 'Import dữ liệu thành công.'});
}

function createTableAreas() {
  db.run('DROP TABLE IF EXISTS areas;');
  db.run('CREATE TABLE areas(areaId INTEGER PRIMARY KEY, name TEXT);');
}

function createTableProvinces() {
  db.run('DROP TABLE IF EXISTS provinces;');
  db.run('CREATE TABLE provinces(provinceId INTEGER PRIMARY KEY, name TEXT, areaId INTEGER, FOREIGN KEY(areaId) REFERENCES areas(areaId));');
}

function createTableCustomers() {
  db.run('DROP TABLE IF EXISTS customers;');
  db.run('CREATE TABLE customers(customerId INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, provinceId INTEGER, areaName TEXT, provinceName TEXT, schoolName TEXT, age INTEGER, phoneNumber TEXT, parentPhoneNumber TEXT, facebook TEXT, email TEXT, contactInformation TEXT, kotexData TEXT, dianaData TEXT, laurierData TEXT, whisperData TEXT, othersData TEXT, createdAt TEXT, notes TEXT, receivedGift TEXT, groupName TEXT, groupId INTEGER, batch TEXT,\
    hasError INTEGER DEFAULT 0,\
    missingData INTEGER DEFAULT 0,\
    missingName INTEGER DEFAULT 0,\
    missingLivingCity INTEGER DEFAULT 0,\
    missingContactInformation INTEGER DEFAULT 0,\
    missingAge INTEGER DEFAULT 0,\
    missingSchoolName INTEGER DEFAULT 0,\
    missingBrandUsing INTEGER DEFAULT 0,\
    missingGroup INTEGER DEFAULT 0,\
    illogicalData INTEGER DEFAULT 0,\
    illogicalPhone INTEGER DEFAULT 0,\
    illogicalAge INTEGER DEFAULT 0,\
    illogicalAgePupil INTEGER DEFAULT 0,\
    illogicalAgeStudent INTEGER DEFAULT 0,\
    illogicalAgeOthers INTEGER DEFAULT 0,\
    duplicatedPhone INTEGER DEFAULT 0,\
    duplicatedPhoneBetweenPupilAndStudent INTEGER DEFAULT 0,\
    duplicatedPhoneBetweenPupilAndOthers INTEGER DEFAULT 0,\
    duplicatedPhoneBetweenStudentAndOthers INTEGER DEFAULT 0,\
    duplicatedPhoneWithinPupil INTEGER DEFAULT 0,\
    duplicatedPhoneWithinStudent INTEGER DEFAULT 0,\
    duplicatedPhoneWithinOthers INTEGER DEFAULT 0,\
    FOREIGN KEY(provinceId) REFERENCES provinces(provinceId));');
}

function insertTableAreas() {
  db.run('INSERT INTO areas(areaId, name) VALUES(?, ?);', 1, 'Hồ Chí Minh');
  db.run('INSERT INTO areas(areaId, name) VALUES(?, ?);', 2, 'Tây Nam Bộ');
  db.run('INSERT INTO areas(areaId, name) VALUES(?, ?);', 3, 'Đông Nam Bộ');
  db.run('INSERT INTO areas(areaId, name) VALUES(?, ?);', 4, 'Miền Trung');
}

function insertTableProvinces() {
  db.run('INSERT INTO provinces(provinceId, areaId, name) VALUES(?, ?, ?);', 1, 1, 'Hồ Chí Minh');
  db.run('INSERT INTO provinces(provinceId, areaId, name) VALUES(?, ?, ?);', 2, 2, 'Cần Thơ');
  db.run('INSERT INTO provinces(provinceId, areaId, name) VALUES(?, ?, ?);', 3, 2, 'Vĩnh Long');
  db.run('INSERT INTO provinces(provinceId, areaId, name) VALUES(?, ?, ?);', 4, 3, 'Đồng Nai');
  db.run('INSERT INTO provinces(provinceId, areaId, name) VALUES(?, ?, ?);', 5, 4, 'Đà Nẵng');
  db.run('INSERT INTO provinces(provinceId, areaId, name) VALUES(?, ?, ?);', 6, 4, 'Huế');
}
