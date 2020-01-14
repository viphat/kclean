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
  db.run('CREATE TABLE areas(area_id INTEGER PRIMARY KEY, name TEXT);');
}

function createTableProvinces() {
  db.run('DROP TABLE IF EXISTS provinces;');
  db.run('CREATE TABLE provinces(province_id INTEGER PRIMARY KEY, name TEXT, area_id INTEGER, FOREIGN KEY(area_id) REFERENCES areas(area_id));');
}

function createTableCustomers() {
  db.run('DROP TABLE IF EXISTS customers;');
  db.run('CREATE TABLE customers(customer_id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, province_id INTEGER, school_name TEXT, age INTEGER, phone_number TEXT, parent_phone_number TEXT, facebook TEXT, email TEXT, kotex_col_data TEXT, diana_col_data TEXT, laurier_col_data TEXT, whisper_col_data TEXT, others_col_data TEXT, created_at TEXT, received_gift TEXT, group_id INTEGER,\
    missingData INTEGER DEFAULT 0,\
    missingName INTEGER DEFAULT 0,\
    missingSchoolName INTEGER DEFAULT 0,\
    missingAddress INTEGER DEFAULT 0,\
    missingPhone INTEGER DEFAULT 0,\
    missingAge INTEGER DEFAULT 0,\
    missingUsingBrand INTEGER DEFAULT 0,\
    missingGroup INTEGER DEFAULT 0,\
    illogicalData INTEGER DEFAULT 0,\
    illogicalPhone INTEGER DEFAULT 0,\
    illogicalAge INTEGER DEFAULT 0,\
    duplicatedPhone INTEGER DEFAULT 0,\
    duplicatedPhoneBetweenGroups INTEGER DEFAULT 0,\
    duplicatedPhoneWithinGroupPupil INTEGER DEFAULT 0,\
    duplicatedPhoneWithinGroupStudent INTEGER DEFAULT 0,\
    FOREIGN KEY(province_id) REFERENCES provinces(province_id));');
}

function insertTableAreas() {
  db.run('INSERT INTO areas(area_id, name) VALUES(?, ?);', 1, 'Hồ Chí Minh');
  db.run('INSERT INTO areas(area_id, name) VALUES(?, ?);', 2, 'Tây Nam Bộ');
  db.run('INSERT INTO areas(area_id, name) VALUES(?, ?);', 3, 'Đông Nam Bộ');
  db.run('INSERT INTO areas(area_id, name) VALUES(?, ?);', 4, 'Miền Trung');
}

function insertTableProvinces() {
  db.run('INSERT INTO provinces(province_id, area_id, name) VALUES(?, ?, ?);', 1, 1, 'Hồ Chí Minh');
  db.run('INSERT INTO provinces(province_id, area_id, name) VALUES(?, ?, ?);', 2, 2, 'Cần Thơ');
  db.run('INSERT INTO provinces(province_id, area_id, name) VALUES(?, ?, ?);', 3, 2, 'Vĩnh Long');
  db.run('INSERT INTO provinces(province_id, area_id, name) VALUES(?, ?, ?);', 4, 3, 'Đồng Nai');
  db.run('INSERT INTO provinces(province_id, area_id, name) VALUES(?, ?, ?);', 5, 4, 'Đà Nẵng');
  db.run('INSERT INTO provinces(province_id, area_id, name) VALUES(?, ?, ?);', 6, 4, 'Huế');
}
