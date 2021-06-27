const electron = require('electron');
const dialog = electron.dialog;
const sqlite3 = require('sqlite3').verbose();
export const db = new sqlite3.Database('db.sqlite3');

export const setupDatabase = () => {
  db.serialize(()=>{
    createTableCustomers();
  });

  dialog.showMessageBox({type: 'info', title: 'Thông báo', message: 'Đã khởi tạo Database thành công, bạn có thể tiếp tục sử dụng ứng dụng.'});
}

function createTableCustomers() {
  db.run('DROP TABLE IF EXISTS customers;');
  db.run('CREATE TABLE customers(customerId INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, phoneNumber TEXT, address TEXT, city TEXT, model TEXT, batch TEXT,\
    hasError INTEGER DEFAULT 0,\
    missingData INTEGER DEFAULT 0,\
    missingName INTEGER DEFAULT 0,\
    missingPhoneNumber INTEGER DEFAULT 0,\
    missingAddress INTEGER DEFAULT 0,\
    missingModel INTEGER DEFAULT 0,\
    illogicalPhone INTEGER DEFAULT 0,\
    illogicalPhoneFormat INTEGER DEFAULT 0,\
    illogicalPhoneProvider INTEGER DEFAULT 0,\
    duplicatedPhone INTEGER DEFAULT 0,\
    duplicatedPhoneSameModel INTEGER DEFAULT 0,\
    duplicatedPhoneDiffModel INTEGER DEFAULT 0);');
}
