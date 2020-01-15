import { db } from './database';
import { provinces } from './provinces';
import _ from 'lodash'

export const checkMissingData = (customer) => {
  customer.missingData = 0
  customer.missingLivingCity = 0
  customer.missingName = 0
  customer.missingContactInformation = 0
  customer.missingAge = 0
  customer.missingSchoolName = 0
  customer.missingBrandUsing = 0
  customer.missingGroup = 0

  if (_.isEmpty(customer.name)) {
    customer.missingName = 1
    customer.missingData = 1
  }

  if (_.isEmpty(customer.provinceId)) {
    customer.missingLivingCity = 1
    customer.missingData = 1
  }

  if (_.isEmpty(customer.phoneNumber) && _.isEmpty(customer.parentPhoneNumber) && _.isEmpty(customer.facebook) && _.isEmpty(customer.email)) {
    customer.missingContactInformation = 1
    customer.missingData = 1
  }

  if (customer.missingContactInformation === 0 && cusomer.groupId === 3 && _.isEmpty(customer.phoneNumber) && _.isEmpty(customer.facebook) && _.isEmpty(customer.email)) {
    customer.missingContactInformation = 1
    customer.missingData = 1
  }

  if (_.isEmpty(customer.age)) {
    customer.missingAge = 1
    customer.missingData = 1
  }

  if (_.isEmpty(customer.schoolName)) {
    customer.missingSchoolName = 1
    customer.missingData = 1
  }

  if (_.isEmpty(customer.schoolName)) {
    customer.missingSchoolName = 1
    customer.missingData = 1
  }

  if (_.isEmpty(customer.kotexData) && _.isEmpty(customer.dianaData) && _.isEmpty(customer.laurierData) && _.isEmpty(customer.othersData) && _.isEmpty(customer.whisperData)) {
    customer.missingBrandUsing = 1
    customer.missingData = 1
  }

    if (_.isEmpty(customer.group)) {
    customer.missingGroup = 1
    customer.missingData = 1
  }

  return customer;
}

export const createCustomer = (customer) => {
  return new Promise((resolve, reject) => {
    if (customer.parentPhoneNumber && customer.parentPhoneNumber.length > 0) {
      customer.phoneNumber = '' + customer.phoneNumber.replace(/[\.\-\_\s\+\(\)]/g,'');
    }

    if (customer.parentPhoneNumber && customer.parentPhoneNumber.length > 0) {
      customer.parentPhoneNumber = '' + customer.parentPhoneNumber.replace(/[\.\-\_\s\+\(\)]/g,'');
    }

    _.each(provinces, (province) => {
      if (customer.province === province.name) {
        customer.provinceId = province.provinceId
      }
    })

    customer.contactInformation = customer.phoneNumber || customer.parentPhoneNumber || customer.facebook || customer.email

    if (customer.group === 'Học sinh') {
      customer.groupId = 1
    } else if (customer.group === 'Sinh viên') {
      customer.groupId = 2
    } else if (customer.group === 'Khác') {
      customer.groupId = 3
    } else {
      customer.groupId = 0
    }

    customer = checkMissingData(customer)

    db.run('INSERT INTO customers(\
          name, provinceId, schoolName, age, phoneNumber, parentPhoneNumber, facebook, email, contactInformation, kotexData, dianaData, laurierData, whisperData, othersData, createdAt, notes, receivedGift, groupId, batch, missingData, missingLivingCity, missingName, missingContactInformation, missingAge, missingSchoolName, missingBrandUsing, missingGroup\
        ) \
        VALUES($name, $provinceId, $schoolName, $age, $phoneNumber, $parentPhoneNumber, $facebook, $email, $contactInformation, $kotexData, $dianaData, $laurierData, $whisperData, $othersData, $createdAt, $notes, $receivedGift, $groupId, $batch, $missingData, $missingLivingCity, $missingName, $missingContactInformation, $missingAge, $missingSchoolName, $missingBrandUsing, $missingGroup);',
    {
      $name: customer.name,
      $provinceId: customer.provinceId,
      $schoolName: customer.schoolName,
      $age: customer.age,
      $phoneNumber: customer.phoneNumber,
      $parentPhoneNumber: customer.parentPhoneNumber,
      $facebook: customer.facebook,
      $email: customer.email,
      $contactInformation: customer.contactInformation,
      $kotexData: customer.kotexData,
      $dianaData: customer.dianaData,
      $laurierData: customer.laurierData,
      $whisperData: customer.whisperData,
      $othersData: customer.othersData,
      $createdAt: customer.createdAt,
      $notes: customer.notes,
      $receivedGift: customer.receivedGift,
      $groupId: customer.groupId,
      $batch: customer.batch,
      $missingData: customer.missingData,
      $missingLivingCity: customer.missingLivingCity,
      $missingName: customer.missingName,
      $missingContactInformation: customer.missingContactInformation,
      $missingAge: customer.missingAge,
      $missingSchoolName: customer.missingSchoolName,
      $missingBrandUsing: customer.missingBrandUsing,
      $missingGroup: customer.missingGroup
    }, (errRes) => {
      db.get('SELECT last_insert_rowid() as customerId', (err, row) => {
        customer.customerId = row.customerId;
        // isPhoneDuplicate(customer).then((customer) => {
        //   resolve(customer);
        // });
      });
    });
  });
}
