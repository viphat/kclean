import { db } from './database';
import { provinces } from './provinces';
import _ from 'lodash'

const checkIllogicalData = (customer) => {
  customer.illogicalData = 0
  customer.illogicalPhone = 0
  customer.illogicalAge = 0
  customer.illogicalAgePupil = 0
  customer.illogicalAgeStudent = 0
  customer.illogicalAgeOthers = 0

  var phone = customer.phoneNumber || customer.parentPhoneNumber

  if (!_.isEmpty(phone)) {
    phone = '' + phone.replace(/[\.\-\_\s\+\(\)]/g,'');
    if (!_.isNumber(phone)) {
      customer.illogicalPhone = 1;
      customer.illogicalData = 1;
    } else {
      if (phone.length < 8 || phone.length > 12) {
        customer.illogicalPhone = 1;
        customer.illogicalData = 1;
      }
    }
  }

  if (!_.isEmpty(customer.age)) {
    if (!_.isNumber(customer.age)) {
      customer.illogicalAge = 1
      customer.illogicalData = 1
    } else {
      var age = parseInt(customer.age)
      if (customer.groupId === 1 && age < 10 && age >= 18) {
        customer.illogicalAgePupil = 1
        customer.illogicalData = 1
      } else if (customer.groupId === 2 && age < 18 && age >= 30)  {
        customer.illogicalAgeStudent = 1
        customer.illogicalData = 1
      } else (customer.groupId === 3 && age < 18) {
        customer.illogicalAgeOthers = 1
        customer.illogicalData = 1
      }
    }
  }

  return customer;
}

const checkMissingData = (customer) => {
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

  if (customer.missingContactInformation === 0 && customer.groupId === 3 && _.isEmpty(customer.phoneNumber) && _.isEmpty(customer.facebook) && _.isEmpty(customer.email)) {
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
    customer = checkIllogicalData(customer)

    db.run('INSERT INTO customers(\
          name, provinceId, schoolName, age, phoneNumber, parentPhoneNumber, facebook, email, contactInformation, kotexData, dianaData, laurierData, whisperData, othersData, createdAt, notes, receivedGift, groupId, batch, hasError, missingData, missingLivingCity, missingName, missingContactInformation, missingAge, missingSchoolName, missingBrandUsing, missingGroup, illogicalData, illogicalPhone, illogicalAge, illogicalAgePupil, illogicalAgeStudent, illogicalAgeOthers\
        ) \
        VALUES($name, $provinceId, $schoolName, $age, $phoneNumber, $parentPhoneNumber, $facebook, $email, $contactInformation, $kotexData, $dianaData, $laurierData, $whisperData, $othersData, $createdAt, $notes, $receivedGift, $groupId, $batch, $hasError, $missingData, $missingLivingCity, $missingName, $missingContactInformation, $missingAge, $missingSchoolName, $missingBrandUsing, $missingGroup, $illogicalData, $illogicalPhone, $illogicalAge, $illogicalAgePupil, $illogicalAgeStudent, $illogicalAgeOthers);',
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
      $hasError: customer.hasError,
      $missingData: customer.missingData,
      $missingLivingCity: customer.missingLivingCity,
      $missingName: customer.missingName,
      $missingContactInformation: customer.missingContactInformation,
      $missingAge: customer.missingAge,
      $missingSchoolName: customer.missingSchoolName,
      $missingBrandUsing: customer.missingBrandUsing,
      $missingGroup: customer.missingGroup,
      $illogicalData: customer.illogicalData,
      $illogicalPhone: customer.illogicalPhone,
      $illogicalAge: customer.illogicalAge,
      $illogicalAgePupil: customer.illogicalAgePupil,
      $illogicalAgeStudent: customer.illogicalAgeStudent,
      $illogicalAgeOthers: customer.illogicalAgeOthers
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
