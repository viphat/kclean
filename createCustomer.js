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

  if (!isBlank(phone)) {
    if (isNaN(parseInt(phone))) {
      customer.illogicalPhone = 1;
      customer.illogicalData = 1;
    } else {
      if (phone.length !== 10) {
        customer.illogicalPhone = 1;
        customer.illogicalData = 1;
      } else {
        if (!phone.startsWith('02') && !phone.startsWith('03') && !phone.startsWith('05') && !phone.startsWith('07') && !phone.startsWith('08') && !phone.startsWith('09')) {
          customer.illogicalPhone = 1;
          customer.illogicalData = 1;
        }
      }
    }
  }

  if (!isBlank(customer.age)) {
    if (isNaN(parseInt(customer.age))) {
      customer.illogicalAge = 1
      customer.illogicalData = 1
    } else {
      var age = parseInt(customer.age)
      if (age < 10 || age > 99) {
        customer.illogicalAge = 1
        customer.illogicalData = 1
      } else if (customer.groupId === 1 && (age < 10 || age > 19)) {
        customer.illogicalAgePupil = 1
        customer.illogicalData = 1
      } else if (customer.groupId === 2 && (age < 18 || age >= 24))  {
        customer.illogicalAgeStudent = 1
        customer.illogicalData = 1
      } else if (customer.groupId === 3 && (age < 18 || age > 60)) {
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

  if (isBlank(customer.name)) {
    customer.missingName = 1
    customer.missingData = 1
  }

  if (isBlank(customer.provinceId)) {
    customer.missingLivingCity = 1
    customer.missingData = 1
  }

  if (customer.groupId === 1) {
    if (isBlank(customer.phoneNumber) && isBlank(customer.parentPhoneNumber) && isBlank(customer.facebook) && isBlank(customer.email)) {
      customer.missingContactInformation = 1
      customer.missingData = 1
    }
  } else if (customer.groupId === 2 || customer.groupId === 3) {
    if (isBlank(customer.phoneNumber)) {
      customer.missingContactInformation = 1
      customer.missingData = 1
    }
  }

  if (customer.missingContactInformation === 0 && customer.groupId === 3 && isBlank(customer.phoneNumber) && isBlank(customer.facebook) && isBlank(customer.email)) {
    customer.missingContactInformation = 1
    customer.missingData = 1
  }

  if (isBlank(customer.age)) {
    customer.missingAge = 1
    customer.missingData = 1
  }

  if (isBlank(customer.schoolName)) {
    customer.missingSchoolName = 1
    customer.missingData = 1
  }

  if (isBlank(customer.kotexData) && isBlank(customer.dianaData) && isBlank(customer.laurierData) && isBlank(customer.othersData) && isBlank(customer.whisperData)) {
    customer.missingBrandUsing = 1
    customer.missingData = 1
  }

  if (isBlank(customer.groupId)) {
    customer.missingGroup = 1
    customer.missingData = 1
  }

  return customer;
}

const checkDuplication = (customer) => {
  return new Promise((resolve, reject) => {
    customer.duplicatedPhone = 0
    customer.duplicatedPhoneBetweenPupilAndStudent = 0
    customer.duplicatedPhoneBetweenPupilAndOthers = 0
    customer.duplicatedPhoneBetweenStudentAndOthers = 0
    customer.duplicatedPhoneWithinPupil = 0
    customer.duplicatedPhoneWithinStudent = 0
    customer.duplicatedPhoneWithinOthers = 0


    if (customer.missingContactInformation === 1 || isBlank(customer.phoneNumber)) {
      return resolve(customer)
    }

    if (customer.illogicalPhone === 1) {
      return resolve(customer)
    }

    db.get('SELECT customers.customerId, customers.name, customers.areaName, customers.provinceName, customers.schoolName, customers.yearOfBirth,\
      customers.phoneNumber, customers.parentPhoneNumber, customers.facebook, customers.email, customers.kotexData, customers.dianaData, customers.laurierData, customers.whisperData, customers.othersData, customers.createdAt, customers.notes, customers.receivedGift, customers.groupName, customers.groupId, customers.batch\
    from customers\
    WHERE customers.phoneNumber = ?',
      customer.phoneNumber, (err, res) => {
      if (err) {
        return reject(err);
      }

      if (res === undefined || res === null) {
        resolve(customer);
      } else {
        customer.duplicatedWith = res;
        customer.duplicatedPhone = 1;

        if (customer.groupId === 1) {
          if (customer.duplicatedWith.groupId === 2) {
            customer.duplicatedPhoneBetweenPupilAndStudent = 1
          } else if (customer.duplicatedWith.groupId === 3) {
            customer.duplicatedPhoneBetweenPupilAndOthers = 1
          } else if (customer.duplicatedWith.groupId === customer.groupId) {
            customer.duplicatedPhoneWithinPupil = 1
          }
        } else if (customer.groupId === 2) {
          if (customer.duplicatedWith.groupId === 1) {
            customer.duplicatedPhoneBetweenPupilAndStudent = 1
          } else if (customer.duplicatedWith.groupId === 3) {
            customer.duplicatedPhoneBetweenStudentAndOthers = 1
          } else if (customer.duplicatedWith.groupId === customer.groupId) {
            customer.duplicatedPhoneWithinStudent = 1
          }
        } else if (customer.groupId === 3) {
          if (customer.duplicatedWith.groupId === 1) {
            customer.duplicatedPhoneBetweenPupilAndOthers = 1
          } else if (customer.duplicatedWith.groupId === 2) {
            customer.duplicatedPhoneBetweenStudentAndOthers = 1
          } else if (customer.duplicatedWith.groupId === customer.groupId) {
            customer.duplicatedPhoneWithinOthers = 1
          }
        }

        resolve(customer);
      }
    });
  })
}

export const createCustomer = (customer) => {
  return new Promise((resolve, reject) => {
    if (customer.phoneNumber && customer.phoneNumber.length > 0) {
      customer.phoneNumber = '' + customer.phoneNumber.replace(/[\.\-\_\s\+\(\)]/g,'');
    }

    if (customer.parentPhoneNumber && customer.parentPhoneNumber.length > 0) {
      customer.parentPhoneNumber = '' + customer.parentPhoneNumber.replace(/[\.\-\_\s\+\(\)]/g,'');
    }

    _.each(provinces, (province) => {
      if (customer.provinceName === province.name) {
        customer.provinceId = province.provinceId
      }
    })

    customer.contactInformation = customer.phoneNumber || customer.parentPhoneNumber || customer.facebook || customer.email

    if (customer.groupName === 'Học sinh') {
      customer.groupId = 1
    } else if (customer.groupName === 'Sinh viên') {
      customer.groupId = 2
    } else if (customer.groupName === 'Khác') {
      customer.groupId = 3
    } else {
      customer.groupId = 0
    }

    customer = checkMissingData(customer)
    customer = checkIllogicalData(customer)
    checkDuplication(customer).then((customer) => {
      if (customer.missingData === 1 || customer.illogicalData === 1 || customer.duplicatedPhone === 1) {
        customer.hasError = 1
      }

      db.run('INSERT INTO customers(\
            name, provinceId, areaName, provinceName, schoolName, yearOfBirth, age, phoneNumber, parentPhoneNumber, facebook, email, contactInformation, kotexData, dianaData, laurierData, whisperData, othersData, createdAt, notes, receivedGift, groupName, groupId, batch, hasError, missingData, missingLivingCity, missingName, missingContactInformation, missingAge, missingSchoolName, missingBrandUsing, missingGroup, illogicalData, illogicalPhone, illogicalAge, illogicalAgePupil, illogicalAgeStudent, illogicalAgeOthers,\
              duplicatedPhone, duplicatedPhoneBetweenPupilAndStudent, duplicatedPhoneBetweenPupilAndOthers, duplicatedPhoneBetweenStudentAndOthers, duplicatedPhoneWithinPupil, duplicatedPhoneWithinStudent, duplicatedPhoneWithinOthers\
          ) \
          VALUES($name, $provinceId, $areaName, $provinceName, $schoolName, $yearOfBirth, $age, $phoneNumber, $parentPhoneNumber, $facebook, $email, $contactInformation, $kotexData, $dianaData, $laurierData, $whisperData, $othersData, $createdAt, $notes, $receivedGift, $groupName, $groupId, $batch, $hasError, $missingData, $missingLivingCity, $missingName, $missingContactInformation, $missingAge, $missingSchoolName, $missingBrandUsing, $missingGroup, $illogicalData, $illogicalPhone, $illogicalAge, $illogicalAgePupil, $illogicalAgeStudent, $illogicalAgeOthers, $duplicatedPhone, $duplicatedPhoneBetweenPupilAndStudent, $duplicatedPhoneBetweenPupilAndOthers, $duplicatedPhoneBetweenStudentAndOthers, $duplicatedPhoneWithinPupil, $duplicatedPhoneWithinStudent, $duplicatedPhoneWithinOthers);',
      {
        $name: customer.name,
        $provinceId: customer.provinceId,
        $areaName: customer.areaName,
        $provinceName: customer.provinceName,
        $schoolName: customer.schoolName,
        $yearOfBirth: customer.yearOfBirth,
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
        $groupName: customer.groupName,
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
        $illogicalAgeOthers: customer.illogicalAgeOthers,
        $duplicatedPhone: customer.duplicatedPhone,
        $duplicatedPhoneBetweenPupilAndStudent: customer.duplicatedPhoneBetweenPupilAndStudent,
        $duplicatedPhoneBetweenPupilAndOthers: customer.duplicatedPhoneBetweenPupilAndOthers,
        $duplicatedPhoneBetweenStudentAndOthers: customer.duplicatedPhoneBetweenStudentAndOthers,
        $duplicatedPhoneWithinPupil: customer.duplicatedPhoneWithinPupil,
        $duplicatedPhoneWithinStudent: customer.duplicatedPhoneWithinStudent,
        $duplicatedPhoneWithinOthers: customer.duplicatedPhoneWithinOthers,
      }, (errRes) => {
        db.get('SELECT last_insert_rowid() as customerId', (err, row) => {
          customer.customerId = row.customerId;
          resolve(customer)
        });
      })
    })
  });
}

const isBlank = (value) => {
  return _.isEmpty(value) && (!_.isNumber(value) || _.isNaN(value))
}
