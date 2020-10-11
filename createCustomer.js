const _ = require('lodash');

import { db } from './database';

const checkIllogicalData = (customer) => {
  customer.illogicalData = 0
  customer.illogicalPhone = 0
  customer.illogicalAge = 0
  customer.illogicalAgePupil = 0
  customer.illogicalAgeStudent = 0

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
      customer.illogicalData = 1
      customer.illogicalAge = 1

      if (customer.source === 'BrandMax') {
        customer.illogicalAgePupil = 1
      } else if (customer.groupId === 'Focus MKT') {
        customer.illogicalAgeStudent = 1
      }
    } else {
      var age = parseInt(customer.age)
      if (customer.source === 'BrandMax' && (age < 12 || age > 20)) {
        customer.illogicalData = 1
        customer.illogicalAge = 1
        customer.illogicalAgePupil = 1
      } else if (customer.groupId === 'Focus MKT' && (age < 17 || age > 24))  {
        customer.illogicalData = 1
        customer.illogicalAge = 1
        customer.illogicalAgeStudent = 1
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
  customer.missingCollectedDate = 0
  customer.missingSamplingType = 0

  if (isBlank(customer.firstName) && isBlank(customer.lastName)) {
    customer.missingName = 1
    customer.missingData = 1
  }

  if (isBlank(customer.provinceName) || isBlank(customer.districtName)) {
    customer.missingLivingCity = 1
    customer.missingData = 1
  }

  if (customer.source === 'BrandMax') {
    if (isBlank(customer.phoneNumber) && isBlank(customer.parentPhoneNumber)) {
      customer.missingContactInformation = 1
      customer.missingData = 1
    }
  } else if (customer.source === 'Focus MKT') {
    if (isBlank(customer.phoneNumber)) {
      customer.missingContactInformation = 1
      customer.missingData = 1
    }
  }

  if (isBlank(customer.age)) {
    customer.missingAge = 1
    customer.missingData = 1
  }

  if (isBlank(customer.collectedDate)) {
    customer.missingCollectedDate = 1
    customer.missingData = 1
  }

  if (isBlank(customer.schoolName)) {
    customer.missingSchoolName = 1
    customer.missingData = 1
  }

  if (isBlank(customer.brand) || isBlank(customer.subBrand)) {
    customer.missingBrandUsing = 1
    customer.missingData = 1
  }

  if (isBlank(customer.samplingProduct)) {
    customer.missingSamplingType = 1
    customer.missingData = 1
  }

  return customer;
}

const checkDuplication = (customer) => {
  return new Promise((resolve, reject) => {
    customer.duplicatedPhone = 0
    customer.duplicatedPhoneBetweenPupilAndStudent = 0
    customer.duplicatedPhoneWithinPupil = 0
    customer.duplicatedPhoneWithinStudent = 0

    if (customer.missingContactInformation === 1) {
      return resolve(customer)
    }

    if (customer.illogicalPhone === 1) {
      return resolve(customer)
    }

    db.get('SELECT customers.customerId, customers.firstName, customers.lastName,\
      customers.districtId, customers.districtName, customers.provinceId, customers.provinceName,\
      customers.schoolName, customers.dateOfBirth, customers.collectedDate, customers.collectedTime,\
      customers.phoneNumber, customers.parentPhoneNumber,\
      customers.brand, customers.subBrand, customers.samplingProduct,\
      customers.gender, customers.optIn,\
      customers.source, customers.batch\
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

        if (customer.source === 'BrandMax') {
          if (customer.duplicatedWith.source === 'Focus MKT') {
            customer.duplicatedPhoneBetweenPupilAndStudent = 1
          } else if (customer.duplicatedWith.source === customer.source) {
            customer.duplicatedPhoneWithinPupil = 1
          }
        } else if (customer.source === 'Focus MKT') {
          if (customer.duplicatedWith.source === 'BrandMax') {
            customer.duplicatedPhoneBetweenPupilAndStudent = 1
          } else if (customer.duplicatedWith.source === customer.source) {
            customer.duplicatedPhoneWithinStudent = 1
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

    customer = checkMissingData(customer)
    customer = checkIllogicalData(customer)

    checkDuplication(customer).then((customer) => {
      if (customer.missingData === 1 || customer.illogicalData === 1 || customer.duplicatedPhone === 1) {
        customer.hasError = 1
      }

      db.run('INSERT INTO customers(\
            customerIndex, firstName, lastName, districtId, provinceId, districtName, provinceName, schoolName, phoneNumber, parentPhoneNumber, collectedDate, collectedTime, dateOfBirth, yearOfBirth, brand, subBrand, samplingProduct, gender, optIn, source, batch,\
            hasError, missingData, missingName, missingLivingCity, missingSchoolName, missingContactInformation, missingAge, missingCollectedDate, missingSchoolName, missingBrandUsing, missingSamplingType,\
            illogicalData, illogicalPhone, illogicalAge, illogicalAgePupil, illogicalAgeStudent,\
            duplicatedPhone, duplicatedPhoneBetweenPupilAndStudent, duplicatedPhoneWithinPupil, duplicatedPhoneWithinStudent) \
          VALUES($customerIndex, $firstName, $lastName, $districtId, $provinceId, $districtName, $provinceName, $schoolName, $phoneNumber, $parentPhoneNumber, $collectedDate, $collectedTime, $dateOfBirth, $yearOfBirth, $brand, $subBrand, $samplingProduct, $gender, $optIn, $source, $batch,\
            $hasError, $missingData, $missingName, $missingLivingCity, $missingSchoolName, $missingContactInformation, $missingAge, $missingCollectedDate, $missingBrandUsing, $missingSamplingType,\
            $illogicalData, $illogicalPhone, $illogicalAge, $illogicalAgePupil, $illogicalAgeStudent,\
            $duplicatedPhone, $duplicatedPhoneBetweenPupilAndStudent, $duplicatedPhoneWithinPupil, $duplicatedPhoneWithinStudent);',
      {
        $customerIndex: customer.customerIndex,
        $firstName: customer.firstName,
        $lastName: customer.lastName,
        $districtId: customer.districtId,
        $provinceId: customer.provinceId,
        $districtName: customer.districtName,
        $provinceName: customer.provinceName,
        $schoolName: customer.schoolName,
        $phoneNumber: customer.phoneNumber,
        $parentPhoneNumber: customer.parentPhoneNumber,
        $collectedDate: customer.collectedDate,
        $collectedTime: customer.collectedTime,
        $dateOfBirth: customer.dateOfBirth,
        $yearOfBirth: customer.yearOfBirth,
        $brand: customer.brand,
        $subBrand: customer.subBrand,
        $samplingProduct: customer.samplingProduct,
        $gender: customer.gender,
        $optIn: customer.optIn,
        $source: customer.source,
        $batch: customer.batch,
        $hasError: customer.hasError,
        $missingData: customer.missingData,
        $missingName: customer.missingName,
        $missingLivingCity: customer.missingLivingCity,
        $missingSchoolName: customer.missingSchoolName,
        $missingContactInformation: customer.missingContactInformation,
        $missingAge: customer.missingAge,
        $missingCollectedDate: customer.missingCollectedDate,
        $missingBrandUsing: customer.missingBrandUsing,
        $missingSamplingType: customer.missingSamplingType,
        $illogicalData: customer.illogicalData,
        $illogicalPhone: customer.illogicalPhone,
        $illogicalAge: customer.illogicalAge,
        $illogicalAgePupil: customer.illogicalAgePupil,
        $illogicalAgeStudent: customer.illogicalAgeStudent,
        $duplicatedPhone: customer.duplicatedPhone,
        $duplicatedPhoneBetweenPupilAndStudent: customer.duplicatedPhoneBetweenPupilAndStudent,
        $duplicatedPhoneWithinPupil: customer.duplicatedPhoneWithinPupil,
        $duplicatedPhoneWithinStudent: customer.duplicatedPhoneWithinStudent,
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
