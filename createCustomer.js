import { db } from './database';
import _ from 'lodash'

const checkIllogicalData = (customer) => {
  customer.illogicalPhone = 0;
  customer.illogicalPhoneFormat = 0;
  customer.illogicalPhoneProvider = 0;

  var phone = customer.phoneNumber;

  if (!isBlank(phone)) {
    if (isNaN(parseInt(phone))) {
      customer.illogicalPhoneFormat = 1;
      customer.illogicalPhone = 1;
    } else {
      if (phone.length !== 10) {
        customer.illogicalPhoneFormat = 1;
        customer.illogicalPhone = 1;
      } else {
        if (!phone.startsWith('02') && !phone.startsWith('03') && !phone.startsWith('05') && !phone.startsWith('07') && !phone.startsWith('08') && !phone.startsWith('09')) {
          customer.illogicalPhoneProvider = 1;
          customer.illogicalPhone = 1;
        }
      }
    }
  }

  return customer;
}

const checkMissingData = (customer) => {
  customer.missingData = 0
  customer.missingName = 0
  customer.missingPhoneNumber = 0
  customer.missingAddress = 0
  customer.missingModel = 0

  if (isBlank(customer.name)) {
    customer.missingName = 1
    customer.missingData = 1
  }

  if (isBlank(customer.phoneNumber)) {
    customer.missingPhoneNumber = 1
    customer.missingData = 1
  }

  // if (isBlank(customer.address)) {
  //   customer.missingAddress = 1
  //   customer.missingData = 1
  // }

  if (isBlank(customer.model)) {
    customer.missingModel = 1
    customer.missingData = 1
  }

  return customer;
}

const checkDuplication = (customer) => {
  return new Promise((resolve, reject) => {
    customer.duplicatedPhone = 0
    customer.duplicatedPhoneSameModel = 0
    customer.duplicatedPhoneDiffModel = 0

    if (isBlank(customer.phoneNumber)) {
      return resolve(customer)
    }

    if (customer.illogicalPhone === 1) {
      return resolve(customer)
    }

    db.get('SELECT customers.customerId, customers.name,\
      customers.phoneNumber,\
      customers.address,\
      customers.city,\
      customers.model,\
      customers.batch\
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

        if (customer.model === customer.duplicatedWith.model) {
          customer.duplicatedPhoneSameModel = 1;
        } else {
          customer.duplicatedPhoneDiffModel = 1;
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

    customer = checkMissingData(customer)
    customer = checkIllogicalData(customer)

    checkDuplication(customer).then((customer) => {
      if (customer.missingPhoneNumber === 1 || customer.illogicalPhone === 1 || customer.duplicatedPhone === 1) {
        customer.hasError = 1
      }

      db.run('INSERT INTO customers(\
            name, phoneNumber, address, city, model, batch,\
            hasError,\
            missingData, missingName, missingPhoneNumber, missingAddress, missingModel,\
            illogicalPhone, illogicalPhoneFormat, illogicalPhoneProvider,\
             duplicatedPhone, duplicatedPhoneSameModel, duplicatedPhoneDiffModel\
          ) \
          VALUES($name, $phoneNumber, $address, $city, $model, $batch,\
          $hasError,\
          $missingData, $missingName, $missingPhoneNumber, $missingAddress, $missingModel,\
          $illogicalPhone, $illogicalPhoneFormat, $illogicalPhoneProvider,\
          $duplicatedPhone, $duplicatedPhoneSameModel, $duplicatedPhoneDiffModel);',
      {
        $name: customer.name,
        $phoneNumber: customer.phoneNumber,
        $address: customer.address,
        $city: customer.city,
        $model: customer.model,
        $batch: customer.batch,
        $hasError: customer.hasError,
        $missingData: customer.missingData,
        $missingName: customer.missingName,
        $missingPhoneNumber: customer.missingPhoneNumber,
        $missingAddress: customer.missingAddress,
        $missingModel: customer.missingModel,
        $illogicalPhone: customer.illogicalPhone,
        $illogicalPhoneFormat: customer.illogicalPhoneFormat,
        $illogicalPhoneProvider: customer.illogicalPhoneProvider,
        $duplicatedPhone: customer.duplicatedPhone,
        $duplicatedPhoneSameModel: customer.duplicatedPhoneSameModel,
        $duplicatedPhoneDiffModel: customer.duplicatedPhoneDiffModel,
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
