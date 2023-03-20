import { db } from './database';

export const clearCustomerData = (batch) => {
  return new Promise((resolve, reject) => {
    resolve(db.run('DELETE FROM customers WHERE customers.batch = ?', batch));
  });
}
