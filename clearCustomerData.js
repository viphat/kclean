import { db } from './database.js';

export const clearCustomerData = (batch, source) => {
  return new Promise((resolve, reject) => {
    resolve(db.run('DELETE FROM customers WHERE customers.batch = ? AND customers.source = ?', batch, source));
  });
}
