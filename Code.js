const HEADERS = ['Date', 'Customer', 'Brand', 'Category', 'Qty', 'Amount', 'OrderStatus', 'Flag'];

const CUSTOMERS = ['Sara Ahmadi', 'Lena Fischer', 'Maya Chen', 'Nora Haddad', 'Emma Rossi', 'Lina Karimi', 'Chloe Martin', 'Aisha Khan'];

// Each product ties a brand to a category it actually sells, with a unit price.
const PRODUCTS = [
  { brand: 'Dior', category: 'Perfume', price: 120 },
  { brand: 'Dior', category: 'Makeup', price: 45 },
  { brand: 'Lancôme', category: 'Skincare', price: 85 },
  { brand: 'Lancôme', category: 'Makeup', price: 38 },
  { brand: 'Lancôme', category: 'Perfume', price: 98 },
  { brand: 'The Ordinary', category: 'Skincare', price: 12 },
];

/**
 * Clears the active sheet and fills it with 30 mock orders:
 * 27 random orders, 2 of them with a blank Customer or Amount, plus 3 exact duplicates.
 */
function seedOrders() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();

  const orders = [];
  for (let i = 0; i < 27; i++) {
    orders.push(randomOrder_());
  }

  // Planted problems: one missing Customer, one missing Amount.
  orders[3][1] = '';
  orders[8][5] = '';

  // Planted problems: exact copies of three existing orders.
  [0, 5, 10].forEach(i => orders.push(orders[i].slice()));

  const rows = [HEADERS].concat(orders);
  sheet.getRange(1, 1, rows.length, HEADERS.length).setValues(rows);
}

/**
 * Returns one random order row from the last 7 days, matching HEADERS.
 */
function randomOrder_() {
  const date = new Date();
  date.setDate(date.getDate() - Math.floor(Math.random() * 7));
  date.setHours(0, 0, 0, 0);

  const customer = CUSTOMERS[Math.floor(Math.random() * CUSTOMERS.length)];
  const product = PRODUCTS[Math.floor(Math.random() * PRODUCTS.length)];
  const qty = 1 + Math.floor(Math.random() * 3);
  const status = Math.random() < 0.8 ? 'Shipped' : 'Returned';

  return [date, customer, product.brand, product.category, qty, qty * product.price, status, ''];
}

/**
 * Rewrites the Flag column: MISSING_DATA for rows without Customer or Amount,
 * DUPLICATE for repeats of an earlier row, empty otherwise. Returns the counts.
 */
function flagOrders() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const counts = { duplicates: 0, missing: 0, clean: 0 };
  const numRows = sheet.getLastRow() - 1;
  if (numRows < 1) {
    console.log(counts);
    return counts;
  }

  const rows = sheet.getRange(2, 1, numRows, HEADERS.length).getValues();
  const customerCol = HEADERS.indexOf('Customer');
  const amountCol = HEADERS.indexOf('Amount');
  const seenKeys = new Set();

  const flags = rows.map(row => {
    if (row[customerCol] === '' || row[amountCol] === '') {
      counts.missing++;
      return ['MISSING_DATA'];
    }
    const key = orderKey_(row);
    if (seenKeys.has(key)) {
      counts.duplicates++;
      return ['DUPLICATE'];
    }
    seenKeys.add(key);
    counts.clean++;
    return [''];
  });

  sheet.getRange(2, HEADERS.indexOf('Flag') + 1, numRows, 1).setValues(flags);
  console.log(counts);
  return counts;
}

/**
 * Returns a text key from Date + Customer + Brand + Amount, used to spot duplicate orders.
 */
function orderKey_(row) {
  const date = row[HEADERS.indexOf('Date')];
  return JSON.stringify([
    date instanceof Date ? date.getTime() : date,
    row[HEADERS.indexOf('Customer')],
    row[HEADERS.indexOf('Brand')],
    row[HEADERS.indexOf('Amount')],
  ]);
}
