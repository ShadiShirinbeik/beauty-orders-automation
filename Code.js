const ORDERS_SHEET = 'Orders';
const SUMMARY_SHEET = 'Summary';

const HEADERS = ['Date', 'Customer', 'Brand', 'Category', 'Qty', 'Amount', 'OrderStatus', 'Flag'];

// Zero-based column positions in an Orders row, matching HEADERS.
const COL = { DATE: 0, CUSTOMER: 1, BRAND: 2, CATEGORY: 3, QTY: 4, AMOUNT: 5, STATUS: 6, FLAG: 7 };

const EUR_FORMAT = '[$€]#,##0.00';

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
 * Returns the sheet with the given name, creating it if it does not exist.
 */
function getOrCreateSheet_(name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
}

/**
 * Clears the Orders sheet and fills it with 30 mock orders:
 * 27 random orders, 2 of them with a blank Customer or Amount, plus 3 exact duplicates.
 */
function seedOrders() {
  const sheet = getOrCreateSheet_(ORDERS_SHEET);
  sheet.clear();

  const orders = [];
  for (let i = 0; i < 27; i++) {
    orders.push(randomOrder_());
  }

  // Planted problems: one missing Customer, one missing Amount.
  orders[3][COL.CUSTOMER] = '';
  orders[8][COL.AMOUNT] = '';

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
  const sheet = getOrCreateSheet_(ORDERS_SHEET);
  const counts = { duplicates: 0, missing: 0, clean: 0 };
  const numRows = sheet.getLastRow() - 1;
  if (numRows < 1) {
    console.log(counts);
    return counts;
  }

  const rows = sheet.getRange(2, 1, numRows, HEADERS.length).getValues();
  const seenKeys = new Set();

  const flags = rows.map(row => {
    if (row[COL.CUSTOMER] === '' || row[COL.AMOUNT] === '') {
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

  sheet.getRange(2, COL.FLAG + 1, numRows, 1).setValues(flags);
  console.log(counts);
  return counts;
}

/**
 * Returns a text key from Date + Customer + Brand + Amount, used to spot duplicate orders.
 */
function orderKey_(row) {
  const date = toDate_(row[COL.DATE]);
  return JSON.stringify([
    date ? date.getTime() : row[COL.DATE],
    row[COL.CUSTOMER],
    row[COL.BRAND],
    row[COL.AMOUNT],
  ]);
}

/**
 * Rebuilds the Summary sheet from the unflagged rows on the Orders sheet and returns the stats.
 */
function buildSummary() {
  const stats = computeStats_(readCleanOrders_());
  const sheet = getOrCreateSheet_(SUMMARY_SHEET);
  sheet.clear();
  writeSummary_(sheet, stats);
  return stats;
}

/**
 * Reads all Orders data rows in one call and returns only those with an empty Flag.
 */
function readCleanOrders_() {
  const sheet = getOrCreateSheet_(ORDERS_SHEET);
  const numRows = sheet.getLastRow() - 1;
  if (numRows < 1) return [];
  const rows = sheet.getRange(2, 1, numRows, HEADERS.length).getValues();
  return rows.filter(row => row[COL.FLAG] === '');
}

/**
 * Computes summary figures. Returned orders count as orders but add no revenue.
 */
function computeStats_(orders) {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayKey = dayKey_(yesterday);

  const stats = { totalOrders: orders.length, totalRevenue: 0, yesterdayRevenue: 0, returned: 0, revenueByCategory: {} };
  const revenueByBrand = {};

  orders.forEach(row => {
    if (row[COL.STATUS] === 'Returned') {
      stats.returned++;
      return;
    }
    const amount = row[COL.AMOUNT];
    stats.totalRevenue += amount;
    const date = toDate_(row[COL.DATE]);
    if (date && dayKey_(date) === yesterdayKey) {
      stats.yesterdayRevenue += amount;
    }
    revenueByBrand[row[COL.BRAND]] = (revenueByBrand[row[COL.BRAND]] || 0) + amount;
    stats.revenueByCategory[row[COL.CATEGORY]] = (stats.revenueByCategory[row[COL.CATEGORY]] || 0) + amount;
  });

  stats.topBrand = sortedByValue_(revenueByBrand)[0] || 'n/a';
  return stats;
}

/**
 * Converts a cell value (a date, or text like an ISO timestamp) to a Date, or null if it isn't one.
 * Avoids `instanceof Date`, which can fail for dates returned by getValues().
 */
function toDate_(value) {
  if (!value || (typeof value !== 'object' && typeof value !== 'string')) return null;
  const date = new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

/**
 * Returns a date as 'yyyy-MM-dd' in the script's time zone, so dates can be compared by day.
 */
function dayKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Returns the keys of a { key: number } object, highest value first.
 */
function sortedByValue_(totals) {
  return Object.keys(totals).sort((a, b) => totals[b] - totals[a]);
}

/**
 * Writes the metrics table and the revenue-by-category table in one call, then formats them.
 */
function writeSummary_(sheet, stats) {
  const metrics = [
    ['Metric', 'Value'],
    ['Total orders', stats.totalOrders],
    ['Total revenue', stats.totalRevenue],
    ['Revenue yesterday', stats.yesterdayRevenue],
    ['Returned orders', stats.returned],
    ['Top brand by revenue', stats.topBrand],
    ['Last updated', new Date()],
  ];
  const categories = sortedByValue_(stats.revenueByCategory)
    .map(category => [category, stats.revenueByCategory[category]]);

  // Layout: metrics, a blank row, a title row, a header row, then one row per category.
  const rows = metrics.concat([['', ''], ['Revenue by category', ''], ['Category', 'Revenue']], categories);
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);

  const titleRow = metrics.length + 2;
  const firstCategoryRow = titleRow + 2;
  const lastCategoryRow = firstCategoryRow + categories.length - 1;

  sheet.getRangeList(['A1:B1', `A${titleRow}:B${titleRow + 1}`]).setFontWeight('bold');
  const currencyCells = ['B3:B4'];
  if (categories.length > 0) currencyCells.push(`B${firstCategoryRow}:B${lastCategoryRow}`);
  sheet.getRangeList(currencyCells).setNumberFormat(EUR_FORMAT);
  sheet.getRange(metrics.length, 2).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.autoResizeColumns(1, 2);
}

/**
 * Daily job: flags bad rows, rebuilds the Summary sheet, then emails the digest.
 */
function runDailyPipeline() {
  const counts = flagOrders();
  const stats = buildSummary();
  stats.flagged = counts.duplicates + counts.missing;
  sendDailyDigest(stats);
}

/**
 * Emails a short HTML digest of yesterday's figures, with a link to the spreadsheet, to the script user.
 */
function sendDailyDigest(stats) {
  // getActiveUser() can be empty when run from a trigger; the effective user is then the trigger owner.
  const recipient = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const day = dayKey_(yesterday);
  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const lines = [
    ['Revenue yesterday (net)', formatEur_(stats.yesterdayRevenue)],
    ['Total orders', stats.totalOrders],
    ['Returned orders', stats.returned],
    ['Top brand', stats.topBrand],
    ['Flagged rows', stats.flagged],
  ];
  const plainBody = lines.map(([label, value]) => `${label}: ${value}`).join('\n') + `\n\nSpreadsheet: ${url}`;
  const tableRows = lines
    .map(([label, value]) => `<tr><td style="padding:4px 16px 4px 0">${label}</td><td><b>${value}</b></td></tr>`)
    .join('');
  const htmlBody = `<h3>Beauty orders – ${day}</h3><table>${tableRows}</table>` +
    `<p><a href="${url}">Open the spreadsheet</a></p>`;

  GmailApp.sendEmail(recipient, `Beauty orders digest – ${day}`, plainBody, { htmlBody: htmlBody });
}

/**
 * Formats a number as euros for the email, e.g. 1234.5 -> "€1234.50".
 */
function formatEur_(amount) {
  return `€${amount.toFixed(2)}`;
}

/**
 * Replaces any existing runDailyPipeline trigger with one that runs daily between 7 and 8 AM.
 */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'runDailyPipeline')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('runDailyPipeline')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .inTimezone(Session.getScriptTimeZone())
    .create();
}
