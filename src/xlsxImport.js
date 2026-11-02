const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

async function unzip(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const view = new DataView(bytes.buffer);
  let end = bytes.length - 22;
  while (end >= 0 && view.getUint32(end, true) !== 0x06054b50) end--;
  if (end < 0) throw new Error('This file is not a valid .xlsx workbook.');
  const count = view.getUint16(end + 10, true);
  let offset = view.getUint32(end + 16, true);
  const entries = new Map();
  for (let index = 0; index < count; index++) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error('The workbook archive is invalid.');
    const method = view.getUint16(offset + 10, true);
    const size = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = new TextDecoder().decode(bytes.slice(offset + 46, offset + 46 + nameLength));
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const start = localOffset + 30 + localNameLength + localExtraLength;
    let contents = bytes.slice(start, start + size);
    if (method === 8) {
      if (!('DecompressionStream' in window)) throw new Error('This browser cannot unpack Excel files. Please use a recent version of Chrome, Edge, or Firefox.');
      const stream = new Blob([contents]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      contents = new Uint8Array(await new Response(stream).arrayBuffer());
    } else if (method !== 0) throw new Error('This workbook uses an unsupported compression method.');
    entries.set(name, new TextDecoder().decode(contents));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

function xml(text) { return new DOMParser().parseFromString(text, 'application/xml'); }
function childText(node, tag) { return node.getElementsByTagNameNS(MAIN, tag)[0]?.textContent || ''; }
function columnNumber(reference) {
  let result = 0;
  for (const letter of reference.match(/[A-Z]+/)[0]) result = result * 26 + letter.charCodeAt(0) - 64;
  return result - 1;
}
function normalize(value) { return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }
const aliases = {
  name: ['name', 'donor', 'donor name', 'contributor', 'contributor name', 'person', 'stallholder', 'stall owner'],
  flat: ['flat', 'flat no', 'flat number', 'flatno', 'unit', 'unit no', 'apartment'],
  amount: ['amount', 'amount paid', 'payment', 'paid amount', 'total amount', 'amount inr', 'inr'],
  date: ['date', 'payment date', 'transaction date'], time: ['time', 'payment time', 'transaction time'],
  remarks: ['remarks', 'remark', 'notes', 'note', 'comments'],
  item: ['particular', 'particulars', 'item', 'expense', 'expense item', 'details', 'description'],
  paidTo: ['paid to', 'paidto', 'payee', 'vendor', 'recipient'],
};

function parseRows(rows, sheetName) {
  const fields = Object.fromEntries(Object.entries(aliases).map(([field, values]) => [field, values.map(normalize)]));
  let headerAt = -1, columns = {};
  for (let index = 0; index < Math.min(rows.length, 10); index++) {
    const found = {};
    rows[index].forEach((value, column) => {
      const header = normalize(value);
      for (const [field, values] of Object.entries(fields)) if (values.includes(header) && !(field in found)) found[field] = column;
      if (!('amount' in found) && (header.startsWith('amount ') || header.startsWith('total amount '))) found.amount = column;
      if (!('date' in found) && header.startsWith('date ')) found.date = column;
    });
    if (found.amount !== undefined && found.date !== undefined && (found.name !== undefined || found.item !== undefined)) { headerAt = index; columns = found; break; }
  }
  if (headerAt < 0) return null;
  const get = (row, field) => String(row[columns[field]] ?? '').trim();
  const entries = [];
  for (const row of rows.slice(headerAt + 1)) {
    const name = get(row, 'name'), item = get(row, 'item'), amount = get(row, 'amount'), date = get(row, 'date');
    if (normalize(item) === 'total expenses') continue;
    if (!name && !item && !amount && !date) continue;
    if (!amount || !date || (!name && !item)) throw new Error(`Incomplete payment row in sheet “${sheetName}”.`);
    const numericAmount = Number(amount.replace(/[₹,\s]/g, ''));
    if (!Number.isFinite(numericAmount) || numericAmount <= 0) throw new Error(`Invalid amount in sheet “${sheetName}”: ${amount}`);
    let parsedDate;
    if (/^\d+(\.\d+)?$/.test(date)) parsedDate = new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(date) * 86400000));
    else {
      const match = /^(\d{1,2})[ /-]([A-Za-z]{3,9})[ /-](\d{4})$/.exec(date);
      parsedDate = match ? new Date(`${match[2]} ${match[1]}, ${match[3]} UTC`) : new Date(date);
    }
    if (!Number.isFinite(parsedDate.getTime())) throw new Error(`Invalid date in sheet “${sheetName}”: ${date}`);
    const normalizedDate = `${parsedDate.getUTCDate()} ${parsedDate.toLocaleString('en-GB', {month:'short', timeZone:'UTC'})} ${parsedDate.getUTCFullYear()}`;
    if (columns.name !== undefined) entries.push({name, sourceSheet:sheetName, flat:get(row,'flat'), amount:String(numericAmount), date:normalizedDate, time:get(row,'time'), remarks:name.toLowerCase()==='anonymous'?'':get(row,'remarks')});
    else entries.push({date:normalizedDate, item, sourceSheet:sheetName, amount:String(numericAmount), paidTo:get(row,'paidTo'), remarks:get(row,'remarks')});
  }
  return entries;
}

export async function importWorkbook(file) {
  if (!file.name.toLowerCase().endsWith('.xlsx')) throw new Error('Choose an .xlsx workbook.');
  const files = await unzip(file);
  const workbook = xml(files.get('xl/workbook.xml') || '');
  const relationships = xml(files.get('xl/_rels/workbook.xml.rels') || '');
  const targets = new Map([...relationships.getElementsByTagNameNS(PKG_REL, 'Relationship')].map(rel => [rel.getAttribute('Id'), rel.getAttribute('Target')]));
  const sharedDoc = files.has('xl/sharedStrings.xml') ? xml(files.get('xl/sharedStrings.xml')) : null;
  const shared = sharedDoc ? [...sharedDoc.getElementsByTagNameNS(MAIN, 'si')].map(node => [...node.getElementsByTagNameNS(MAIN, 't')].map(text => text.textContent).join('')) : [];
  const receipts = [], stalls = [], expenses = [];
  let hasExpenseSheet = false;
  const sheets = [...workbook.getElementsByTagNameNS(MAIN, 'sheet')];
  for (const sheet of sheets) {
    const name = sheet.getAttribute('name');
    let target = targets.get(sheet.getAttributeNS(DOC_REL, 'id'));
    target = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    const doc = xml(files.get(target) || '');
    const rows = [...doc.getElementsByTagNameNS(MAIN, 'row')].map(row => {
      const values = [];
      for (const cell of row.getElementsByTagNameNS(MAIN, 'c')) {
        const value = childText(cell, 'v');
        const text = cell.getAttribute('t') === 's' ? shared[Number(value)] : cell.getAttribute('t') === 'inlineStr' ? childText(cell, 't') : value;
        values[columnNumber(cell.getAttribute('r'))] = text;
      }
      return values;
    });
    const records = parseRows(rows, name);
    if (!records) continue;
    if ((records.length && 'item' in records[0]) || /expense/i.test(name)) { hasExpenseSheet = true; expenses.push(...records); }
    else if (/stall/i.test(name)) stalls.push(...records);
    else receipts.push(...records);
  }
  if (!receipts.length && !stalls.length) throw new Error('No receipt sheets found. Include Name, Amount, and Date columns.');
  if (!hasExpenseSheet) throw new Error('No expense sheet found. Include Particular/Item, Amount, and Date columns.');
  const sum = records => records.reduce((total, row) => total + Number(row.amount), 0);
  const stem = file.name.replace(/\.xlsx$/i, '').replace(/[_-]+/g, ' ').trim();
  const name = stem.replace(/\b\w/g, letter => letter.toUpperCase());
  const year = Number(name.match(/(?:19|20)\d{2}/)?.[0]) || new Date().getFullYear();
  const id = `${name.toLowerCase().replace(/[^a-z0-9]+/g, '-')}-${Date.now().toString(36)}`;
  return {id, name, year, sourceSheets:sheets.map(sheet=>sheet.getAttribute('name')), receipts, stalls, expenses, totals:{contributions:sum(receipts), stallFees:sum(stalls), expenses:sum(expenses)}};
}
