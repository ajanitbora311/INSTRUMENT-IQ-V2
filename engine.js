const HEAD = {
  boq: ['SL. NO.', 'SHORT DESCRIPTION', 'JOB DESCRIPTION/ ACTIVITIES/LONG DESCRIPTION', 'UNIT', 'QUANTITY'],
  sor: ['ITEM CODE', 'DESCRIPTION', 'UNIT', 'RATE (₹)'],
  po: ['PO NUMBER', 'ITEM NUMBER', 'SERVICE LINE', 'DESCRIPTION', 'UNIT', 'RATE (₹)']
};

const OUTPUT = ['RATE (₹)', 'AMOUNT (₹)', 'RATE SOURCE (NRL SOR 2023)', 'RATE SOURCE (NRL POs)',
  'RATE SOURCE (to be mentioned)', 'Reference', 'Unit Conversion', 'Status', 'Review Reason',
  'Source Workbook', 'Source Sheet', 'Source Row', 'Source Description', 'Original Rate (₹)', 'Original Unit', 'Match Method', 'Input Sheet', 'Input Row'];

const norm = value => String(value ?? '').normalize('NFKC').toLowerCase().replace(/[‐‑–—]/g, '-').replace(/\s+/g, ' ').trim();
const header = value => norm(value).replace(/[^a-z0-9]/g, '');
const present = value => String(value ?? '').trim() !== '';
const aliases = {
  slno: ['sl', 'sno', 'serialno'], shortdescription: ['shorttext'],
  jobdescriptionactivitieslongdescription: ['jobdescription', 'activities', 'longdescription', 'jobdescriptionlongdescription'],
  unit: ['uom', 'unitofmeasure'], quantity: ['qty'], itemcode: ['soritemcode', 'soritemno', 'soritem', 'itemno', 'itemnumber', 'sorcode', 'scheduleitemcode', 'code'],
  description: ['jobdescription', 'servicedescription', 'servicedetails', 'servicetext', 'longdescription', 'longtext', 'activitydescription', 'materialdescription', 'itemdescription', 'shortdescription', 'shorttext', 'jobdescriptionactivitieslongdescription'],
  rate: ['unitrate', 'unitprice', 'rateinr', 'unitrateinr', 'rateperunit', 'netprice', 'priceperunit', 'price', 'unitpriceinr', 'netunitprice'],
  ponumber: ['po', 'pono', 'purchaseorderno', 'purchaseordernumber', 'purchasingdocument', 'ebeln'], itemnumber: ['itemno', 'poitem', 'poitemnumber', 'poitemno', 'ebelp', 'item'],
  serviceline: ['servicelineno', 'servicelinenumber', 'servicelineid', 'line', 'lineitem', 'serviceno']
};

function readBook(workbook, kind, name) {
  const result = [];
  let found = false;
  let missing = HEAD[kind];
  for (const sheet of workbook.SheetNames) {
    const ws = workbook.Sheets[sheet];
    if (!ws['!ref']) continue;
    const end = XLSX.utils.decode_range(ws['!ref']).e;
    if (end.r > 100000 || end.c > 300) throw new Error(name + ' has a sheet exceeding 100,001 rows or 301 columns.');
    const raw = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: '', blankrows: true, range: 0});
    let indexes;
    const start = raw.findIndex((row, i) => {
      if (i > 99) return false;
      const keys = row.map(header);
      indexes = HEAD[kind].map(label => {
        const key = header(label);
        const exact = keys.indexOf(key);
        return exact >= 0 ? exact : keys.findIndex(k => aliases[key]?.includes(k));
      });
      const absent = HEAD[kind].filter((_, column) => indexes[column] < 0);
      if (absent.length < missing.length) missing = absent;
      if (kind === 'boq') return indexes[3] >= 0 && indexes[4] >= 0 && (indexes[1] >= 0 || indexes[2] >= 0);
      return !absent.length && new Set(indexes).size === indexes.length;
    });
    if (start < 0) continue;
    found = true;
    const currencyIndex = raw[start].findIndex(value => ['currency', 'currencycode', 'waers'].includes(header(value)));
    for (let i = start + 1; i < raw.length; i++) {
      const cells = indexes.map(index => {
        if (index < 0) return '';
        const cell = ws[XLSX.utils.encode_cell({r:i, c:index})];
        return kind === 'boq' && cell?.f && !present(raw[i][index]) ? '=' + cell.f : raw[i][index];
      });
      if (!cells.some(present)) continue;
      const displayed = j => {
        if (indexes[j] < 0) return '';
        const cell = ws[XLSX.utils.encode_cell({r: i, c: indexes[j]})];
        return cell ? (cell.w ?? XLSX.utils.format_cell(cell)) : '';
      };
      const provenance = {name, sheet, row: i + 1};
      if (kind === 'boq') {
        const summary = /^(grand total|sub[ -]?total|total)$/.test(norm(cells[1] || cells[0]))
          && !present(cells[2]) && !present(cells[3]);
        result.push({...provenance, summary, input:[displayed(0) || cells[0], ...cells.slice(1)]});
        continue;
      }
      const [description, unit, rate] = cells.slice(-3);
      const currency = currencyIndex >= 0 ? norm(raw[i][currencyIndex]) : '';
      const references = cells.slice(0, -3).map((_, j) => displayed(j));
      if (/^(grand total|sub[ -]?total|total)$/.test(norm(description)) && !present(unit)) continue;
      if (!present(description)) continue;
      result.push({...provenance, kind, description, unit, rate: currency && !['inr', '₹', 'rs', 'rs.'].includes(currency) ? null : numeric(rate),
        reference: references.every(present) ? references.join(' / ') : ''});
    }
  }
  if (!found) throw new Error(name + (missing.length && missing.length < HEAD[kind].length ? ' is missing these headers: ' + missing.join(', ') : ' needs these headers within the first 100 rows: ' + HEAD[kind].join(', ')));
  if (!result.length) throw new Error(name + ' has no ' + (kind === 'boq' ? 'estimate' : 'rate') + ' rows.');
  return result;
}

function numeric(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const input = String(value ?? '').trim().replace(/^(₹|INR)\s*/i, '');
  if (!/^(?:\d+(?:\.\d+)?|\.\d+|\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d{1,2}(?:,\d{2})*,\d{3}(?:\.\d+)?)$/.test(input)) return null;
  const number = Number(input.replace(/,/g, ''));
  return Number.isFinite(number) ? number : null;
}

const units = new Map();
for (const [dimension, factor, symbols] of [
  ['count', 1, 'ea|each|no|nos|number|numbers|pc|pcs|piece|pieces'],
  ['length', 1, 'm|mtr|mtrs|metre|meter|metres|meters'],
  ['length', .001, 'mm|millimetre|millimeter'],
  ['length', .01, 'cm|centimetre|centimeter'],
  ['length', 1000, 'km|kilometre|kilometer'],
  ['length', .3048, 'ft|foot|feet'],
  ['area', 1, 'm2|m^2|sq m|sqm'],
  ['area', .09290304, 'ft2|ft^2|sq ft|sqft'],
  ['mass', 1, 'kg|kilogram|kilograms'],
  ['mass', 1000, 't|mt|tonne|tonnes|metric ton'],
  ['time', 1, 'h|hr|hrs|hour|hours']
]) {
  for (const symbol of symbols.split('|')) units.set(symbol, {dimension, factor});
}

function conversion(from, to) {
  const a = norm(from).replace(/\.$/, '');
  const b = norm(to).replace(/\.$/, '');
  if (!a || !b) return null;
  if (a === b) return 1;
  const source = units.get(a);
  const target = units.get(b);
  return source && target && source.dimension === target.dimension ? target.factor / source.factor : null;
}

const decimal = number => Number(number.toPrecision(12)).toString();
const money = amount => Math.round((amount + Number.EPSILON * Math.abs(amount)) * 100) / 100;

function priceRows(boq, sor, po, analysis = []) {
  const normalize = typeof semanticKey === 'function' ? semanticKey : norm;
  const signature = typeof pricingSignature === 'function' ? pricingSignature : () => '';
  const makeIndex = (records, keyOf) => {
    const index = new Map();
    for (const record of records) {
      for (const key of new Set([keyOf(record.description), keyOf(record.short)].filter(Boolean))) {
        if (!index.has(key)) index.set(key, []);
        index.get(key).push(record);
      }
    }
    return index;
  };
  const catalog = [sor, po, analysis].flatMap(records => [
    {index:makeIndex(records, norm), keyOf:norm, method:'Exact description'},
    {index:makeIndex(records, normalize), keyOf:normalize, method:'Equivalent terminology'},
    {index:makeIndex(records, signature), keyOf:signature, method:'Semantic match'}
  ]);
  return boq.map(row => {
    const [, short, long, unit, quantity] = row.input;
    // A full activity description takes precedence over its short label.
    const key = norm(present(long) ? long : short);
    const qty = numeric(quantity);
    const review = reason => ({...row, status: 'Review', reason});
    if (row.summary) return review('Subtotal or total copied from the input workbook.');
    if (!key || !present(unit) || qty === null || qty < 0) return review('Check the description, unit and quantity.');
    for (const {index, keyOf, method} of catalog) {
      const matchKey = keyOf(present(long) ? long : short);
      const matches = (index.get(matchKey) || []).flatMap(source => {
        if (present(long) && keyOf(source.description) !== matchKey
            && !(method === 'Semantic match' && keyOf(source.short) === matchKey)) return [];
        const factor = conversion(source.unit, unit);
        if (factor === null || source.rate === null || source.rate < 0 || !source.reference) return [];
        const rate = source.rate * factor;
        const amount = money(rate * qty);
        return Number.isFinite(amount) ? [{source, factor, rate, amount}] : [];
      });
      const unique = new Map(matches.map(match => [
        JSON.stringify([match.source.reference, norm(match.source.unit), match.source.rate]), match
      ]));
      if (unique.size > 1) {
        const labels = {sor:'NRL SOR', po:'service PO', analysis:'rate analysis'};
        return review('Multiple defensible ' + labels[matches[0].source.kind] + ' records. Verify the intended reference.');
      }
      if (unique.size === 1) {
        const match = unique.values().next().value;
        const {source, factor, rate} = match;
        const audit = norm(source.unit) === norm(unit) ? '' :
          '1 ' + unit + ' = ' + decimal(factor) + ' ' + source.unit + '; ₹' +
          decimal(source.rate) + ' × ' + decimal(factor) + ' = ₹' + decimal(rate) + ' / ' + unit;
        return {...row, ...match, audit, method, status: 'Priced'};
      }
    }
    return review(catalog.some(({index, keyOf}) => index.has(keyOf(present(long) ? long : short)))
      ? 'Matching description found, but its rate, reference or unit conversion could not be verified.'
      : 'No defensible description match in the fixed rate catalog.');
  });
}

function exportRows(rows, material = []) {
  const workbook = XLSX.utils.book_new();
  const data = rows.map(result => {
    const s = result.source;
    return [...result.input, result.rate, result.amount, s?.kind === 'sor' ? s.reference : '',
      s?.kind === 'po' ? s.reference : '', s?.kind === 'analysis' ? 'Rate analysis · ' + s.reference : '', s?.reference, result.audit, result.status,
      result.reason, s?.name, s?.sheet, s?.row, s?.description, s?.rate, s?.unit, result.method, result.sheet, result.row];
  });
  const totalRow = Array(HEAD.boq.length + OUTPUT.length).fill('');
  totalRow[1] = 'TOTAL (priced rows)';
  totalRow[6] = estimateTotal(rows);
  const ws = XLSX.utils.aoa_to_sheet([[...HEAD.boq, ...OUTPUT], ...data, totalRow]);
  ws['!cols'] = [10, 30, 50, 12, 14, 16, 18, 27, 27, 27, 26, 52, 14, 55, 26, 20, 14, 45, 20, 16, 26, 22, 12].map(wch => ({wch}));
  ws['!autofilter'] = {ref: 'A1:W' + (rows.length + 1)};
  for (let i = 2; i <= rows.length + 2; i++) {
    if (ws['F' + i]?.t === 'n') ws['F' + i].z = '0.00######';
    if (ws['G' + i]?.t === 'n') ws['G' + i].z = '#,##0.00';
  }
  XLSX.utils.book_append_sheet(workbook, ws, 'Instrumentation estimate');
  const materialIndexes = {short:new Map(), full:new Map()};
  for (const item of material) {
    const prefix = item.short + ' · ';
    const long = item.short && item.description.startsWith(prefix)
      ? item.description.slice(prefix.length) : item.description;
    const keys = {short:[item.short, item.description], full:[item.description, long, item.short + ' ' + long]};
    for (const [type, values] of Object.entries(keys)) {
      const index = materialIndexes[type];
      for (const key of new Set(values.map(norm).filter(Boolean))) {
        if (!index.has(key)) index.set(key, []);
        index.get(key).push(item);
      }
    }
  }
  const evidence = [];
  for (const row of rows) {
    if (row.status !== 'Review') continue;
    const [serial, short, long] = row.input;
    const matches = present(long) ? materialIndexes.full.get(norm(long)) : materialIndexes.short.get(norm(short));
    for (const item of matches || []) {
      evidence.push([serial, row.sheet, row.row, item.reference, item.description, item.unit, item.rate,
        item.currency, item.name, item.sheet, item.row]);
    }
  }
  if (evidence.length) {
    const sheet = XLSX.utils.aoa_to_sheet([['Take-off SL. NO.', 'Take-off Sheet', 'Take-off Row',
      'Material PO / Item', 'Description', 'Unit', 'Reported Price', 'Currency', 'Source File', 'Order', 'Item Index'], ...evidence]);
    sheet['!cols'] = [16, 20, 14, 26, 62, 12, 18, 14, 48, 14, 14].map(wch => ({wch}));
    XLSX.utils.book_append_sheet(workbook, sheet, 'Material references for review');
  }
  return workbook;
}

const estimateTotal = rows => money(rows.reduce((sum, row) => sum + (row.amount ?? 0), 0));
