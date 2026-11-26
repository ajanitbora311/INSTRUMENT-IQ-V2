const HEAD = {
  boq: ['SL. NO.', 'SHORT DESCRIPTION', 'JOB DESCRIPTION/ ACTIVITIES/LONG DESCRIPTION', 'UNIT', 'QUANTITY'],
  sor: ['ITEM CODE', 'DESCRIPTION', 'UNIT', 'RATE (₹)'],
  po: ['PO NUMBER', 'ITEM NUMBER', 'SERVICE LINE', 'DESCRIPTION', 'UNIT', 'RATE (₹)']
};

const OUTPUT = ['RATE (₹)', 'AMOUNT (₹)', 'RATE SOURCE (NRL SOR 2023)', 'RATE SOURCE (NRL POs)',
  'RATE SOURCE (to be mentioned)', 'Reference', 'Unit Conversion', 'Status', 'Review Reason',
  'Source Workbook', 'Source Sheet', 'Source Row', 'Source Description', 'Original Source Rate', 'Original Unit', 'Match Method', 'Input Sheet', 'Input Row', 'Input Item Code', 'Original Currency'];

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
    const itemColumn = kind === 'boq' ? raw[start].findIndex(value =>
      ['itemcode', 'itemreference', 'sourceitemcode', 'soritemcode', 'materialcode', 'item'].includes(header(value))) : -1;
    const currencyIndex = raw[start].findIndex(value => ['currency', 'currencycode', 'waers'].includes(header(value)));
    for (let i = start + 1; i < raw.length; i++) {
      const cells = indexes.map(index => {
        if (index < 0) return '';
        const cell = ws[XLSX.utils.encode_cell({r:i, c:index})];
        return kind === 'boq' && cell?.f && !present(raw[i][index]) ? '=' + cell.f : raw[i][index];
      });
      const itemCell = itemColumn < 0 ? undefined : ws[XLSX.utils.encode_cell({r:i, c:itemColumn})];
      const itemCode = itemCell ? XLSX.utils.format_cell(itemCell) : '';
      if (!cells.some(present) && !present(itemCode)) continue;
      const displayed = j => {
        if (indexes[j] < 0) return '';
        const cell = ws[XLSX.utils.encode_cell({r: i, c: indexes[j]})];
        return cell ? (cell.w ?? XLSX.utils.format_cell(cell)) : '';
      };
      const provenance = {name, sheet, row: i + 1};
      if (kind === 'boq') {
        const summary = /^(grand total|sub[ -]?total|total)$/.test(norm(cells[1] || cells[0]))
          && !present(cells[2]) && !present(cells[3]);
        if (summary && !present(itemCode)) continue;
        result.push({...provenance, summary, itemCode, input:[displayed(0) || cells[0], ...cells.slice(1)]});
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

function itemKey(value) {
  return norm(value).replace(/^sl\.?\s*no\.?(?:\s+|(?=\d))/, '').replace(/\s*\/\s*/g, '/');
}

function sourceItemKeys(source) {
  return new Set([source.reference, source.material].map(itemKey).filter(Boolean));
}

function materialDescriptions(item) {
  const prefix = item.short + ' · ';
  const long = item.short && item.description.startsWith(prefix)
    ? item.description.slice(prefix.length) : item.description;
  return {short:[item.short, item.description], full:[item.description, long, item.short + ' ' + long]};
}

function rateMatch(source, unit, qty) {
  const factor = conversion(source.unit, unit);
  if (factor === null || !Number.isFinite(source.rate) || source.rate < 0 || !source.reference) return null;
  const rate = source.rate * factor;
  const amount = qty === null ? undefined : money(rate * qty);
  if (!Number.isFinite(rate) || (amount !== undefined && !Number.isFinite(amount))) return null;
  const audit = norm(source.unit) === norm(unit) ? '' :
    '1 ' + unit + ' = ' + decimal(factor) + ' ' + source.unit + '; ₹' +
    decimal(source.rate) + ' × ' + decimal(factor) + ' = ₹' + decimal(rate) + ' / ' + unit;
  return {source, rate, amount, audit};
}

function uniqueRates(matches) {
  return [...new Map(matches.map(match => [
    JSON.stringify([match.source.reference, norm(match.source.unit), match.source.rate]), match
  ])).values()];
}

function priceRows(boq, sor, po, analysis = [], material = []) {
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
  const pricingSources = [sor, po, analysis];
  const itemIndexes = pricingSources.map(records => {
    const index = new Map();
    for (const record of records) for (const key of sourceItemKeys(record)) {
      if (!index.has(key)) index.set(key, []);
      index.get(key).push(record);
    }
    return index;
  });
  const catalog = pricingSources.flatMap((records, tier) => tier === 1
    ? [{index:new Map([['service', records]]), keyOf:() => 'service', method:'Service PO match'}]
    : [
    {index:makeIndex(records, norm), keyOf:norm, method:'Exact description'},
    {index:makeIndex(records, normalize), keyOf:normalize, method:'Equivalent terminology'},
    {index:makeIndex(records, signature), keyOf:signature, method:'Semantic match'}
  ]);
  const materialKey = value => normalize(value).replace(/^(?:supply of|supply only:?|supply|procurement of|purchase of)\s+/, '');
  const materialIndex = new Map();
  for (const source of material) {
    const descriptions = materialDescriptions(source);
    for (const key of new Set([...sourceItemKeys(source), ...descriptions.short.map(materialKey),
      ...descriptions.full.map(materialKey)].filter(Boolean))) {
      if (!materialIndex.has(key)) materialIndex.set(key, []);
      materialIndex.get(key).push(source);
    }
  }
  return boq.map(row => {
    const [, short, long, unit, quantity] = row.input;
    // Non-PO work rates use the full activity description when supplied.
    const key = norm(present(long) ? long : short);
    const parsedQty = numeric(quantity);
    const qty = parsedQty !== null && parsedQty >= 0 ? parsedQty : null;
    const review = reason => ({...row, status: 'Review', reason});
    if (row.summary) return review('Subtotal or total copied from the input workbook.');
    const explicitCode = itemKey(row.itemCode);
    const inlineCode = [short, long].map(itemKey).find(value => itemIndexes.some(index => index.has(value)) || materialIndex.get(value)?.some(source => sourceItemKeys(source).has(value)));
    const code = explicitCode || inlineCode;
    const description = [long, short].find(value => present(value) && (!code || itemKey(value) !== code));
    if ((!key && !code) || !present(unit)) return review('Check the description or item code, unit and quantity.');
    if (explicitCode && inlineCode && explicitCode !== inlineCode) return review('The input contains conflicting item codes.');
    const tiers = code ? itemIndexes.map(index => ({index, keyOf:() => code, method:'Item code'})) : catalog;
    for (const {index, keyOf, method} of tiers) {
      const matchKey = keyOf(description);
      const candidates = index.get(matchKey) || [];
      const service = method === 'Service PO match' || candidates[0]?.kind === 'po';
      if (service) {
        const fields = [[short, 'Short description'], [long, 'Long description']]
          .filter(([value]) => present(value) && (!code || itemKey(value) !== code));
        if (!fields.length && code) fields.push(['', 'Item code']);
        for (const [value, field] of fields) {
          for (const source of candidates) {
            const texts = field === 'Short description' ? [source.short, source.description] : [source.description];
            if (value && ![norm, normalize, signature].some(compare => {
              const key = compare(value);
              return key && texts.some(text => compare(text) === key);
            })) continue;
            const match = rateMatch(source, unit, qty);
            if (match) return {
              ...review('Provisional service PO rate. Verify the full scope before using it.'),
              ...match, method:field + ' · service PO'
            };
          }
        }
        continue;
      }
      const matches = candidates.flatMap(source => {
        if (code && description && ![norm, normalize, signature].some(compare => {
          const value = compare(description);
          return value && [source.description, source.short].some(text => compare(text) === value);
        })) return [];
        if (!code && present(long) && keyOf(source.description) !== matchKey
            && !(method === 'Semantic match' && keyOf(source.short) === matchKey)) return [];
        const match = rateMatch(source, unit, qty);
        return match ? [match] : [];
      });
      const unique = uniqueRates(matches);
      if (unique.length > 1) {
        const labels = {sor:'NRL SOR', po:'service PO', analysis:'rate analysis'};
        return review('Multiple defensible ' + labels[matches[0].source.kind] + ' records. Verify the intended reference.');
      }
      if (unique.length === 1) {
        if (qty === null) return {...review('Provisional rate. Check the quantity.'), ...unique[0], method};
        return {...row, ...unique[0], method, status:'Priced'};
      }
    }
    // A material purchase price supports a matching part, but cannot estimate work on that part.
    const serviceScope = [short, long].some(value => /\b(?:install|installation|calibrat\w*|commission\w*|repair\w*|remov\w*|test\w*|overhaul\w*|fabricat\w*|dismantl\w*|replac\w*|servic\w*|labor|labour|erect\w*|mount\w*|retrofit\w*|maintenan\w*)\b/.test(normalize(value)));
    if (!serviceScope) {
      const fields = [[short, 'short', 'Short description'], [long, 'full', 'Long description']]
        .filter(([value]) => present(value) && (!code || itemKey(value) !== code));
      if (!fields.length && code) fields.push(['', 'short', 'Item code']);
      for (const [value, type, field] of fields) {
        const key = materialKey(value);
        if (!key && value) continue;
        const candidates = materialIndex.get(code || key) || [];
        for (const source of candidates) {
          if (code && !sourceItemKeys(source).has(code)) continue;
          if (key && !materialDescriptions(source)[type].some(text => materialKey(text) === key)) continue;
          const match = rateMatch(source, unit, qty);
          if (!match) continue;
          const method = field + ' · material PO';
          if (!['inr', '₹'].includes(norm(source.currency))) return {
            ...review('Material PO rate shown in source currency. INR conversion requires a verified exchange rate.'),
            source, method
          };
          return {
            ...review(qty === null ? 'Provisional material PO rate. Verify the source and quantity.' : 'Provisional material PO rate. Verify the full specification before using it.'),
            ...match, method
          };
        }
      }
    }
    if (code) return review('No eligible rate for this item code and description. Check the reference, scope and unit.');
    return review(catalog.some(({index, keyOf, method}) => method !== 'Service PO match' && index.has(keyOf(present(long) ? long : short)))
      ? 'Matching description found, but its rate, reference or unit conversion could not be verified.'
      : 'No defensible description match in the fixed rate catalog.');
  });
}

function exportRows(rows, material = []) {
  const workbook = XLSX.utils.book_new();
  const data = rows.map(result => {
    const s = result.source;
    return [...result.input, result.rate, result.amount, s?.kind === 'sor' ? s.reference : '',
      s?.kind === 'po' ? s.reference : '', s?.kind === 'analysis' ? 'Rate analysis · ' + s.reference : s?.kind === 'material' ? 'Provisional material PO · ' + s.reference : '', s?.reference, result.audit, result.status,
      result.reason, s?.name, s?.sheet, s?.row, s?.description, s?.rate, s?.unit, result.method, result.sheet, result.row, result.itemCode, s ? (s.currency || (s.kind === 'material' ? 'Unspecified' : 'INR')) : undefined];
  });
  const ws = XLSX.utils.aoa_to_sheet([[...HEAD.boq, ...OUTPUT], ...data]);
  const lastItemRow = data.length + 1;
  const footerRow = lastItemRow + 1;
  XLSX.utils.sheet_add_aoa(ws, [['TOTAL (priced rows)', null, null, null, null, null, estimateTotal(rows)]],
    {origin:'A' + footerRow});
  ws['!merges'] = [XLSX.utils.decode_range('A' + footerRow + ':F' + footerRow)];
  ws['!cols'] = [10, 30, 50, 12, 14, 16, 18, 27, 27, 27, 26, 52, 14, 55, 26, 20, 14, 45, 20, 16, 26, 22, 12, 20, 18].map(wch => ({wch}));
  ws['!autofilter'] = {ref: 'A1:Y' + lastItemRow};
  for (let i = 2; i <= footerRow; i++) {
    if (ws['F' + i]?.t === 'n') ws['F' + i].z = '0.00######';
    if (ws['G' + i]?.t === 'n') ws['G' + i].z = '#,##0.00';
  }
  XLSX.utils.book_append_sheet(workbook, ws, 'Instrumentation estimate');
  const materialIndexes = {short:new Map(), full:new Map(), item:new Map()};
  for (const item of material) {
    const descriptions = materialDescriptions(item);
    const keys = {short:descriptions.short.map(norm), full:descriptions.full.map(norm), item:sourceItemKeys(item)};
    for (const [type, values] of Object.entries(keys)) {
      const index = materialIndexes[type];
      for (const key of new Set([...values].filter(Boolean))) {
        if (!index.has(key)) index.set(key, []);
        index.get(key).push(item);
      }
    }
  }
  const evidence = [];
  for (const row of rows) {
    if (row.status !== 'Review') continue;
    const [serial, short, long] = row.input;
    const codeMatches = [row.itemCode, short, long].flatMap(value => materialIndexes.item.get(itemKey(value)) || []);
    const descriptionMatches = present(long) ? materialIndexes.full.get(norm(long)) : materialIndexes.short.get(norm(short));
    for (const item of new Set([...codeMatches, ...(descriptionMatches || [])])) {
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

const estimateTotal = rows => money(rows.reduce((sum, row) => sum + (row.status === 'Priced' ? (row.amount ?? 0) : 0), 0));
