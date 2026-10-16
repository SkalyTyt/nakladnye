// Распознавание накладных из Excel и подстановка штрихкодов.
// Работает и в браузере (window.InvoiceParser), и в Node (require).
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.InvoiceParser = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const FIELDS = [
    { key: 'name', label: 'Наименование' },
    { key: 'barcode', label: 'Штрихкод' },
    { key: 'qty', label: 'Количество' },
    { key: 'price', label: 'Цена' },
    { key: 'sum', label: 'Сумма' },
    { key: 'code', label: 'Код / артикул' },
    { key: 'unit', label: 'Ед. изм.' },
    { key: 'vat', label: 'НДС' },
  ];
  const DETECT_FIELDS = ['name', 'barcode', 'qty', 'price', 'sum'];
  const LETTER = /[a-zа-яё]/i;

  // ---------- утилиты ----------

  function normHeader(s) {
    return String(s == null ? '' : s).toLowerCase().replace(/ё/g, 'е')
      .replace(/­/g, '').replace(/-\s+/g, '').replace(/\s+/g, ' ').trim();
  }

  function toNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    if (v == null) return NaN;
    const s = String(v).replace(/[\s ]/g, '').replace(/,/g, '.');
    return /^-?\d+(\.\d+)?$/.test(s) ? parseFloat(s) : NaN;
  }

  const round2 = (x) => Math.round(x * 100) / 100;
  const clean = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();

  function cleanBarcode(v) {
    if (v == null || v === '') return '';
    let s = typeof v === 'number' ? (Number.isInteger(v) ? String(v) : v.toFixed(0)) : String(v);
    if (/^\s*\d+(\.\d+)?e\+\d+\s*$/i.test(s)) s = Number(s).toFixed(0);
    // в ячейке может быть несколько штрихкодов — берём первый
    const codes = s.match(/\d{8,14}/g);
    if (codes) return codes[0];
    return s.replace(/[\s ]/g, '').replace(/^(\d+)\.0+$/, '$1');
  }

  const MONTHS = ['январ', 'феврал', 'март', 'апрел', 'ма', 'июн', 'июл', 'август', 'сентябр', 'октябр', 'ноябр', 'декабр'];
  // «15 октября 2026» → «15.10.2026»
  function normDate(s) {
    const m = String(s).trim().match(/^(\d{1,2})\s+([а-яё]+)\s+(\d{4})/i);
    if (!m) return String(s).trim();
    const word = m[2].toLowerCase();
    let mi = -1;
    MONTHS.forEach((p, i) => { if (word.startsWith(p) && (mi < 0 || p.length > MONTHS[mi].length)) mi = i; });
    return mi < 0 ? String(s).trim() : `${m[1].padStart(2, '0')}.${String(mi + 1).padStart(2, '0')}.${m[3]}`;
  }

  function colLetter(c) {
    let s = '';
    for (c += 1; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
    return s;
  }

  // ---------- лист -> сетка ----------

  function sheetToGrid(ws, XLSX) {
    if (!ws || !ws['!ref']) return { vals: [], texts: [], merges: new Map(), rows: 0, cols: 0 };
    const range = XLSX.utils.decode_range(ws['!ref']);
    const rows = range.e.r + 1, cols = range.e.c + 1;
    const vals = [], texts = [];
    for (let r = 0; r < rows; r++) {
      const vr = new Array(cols).fill(null), tr = new Array(cols).fill('');
      for (let c = range.s.c; c < cols; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell || cell.v == null || cell.v === '') continue;
        vr[c] = cell.v;
        tr[c] = typeof cell.v === 'number' || cell.w == null ? String(cell.v) : String(cell.w);
        if (cell.t === 'd' && cell.w) tr[c] = String(cell.w);
      }
      vals.push(vr); texts.push(tr);
    }
    const merges = new Map();
    for (const m of ws['!merges'] || []) {
      if ((m.e.r - m.s.r + 1) * (m.e.c - m.s.c + 1) > 5000) continue;
      for (let r = m.s.r; r <= m.e.r; r++)
        for (let c = m.s.c; c <= m.e.c; c++)
          if (r !== m.s.r || c !== m.s.c) merges.set(r + ',' + c, [m.s.r, m.s.c]);
    }
    return { vals, texts, merges, rows, cols };
  }

  function textAt(g, r, c, useMerge) {
    let t = g.texts[r] && g.texts[r][c];
    if (!t && useMerge) {
      const m = g.merges.get(r + ',' + c);
      if (m) t = g.texts[m[0]][m[1]];
    }
    return t || '';
  }

  // ---------- распознавание столбцов ----------

  // Насколько заголовок h подходит полю (0 — не подходит).
  function scoreField(field, h) {
    if (!h) return 0;
    const isUnit = /единиц|ед\.?\s*изм|^ед\.?$/.test(h);
    switch (field) {
      case 'unit':
        return isUnit && !/код|океи/.test(h) ? 10 : 0;
      case 'barcode':
        return /штрих|ean|barcode|(^|[^а-я])шк([^а-я]|$)/.test(h) ? 10 : 0;
      case 'name':
        if (isUnit || /штрих|^код|океи/.test(h)) return 0;
        if (/наименование|номенклатура|название/.test(h)) return 10;
        return /товар/.test(h) ? 6 : 0;
      case 'code':
        if (/океи|штрих|наименование|товар/.test(h)) return 0;
        if (/^код(?![а-я])|артикул|^арт\.?$/.test(h)) return 10;
        return /код/.test(h) ? 5 : 0;
      case 'qty':
        if (/брутто|в одном месте|кор/.test(h)) return 0;
        if (/нетто/.test(h)) return 10;
        if (/мест/.test(h)) return 3;
        return /^штук|^шт\.?$|кол-?во|количество/.test(h) ? 8 : 0;
      case 'price':
        if (!/цена/.test(h)) return 0;
        if (/без/.test(h)) return 7;
        return /с ндс|с учетом ндс/.test(h) ? 10 : 8;
      case 'sum':
        if (!/сумма|стоимость/.test(h)) return 0;
        if (/без/.test(h)) return 7;
        if (/с ндс|с учетом|всего/.test(h)) return 10;
        return /ндс/.test(h) ? 0 : 8;
      case 'vat':
        return /ставка/.test(h) || (/ндс/.test(h) && /%/.test(h)) ? 10 : 0;
    }
    return 0;
  }

  function rowHeaderScore(g, r) {
    const found = new Set();
    for (let c = 0; c < g.cols; c++) {
      const h = normHeader(g.texts[r][c]);
      if (!h || h.length > 80) continue;
      for (const f of DETECT_FIELDS) if (scoreField(f, h) > 0) found.add(f);
    }
    return found.size;
  }

  function isDataLike(g, r) {
    let nums = 0, text = false;
    for (let c = 0; c < g.cols; c++) {
      const v = g.vals[r][c];
      if (v == null) continue;
      if (isFinite(toNum(v))) nums++;
      else if (String(v).trim().length >= 3 && LETTER.test(String(v))) text = true;
    }
    return nums >= 2 && text;
  }

  function findHeaderBlocks(g, minScore) {
    const blocks = [];
    for (let r = 0; r < g.rows; r++) {
      if (rowHeaderScore(g, r) < minScore) continue;
      let end = r;
      while (end + 1 < g.rows && end - r < 2 && !isDataLike(g, end + 1) && rowHeaderScore(g, end + 1) >= 1) end++;
      blocks.push({ start: r, end });
      r = end;
    }
    return blocks;
  }

  function extendBlock(g, start) {
    let end = start;
    while (end + 1 < g.rows && end - start < 2 && !isDataLike(g, end + 1) && rowHeaderScore(g, end + 1) >= 1) end++;
    return { start, end };
  }

  function headerTexts(g, block) {
    const out = [];
    for (let c = 0; c < g.cols; c++) {
      const parts = [];
      for (let r = block.start; r <= block.end; r++) {
        const t = textAt(g, r, c, true);
        if (t && parts[parts.length - 1] !== t) parts.push(t);
      }
      out.push(normHeader(parts.join(' ')));
    }
    return out;
  }

  function assignColumns(headers) {
    const cand = [];
    for (const f of FIELDS)
      headers.forEach((h, c) => {
        const s = scoreField(f.key, h);
        if (s > 0) cand.push({ f: f.key, c, s });
      });
    cand.sort((a, b) => b.s - a.s || a.c - b.c);
    const cols = {}, used = new Set();
    for (const x of cand) {
      if (cols[x.f] != null || used.has(x.c)) continue;
      cols[x.f] = x.c;
      used.add(x.c);
    }
    return cols;
  }

  // ---------- позиции и реквизиты ----------

  function extractItems(g, cols, from, to, barcodeOnly) {
    const items = [];
    let total = null, lastRow = from - 1;
    const at = (r, f) => (cols[f] != null ? g.vals[r][cols[f]] : null);
    for (let r = from; r < to; r++) {
      const first = g.texts[r].find((t) => t && t.trim()) || '';
      if (!first) continue;
      if (/^\s*(итого|всего)/i.test(first)) {
        const v = toNum(at(r, 'sum'));
        if (isFinite(v)) { total = v; lastRow = r; }
        continue;
      }
      const name = cols.name != null ? clean(g.texts[r][cols.name]) : '';
      if (!name || !LETTER.test(name)) continue;
      let qty = toNum(at(r, 'qty')), price = toNum(at(r, 'price')), sum = toNum(at(r, 'sum'));
      const barcode = cleanBarcode(at(r, 'barcode'));
      const hasBarcode = /^\d{8,14}$/.test(barcode);
      if (!isFinite(qty) && !isFinite(price) && !isFinite(sum) && !(barcodeOnly && hasBarcode)) continue;
      if (!isFinite(price) && isFinite(sum) && isFinite(qty) && qty) price = round2(sum / qty);
      if (!isFinite(sum) && isFinite(price) && isFinite(qty)) sum = round2(price * qty);
      items.push({
        row: r + 1,
        name,
        barcode,
        qty: isFinite(qty) ? qty : null,
        price: isFinite(price) ? price : null,
        sum: isFinite(sum) ? sum : null,
        code: cols.code != null ? clean(g.texts[r][cols.code]) : '',
        unit: cols.unit != null ? clean(g.texts[r][cols.unit]) : '',
        vat: cols.vat != null ? clean(g.texts[r][cols.vat]) : '',
      });
      lastRow = r;
    }
    return { items, total, lastRow };
  }

  function firstRight(g, r, c) {
    for (let k = c + 1; k < g.cols; k++) if (g.texts[r][k] && g.texts[r][k].trim()) return g.texts[r][k].trim();
    return '';
  }

  function cleanSupplier(s) {
    s = clean(s);
    const short = s.match(/,\s*((?:ООО|АО|ПАО|ЗАО|ОАО)\s*["«][^"»]+["»])/i);
    if (short) return short[1];
    return s.split(/,\s*(?:ИНН|КПП|ОГРН|р\/с|\d{6})/i)[0].slice(0, 120).trim();
  }

  const ORG = /^(ооо|ип|ао|пао|зао|оао|общество|индивидуальный предприниматель)[\s"«]/i;

  function extractMeta(g, a, b) {
    const meta = { number: '', date: '', supplier: '', inn: '' };
    let labeled = false;
    for (let r = Math.max(0, a); r < b; r++) {
      for (let c = 0; c < g.cols; c++) {
        const t = g.texts[r][c];
        if (!t) continue;
        const h = normHeader(t).replace(/[:\s]+$/, '');
        if (h === 'номер документа' && r + 1 < g.rows && g.texts[r + 1][c]) meta.number = g.texts[r + 1][c].trim();
        else if (h === 'дата составления' && r + 1 < g.rows && g.texts[r + 1][c]) meta.date = g.texts[r + 1][c].trim();
        else if (/^(поставщик|продавец)$/.test(h)) {
          const v = firstRight(g, r, c);
          if (v) {
            meta.supplier = cleanSupplier(v);
            const inn = v.match(/ИНН\s*:?\s*(\d{10,12})/i);
            meta.inn = inn ? inn[1] : '';
            labeled = true;
          }
        } else if (!labeled && !meta.supplier && ORG.test(t.trim())) {
          meta.supplier = cleanSupplier(t);
          const inn = t.match(/ИНН\s*:?\s*(\d{10,12})/i);
          meta.inn = inn ? inn[1] : '';
        }
        const m = t.match(/(?:накладн\S*|упд|сч[её]т[- ]?фактур\S*|сч[её]т(?:\s+на\s+оплату)?|заказ\S*)\s*№\s*([^\s,]+)(?:\s*от\s*(\d{1,2}[./]\d{1,2}[./]\d{2,4}|\d{1,2}\s+[а-яё]+\s+\d{4}))?/i);
        if (m) { meta.number = m[1]; if (m[2]) meta.date = normDate(m[2]); }
      }
    }
    return meta;
  }

  // ---------- разбор книги ----------

  // opts: { headerRow: номер строки с 1, cols: { поле: индекс столбца | -1 },
  //         barcodes: true — файл только с названиями и штрихкодами }
  function parseWorkbook(wb, XLSX, opts) {
    opts = opts || {};
    const sheets = [], docs = [], warnings = [];
    wb.SheetNames.forEach((sheetName, si) => {
      const g = sheetToGrid(wb.Sheets[sheetName], XLSX);
      sheets.push({ name: sheetName, grid: g });
      if (!g.rows) return;
      let blocks = opts.headerRow ? [extendBlock(g, opts.headerRow - 1)] : findHeaderBlocks(g, opts.barcodes ? 2 : 3);
      if (!blocks.length && !opts.headerRow) blocks = findHeaderBlocks(g, 2);
      let regionStart = 0;
      blocks.forEach((block, bi) => {
        const headers = headerTexts(g, block);
        const cols = assignColumns(headers);
        for (const [f, c] of Object.entries(opts.cols || {})) {
          if (c === -1 || c === '-1') delete cols[f];
          else if (c != null && c !== '') cols[f] = Number(c);
        }
        const to = bi + 1 < blocks.length ? blocks[bi + 1].start : g.rows;
        const metaStart = regionStart;
        const meta = extractMeta(g, metaStart, block.start);
        const { items, total, lastRow } = extractItems(g, cols, block.end + 1, to, !!opts.barcodes);
        regionStart = lastRow + 1;
        const prev = docs[docs.length - 1];
        // страницы одной накладной склеиваем
        if (prev && prev.sheetIndex === si && meta.number && prev.number === meta.number && prev.supplier === meta.supplier) {
          prev.items.push(...items);
          if (total != null) prev.fileTotal = total;
          prev.lastRow = lastRow;
          return;
        }
        if (!items.length) return;
        docs.push(Object.assign(meta, {
          sheet: sheetName, sheetIndex: si,
          headerRows: [block.start + 1, block.end + 1],
          regionStart: metaStart + 1, lastRow: lastRow + 1,
          cols, headers, items, fileTotal: total,
        }));
      });
    });
    if (!docs.length) warnings.push('Не нашёл таблицу с товарами. Укажите строку заголовка и столбцы в настройках.');
    return { sheets, docs, warnings };
  }

  // ---------- подстановка штрихкодов ----------

  function normName(s) {
    return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]+/g, ' ').trim();
  }
  const numsOf = (n) => n.match(/\d+/g) || [];
  const subset = (a, b) => a.every((x) => b.includes(x));

  function dice(a, b) {
    if (a === b) return 1;
    if (a.length < 2 || b.length < 2) return 0;
    const m = new Map();
    for (let i = 0; i < a.length - 1; i++) { const k = a.substr(i, 2); m.set(k, (m.get(k) || 0) + 1); }
    let inter = 0;
    for (let i = 0; i < b.length - 1; i++) {
      const k = b.substr(i, 2), n = m.get(k);
      if (n) { inter++; m.set(k, n - 1); }
    }
    return (2 * inter) / (a.length + b.length - 2);
  }

  function buildPool(items) {
    const byCode = new Map(), byName = new Map(), list = [];
    for (const it of items) {
      if (!it.barcode) continue;
      const n = normName(it.name);
      const e = { name: it.name, n, barcode: it.barcode, nums: numsOf(n) };
      list.push(e);
      if (it.code && !byCode.has(it.code)) byCode.set(it.code, e);
      if (!byName.has(n)) byName.set(n, e);
    }
    return { byCode, byName, list, size: list.length };
  }

  function matchInPool(item, pool) {
    if (item.code && pool.byCode.has(item.code)) return { e: pool.byCode.get(item.code), kind: 'code', score: 1 };
    const n = normName(item.name);
    if (pool.byName.has(n)) return { e: pool.byName.get(n), kind: 'name', score: 1 };
    const nums = numsOf(n);
    let best = null, bs = 0, second = 0;
    for (const e of pool.list) {
      // «Краска 501» и «Краска 502» — разные товары: цифры должны совпадать
      if (!subset(nums, e.nums) && !subset(e.nums, nums)) continue;
      const s = dice(n, e.n);
      if (s > bs) { second = bs; bs = s; best = e; } else if (s > second) second = s;
    }
    if (best && bs >= 0.75 && bs - second >= 0.03) return { e: best, kind: 'fuzzy', score: bs };
    return null;
  }

  // Справочник: { "<поставщик>\u0001<название>": { s, n, b, t } }
  const SEP = '\u0001';
  function supplierKey(doc) {
    return (doc && (doc.inn || normName(doc.supplier))) || 'без поставщика';
  }
  function dictKey(supKey, name) { return supKey + SEP + normName(name); }
  function indexDict(entries) {
    const exact = new Map(), byNorm = new Map();
    for (const [k, v] of Object.entries(entries || {})) {
      exact.set(k, v.b);
      const n = k.split(SEP)[1];
      if (!byNorm.has(n)) byNorm.set(n, v.b);
    }
    return { exact, byNorm };
  }
  function dictLookup(idx, supKey, name) {
    if (!idx) return '';
    return idx.exact.get(dictKey(supKey, name)) || idx.byNorm.get(normName(name)) || '';
  }

  // Порядок: ручная правка > штрихкод из накладной > второй файл (точно) > справочник > второй файл (похоже)
  function resolveItems(items, ctx) {
    ctx = ctx || {};
    return items.map((it, i) => {
      const out = Object.assign({}, it, { src: it.barcode ? 'file' : '', match: null });
      const ed = ctx.edits && ctx.edits[i];
      if (ed && ed.barcode != null) { out.barcode = ed.barcode; out.src = ed.barcode ? 'manual' : ''; return out; }
      if (out.barcode) return out;
      const m = ctx.pool && ctx.pool.size ? matchInPool(it, ctx.pool) : null;
      if (m && m.kind !== 'fuzzy') {
        out.barcode = m.e.barcode; out.src = 'merge';
        out.match = { name: m.e.name, score: 1, kind: m.kind };
        return out;
      }
      const d = dictLookup(ctx.dict, ctx.supplierKey, it.name);
      if (d) { out.barcode = d; out.src = 'dict'; return out; }
      if (m) {
        out.barcode = m.e.barcode;
        out.src = ed && ed.confirmed ? 'merge' : 'fuzzy';
        out.match = { name: m.e.name, score: m.score, kind: 'fuzzy' };
      }
      return out;
    });
  }

  return {
    FIELDS, colLetter, toNum, round2, normName, dice,
    sheetToGrid, scoreField, parseWorkbook,
    buildPool, matchInPool, resolveItems,
    supplierKey, dictKey, indexDict, dictLookup,
  };
});
