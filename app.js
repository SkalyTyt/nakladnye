(() => {
  'use strict';
  const P = window.InvoiceParser;
  const $ = (s, el = document) => el.querySelector(s);

  // ---------- хранилище ----------
  const LS = {
    get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch { return false; } },
  };
  const K_PROFILES = 'nk.profiles', K_EXPORT = 'nk.export', K_DICT = 'nk.dict';

  const DEFAULT_EXPORT = [
    { key: 'name', title: 'Наименование', on: true },
    { key: 'barcode', title: 'Штрихкод', on: true },
    { key: 'qty', title: 'Количество', on: true },
    { key: 'price', title: 'Цена закупки', on: true },
    { key: 'sum', title: 'Сумма', on: true },
    { key: 'code', title: 'Артикул поставщика', on: false },
    { key: 'unit', title: 'Ед. изм.', on: false },
    { key: 'vat', title: 'НДС', on: false },
    { key: 'supplier', title: 'Поставщик', on: false },
    { key: 'docNumber', title: 'Номер накладной', on: false },
    { key: 'docDate', title: 'Дата накладной', on: false },
  ];
  const EXPORT_LABELS = Object.fromEntries(DEFAULT_EXPORT.map((c) => [c.key, c.title]));

  function getExportCfg() {
    const saved = LS.get(K_EXPORT, null);
    if (!Array.isArray(saved)) return DEFAULT_EXPORT.map((c) => ({ ...c }));
    const known = new Set(saved.map((c) => c.key));
    return saved.filter((c) => EXPORT_LABELS[c.key]).concat(DEFAULT_EXPORT.filter((c) => !known.has(c.key)).map((c) => ({ ...c, on: false })));
  }

  let dictEntries = LS.get(K_DICT, {});
  let dictIdx = P.indexDict(dictEntries);
  function saveDict() { LS.set(K_DICT, dictEntries); dictIdx = P.indexDict(dictEntries); }

  // ---------- состояние ----------
  const state = {
    main: null, // { fileName, wb, result, profileKey, profile }
    bc: null,
    pool: null,
    docIdx: 0,
    filter: 'all',
    q: '',
    view: 'items',
    edits: {}, // docIdx -> { itemIdx: { barcode?, confirmed? } }
    overrides: {}, // docIdx -> { supplier }
    pickTarget: 'main',
  };

  // ---------- форматирование ----------
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmtMoney = (n) => (n == null || !isFinite(n) ? '—' : n.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
  const fmtNum = (n) => (n == null || !isFinite(n) ? '—' : n.toLocaleString('ru-RU', { maximumFractionDigits: 3 }));
  const icon = (id, cls = '') => `<svg class="i ${cls}"><use href="#i-${id}"/></svg>`;
  const plural = (n, a, b, c) => { const m = n % 10, h = n % 100; return m === 1 && h !== 11 ? a : m >= 2 && m <= 4 && (h < 10 || h >= 20) ? b : c; };

  function toast(msg, err) {
    const el = document.createElement('div');
    el.className = 'toast' + (err ? ' err' : '');
    el.textContent = msg;
    $('#toasts').appendChild(el);
    setTimeout(() => el.remove(), err ? 6000 : 3500);
  }

  // ---------- загрузка файлов ----------
  async function readWb(file) {
    const buf = await file.arrayBuffer();
    return XLSX.read(new Uint8Array(buf), { type: 'array' });
  }

  const profileKeyFor = (slot, doc) => slot + '::' + P.supplierKey(doc);

  function parseSlot(slot, wb, forcedProfile) {
    const base = slot === 'bc' ? { barcodes: true } : {};
    let result = P.parseWorkbook(wb, XLSX, base);
    const key = result.docs[0] ? profileKeyFor(slot, result.docs[0]) : slot + '::' + 'без поставщика';
    const profile = forcedProfile !== undefined ? forcedProfile : LS.get(K_PROFILES, {})[key] || null;
    if (profile) result = P.parseWorkbook(wb, XLSX, { ...base, ...profile });
    return { result, profileKey: key, profile };
  }

  async function loadFile(slot, file) {
    let wb;
    try { wb = await readWb(file); } catch (e) { toast('Не удалось открыть файл: ' + (e.message || e), true); return; }
    const parsed = parseSlot(slot, wb);
    state[slot] = { fileName: file.name, wb, ...parsed };
    if (slot === 'main') { state.docIdx = 0; state.edits = {}; state.overrides = {}; state.view = 'items'; state.filter = 'all'; }
    else rebuildPool();
    const docs = parsed.result.docs;
    if (!docs.length) toast(parsed.result.warnings[0] || 'Не нашёл товары в файле', true);
    else if (slot === 'main') {
      const n = docs.reduce((s, d) => s + d.items.length, 0);
      toast(`Загружено: ${docs.length} ${plural(docs.length, 'накладная', 'накладные', 'накладных')}, ${n} ${plural(n, 'позиция', 'позиции', 'позиций')}`);
    }
    render();
  }

  async function loadDemo() {
    try {
      const get = async (f) => new File([await (await fetch('demo/' + f)).blob()], f);
      const [inv, bc] = await Promise.all([get('schet_na_oplatu_primer.xlsx'), get('shtrihkody_primer.xlsx')]);
      await loadFile('main', inv);
      await loadFile('bc', bc);
    } catch {
      toast('Пример открывается только в онлайн-версии. Загрузите свой файл.', true);
    }
  }

  function rebuildPool() {
    if (!state.bc) { state.pool = null; return; }
    const items = state.bc.result.docs.flatMap((d) => d.items);
    state.pool = P.buildPool(items);
    if (!state.pool.size) toast('Во втором файле не нашёл штрихкоды. Проверьте столбцы в настройках.', true);
    else toast(`Из второго файла взято ${state.pool.size} ${plural(state.pool.size, 'штрихкод', 'штрихкода', 'штрихкодов')}`);
  }

  function clearSlot(slot) {
    state[slot] = null;
    if (slot === 'bc') state.pool = null;
    else { state.edits = {}; state.overrides = {}; }
    render();
  }

  // ---------- расчёт ----------
  function resolvedDoc(i) {
    const d = state.main.result.docs[i];
    const supKey = P.supplierKey(d);
    const items = P.resolveItems(d.items, { pool: state.pool, dict: dictIdx, supplierKey: supKey, edits: state.edits[i] || {} });
    const ov = state.overrides[i] || {};
    return { ...d, supplier: ov.supplier != null ? ov.supplier : d.supplier, items, supKey };
  }

  function docStats(doc) {
    const s = { count: doc.items.length, sum: 0, withBc: 0, missing: 0, fuzzy: 0 };
    for (const it of doc.items) {
      s.sum += it.sum || 0;
      if (!it.barcode) s.missing++;
      else if (it.src === 'fuzzy') s.fuzzy++;
      else s.withBc++;
    }
    s.sum = P.round2(s.sum);
    s.totalOk = doc.fileTotal == null ? null : Math.abs(s.sum - doc.fileTotal) < 0.05;
    return s;
  }

  // ---------- отрисовка ----------
  function render() {
    renderSlot('main');
    renderSlot('bc');
    renderSummary();
    renderContent();
  }

  function renderSlot(slot) {
    const el = $('#slot-' + slot);
    const s = state[slot];
    if (!s) {
      el.innerHTML = `<div class="drop" data-pick="${slot}">
        <div class="ico">${icon(slot === 'main' ? 'upload' : 'barcode')}</div>
        <b>${slot === 'main' ? 'Перетащите накладную' : 'Перетащите файл со штрихкодами'}</b>
        <span>или нажмите, чтобы выбрать · .xlsx, .xls, .csv</span></div>`;
      return;
    }
    const docs = s.result.docs;
    const items = docs.reduce((a, d) => a + d.items.length, 0);
    const sub = slot === 'main'
      ? `${docs.length} ${plural(docs.length, 'накладная', 'накладные', 'накладных')} · ${items} поз.`
      : `${state.pool ? state.pool.size : 0} штрихкодов`;
    el.innerHTML = `<div class="file-chip" data-dropzone="${slot}">
        <span class="fi">${icon('file')}</span>
        <div class="meta"><b title="${esc(s.fileName)}">${esc(s.fileName)}</b><span>${sub}${s.profile ? ' · свои настройки' : ''}</span></div>
        <button class="btn ghost icon-btn" data-pick="${slot}" title="Заменить файл">${icon('upload')}</button>
        <button class="btn ghost icon-btn" data-clear="${slot}" title="Убрать">${icon('x')}</button>
      </div>`;
  }

  function renderSummary() {
    const el = $('#summary');
    if (!state.main || !state.main.result.docs.length) { el.hidden = true; return; }
    el.hidden = false;
    const doc = resolvedDoc(state.docIdx);
    const st = docStats(doc);
    const many = state.main.result.docs.length > 1;
    el.innerHTML = `
      <div class="step" style="margin-bottom:12px">${icon('check')}Итог по накладной</div>
      <div class="stats">
        <div class="stat"><span>Позиций</span><b>${st.count}</b></div>
        <div class="stat"><span>Сумма, ₽</span><b>${fmtMoney(st.sum)}</b></div>
        <div class="stat ok"><span>Со штрихкодом</span><b>${st.withBc}</b></div>
        <div class="stat ${st.missing ? 'bad' : ''}"><span>Без штрихкода</span><b>${st.missing}</b></div>
        ${st.fuzzy ? `<div class="stat warn" style="grid-column:1/-1"><span>Нужно проверить совпадения</span><b>${st.fuzzy}</b></div>` : ''}
      </div>
      <div class="actions">
        <button class="btn primary block" data-act="export">${icon('download')}Скачать Excel</button>
        ${many ? `<button class="btn block" data-act="export-all">${icon('download')}Скачать все ${state.main.result.docs.length} накладных</button>` : ''}
        <button class="btn block" disabled title="Подключим после настройки шаблона">${icon('send')}Отправить в Контур.Маркет <span class="soon">скоро</span></button>
      </div>`;
  }

  function renderContent() {
    const el = $('#content');
    if (!state.main) {
      el.innerHTML = `<div class="card empty">
        <div class="big">${icon('file')}</div>
        <h2>Загрузите накладную поставщика</h2>
        <p>Перетащите Excel-файл в окно или выберите его слева. Приложение само найдёт таблицу с товарами, проверит сумму с «Итого» и подготовит файл для Контур.Маркета.</p>
        <div class="features">
          <div>${icon('columns')}ТОРГ-12, УПД и простые накладные</div>
          <div>${icon('link')}Объединение файла с ценами и файла со штрихкодами</div>
          <div>${icon('book')}Запоминает штрихкоды по названию</div>
        </div>
        <button class="btn primary" style="margin-top:22px" data-act="demo">${icon('file')}Открыть пример</button>
        <p class="hint" style="margin-top:8px">Счёт вымышленного поставщика и отдельный файл со штрихкодами.</p>
      </div>`;
      return;
    }
    const docs = state.main.result.docs;
    if (!docs.length) {
      el.innerHTML = `<div class="warn-line">${icon('alert')}${esc(state.main.result.warnings[0] || 'Товары не найдены')}</div>
        <div class="card"><div class="toolbar"><b>Исходный файл</b><div class="spacer"></div><button class="btn sm" data-act="settings">${icon('gear')}Указать столбцы</button></div>
        <div class="raw-wrap">${rawTable(state.main.result.sheets[0].grid, null, { from: 0, to: 80 })}</div></div>`;
      return;
    }
    const doc = resolvedDoc(state.docIdx);
    const st = docStats(doc);

    const tabs = docs.length > 1 ? `<div class="doc-tabs">${docs.map((d, i) => {
      const r = resolvedDoc(i), s = docStats(r);
      const mark = s.missing ? `<span class="badge b-bad" style="margin-left:6px">${s.missing}</span>` : '';
      return `<button class="doc-tab ${i === state.docIdx ? 'active' : ''}" data-doc="${i}"><b>№ ${esc(d.number || i + 1)}${mark}</b><span>${d.items.length} поз. · ${fmtMoney(s.sum)} ₽</span></button>`;
    }).join('')}</div>` : '';

    const totalBadge = st.totalOk == null
      ? `<span class="badge b-muted">${icon('info')}Итого в файле не найдено</span>`
      : st.totalOk
        ? `<span class="badge b-ok">${icon('check')}Сумма совпадает с «Итого»</span>`
        : `<span class="badge b-bad">${icon('alert')}Расхождение: в файле ${fmtMoney(doc.fileTotal)} ₽</span>`;

    const head = `<div class="card doc-head">
      <div class="f"><span>Поставщик</span><input id="supplier-input" size="${Math.max(24, (doc.supplier || '').length + 2)}" value="${esc(doc.supplier)}" placeholder="Не найден — впишите"></div>
      <div class="f"><span>Номер</span><b>${esc(doc.number || '—')}</b></div>
      <div class="f"><span>Дата</span><b>${esc(doc.date || '—')}</b></div>
      <div class="f"><span>ИНН</span><b>${esc(doc.inn || '—')}</b></div>
      <div class="grow"></div>
      ${totalBadge}
    </div>`;

    const counts = { all: st.count, missing: st.missing, fuzzy: st.fuzzy, ok: st.withBc };
    const seg = [['all', 'Все'], ['missing', 'Без штрихкода'], ['fuzzy', 'Проверить'], ['ok', 'Готово']]
      .map(([k, t]) => `<button class="${state.filter === k ? 'on' : ''}" data-filter="${k}">${t}<span class="c">${counts[k]}</span></button>`).join('');
    const views = `<div class="seg"><button class="${state.view === 'items' ? 'on' : ''}" data-view="items">Распознанные товары</button><button class="${state.view === 'raw' ? 'on' : ''}" data-view="raw">Исходный файл</button></div>`;

    let body;
    if (state.view === 'raw') {
      const g = state.main.result.sheets[doc.sheetIndex].grid;
      body = `<div class="toolbar">${views}<div class="spacer"></div><span class="hint" style="margin:0">Лист «${esc(doc.sheet)}», строки ${doc.regionStart}–${doc.lastRow}. Синим выделены столбцы, которые берутся в накладную.</span>
        <button class="btn sm" data-act="settings">${icon('gear')}Изменить столбцы</button></div>
        <div class="raw-wrap">${rawTable(g, doc, { from: doc.regionStart - 1, to: doc.lastRow + 2 })}</div>`;
    } else {
      body = `<div class="toolbar">${views}<div class="seg">${seg}</div>
        <label class="search">${icon('search')}<input id="q" placeholder="Поиск по названию или штрихкоду" value="${esc(state.q)}"></label>
        <div class="spacer"></div>
        <button class="btn sm" data-act="learn" title="Сохранить штрихкоды этой накладной в справочник">${icon('save')}Запомнить штрихкоды</button>
      </div>
      <div class="table-wrap" id="table-wrap">${itemsTable(doc)}</div>`;
    }
    el.innerHTML = tabs + head + `<div class="card">${body}</div>`;
  }

  function statusBadge(it, i) {
    switch (it.src) {
      case 'file': return `<span class="badge b-muted">Из накладной</span>`;
      case 'merge': return `<span class="badge b-ok" title="${esc(it.match ? 'Совпало с: ' + it.match.name : '')}">${icon('link')}Из 2-го файла</span>`;
      case 'fuzzy': return `<button class="badge b-warn" data-confirm="${i}" title="Похоже на: ${esc(it.match.name)}. Нажмите, если это тот же товар">${icon('alert')}Проверить · ${Math.round(it.match.score * 100)}%</button>`;
      case 'dict': return `<span class="badge b-info">${icon('book')}Из справочника</span>`;
      case 'manual': return `<span class="badge b-accent">Вручную</span>`;
      default: return `<span class="badge b-bad">Нет штрихкода</span>`;
    }
  }

  function itemsTable(doc) {
    const q = state.q.trim().toLowerCase();
    const rows = doc.items.map((it, i) => ({ it, i })).filter(({ it }) => {
      if (state.filter === 'missing' && it.barcode) return false;
      if (state.filter === 'fuzzy' && it.src !== 'fuzzy') return false;
      if (state.filter === 'ok' && (!it.barcode || it.src === 'fuzzy')) return false;
      if (q && !it.name.toLowerCase().includes(q) && !String(it.barcode).includes(q)) return false;
      return true;
    });
    if (!rows.length) return `<div class="table-empty">Ничего не найдено</div>`;
    return `<table><thead><tr>
        <th>#</th><th>Наименование</th><th>Штрихкод</th><th class="n">Кол-во</th><th class="n">Цена</th><th class="n">Сумма</th><th>Статус</th>
      </tr></thead><tbody>${rows.map(({ it, i }) => `<tr>
        <td class="idx">${i + 1}</td>
        <td class="name">${esc(it.name)}${it.src === 'fuzzy' ? `<small>≈ ${esc(it.match.name)}</small>` : it.code ? `<small>Код ${esc(it.code)}${it.unit ? ' · ' + esc(it.unit) : ''}</small>` : ''}</td>
        <td><input class="bc-input ${!it.barcode ? 'missing' : it.src === 'fuzzy' ? 'fuzzy' : ''}" data-bc="${i}" value="${esc(it.barcode)}" placeholder="Введите" inputmode="numeric"></td>
        <td class="n">${fmtNum(it.qty)}</td>
        <td class="n">${fmtMoney(it.price)}</td>
        <td class="n">${fmtMoney(it.sum)}</td>
        <td>${statusBadge(it, i)}</td>
      </tr>`).join('')}</tbody></table>`;
  }

  // Таблица исходного листа; пустые столбцы скрываются.
  function rawTable(g, doc, { from, to, cols: colsMap }) {
    from = Math.max(0, from); to = Math.min(g.rows, to, from + 400);
    const map = colsMap || (doc ? doc.cols : {});
    const byCol = {};
    for (const [f, c] of Object.entries(map)) byCol[c] = f;
    const labels = Object.fromEntries(P.FIELDS.map((f) => [f.key, f.label]));
    const visible = [];
    for (let c = 0; c < g.cols; c++) {
      let any = byCol[c] != null;
      for (let r = from; r < to && !any; r++) if (g.texts[r][c]) any = true;
      if (any) visible.push(c);
    }
    const hdr = doc ? [doc.headerRows[0] - 1, doc.headerRows[1] - 1] : [-1, -1];
    const itemRows = new Set(doc ? doc.items.map((it) => it.row - 1) : []);
    let html = `<table class="raw"><thead><tr><th></th>${visible.map((c) => `<th class="${byCol[c] ? 'mapped' : ''}">${P.colLetter(c)}${byCol[c] ? `<span class="tag">${labels[byCol[c]]}</span>` : ''}</th>`).join('')}</tr></thead><tbody>`;
    for (let r = from; r < to; r++) {
      const cls = r >= hdr[0] && r <= hdr[1] ? 'hdr' : itemRows.has(r) ? 'item' : '';
      html += `<tr class="${cls}"><td class="rn">${r + 1}</td>${visible.map((c) => `<td class="${byCol[c] != null ? 'mapped' : ''}" title="${esc(g.texts[r][c])}">${esc(g.texts[r][c])}</td>`).join('')}</tr>`;
    }
    return html + '</tbody></table>';
  }

  // ---------- выгрузка и справочник ----------
  function valueFor(key, it, doc) {
    switch (key) {
      case 'supplier': return doc.supplier;
      case 'docNumber': return doc.number;
      case 'docDate': return doc.date;
      case 'barcode': return it.barcode ? String(it.barcode) : '';
      default: return it[key] == null ? '' : it[key];
    }
  }

  const safeName = (s) => String(s).replace(/[\\/:*?"<>|\[\]]+/g, ' ').replace(/\s+/g, ' ').trim();

  function exportDocs(idxs) {
    const cfg = getExportCfg().filter((c) => c.on);
    if (!cfg.length) { toast('В настройках выгрузки не выбран ни один столбец', true); return; }
    const out = XLSX.utils.book_new();
    const used = new Set();
    let noBc = 0, fuzzy = 0, firstDoc;
    for (const i of idxs) {
      const doc = resolvedDoc(i);
      firstDoc = firstDoc || doc;
      const aoa = [cfg.map((c) => c.title)];
      for (const it of doc.items) {
        if (!it.barcode) noBc++;
        if (it.src === 'fuzzy') fuzzy++;
        aoa.push(cfg.map((c) => valueFor(c.key, it, doc)));
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = cfg.map((c) => ({ wch: c.key === 'name' ? 60 : c.key === 'supplier' ? 30 : c.key === 'barcode' ? 16 : 13 }));
      let name = safeName(doc.number || 'Накладная ' + (i + 1)).slice(0, 28) || 'Лист';
      while (used.has(name)) name = name.slice(0, 26) + '_' + used.size;
      used.add(name);
      XLSX.utils.book_append_sheet(out, ws, name);
      learnDoc(doc, true);
    }
    const fname = idxs.length === 1
      ? safeName(`Накладная ${firstDoc.number || ''} ${firstDoc.supplier || ''}`) + '.xlsx'
      : safeName(`Накладные ${firstDoc.supplier || ''} (${idxs.length})`) + '.xlsx';
    XLSX.writeFile(out, fname);
    let msg = 'Файл сохранён: ' + fname;
    if (noBc) msg += ` · без штрихкода: ${noBc}`;
    if (fuzzy) msg += ` · не проверено: ${fuzzy}`;
    toast(msg, false);
    render();
  }

  function learnDoc(doc, silent) {
    let n = 0;
    for (const it of doc.items) {
      if (!it.barcode || it.src === 'fuzzy') continue;
      const k = P.dictKey(doc.supKey, it.name);
      if (dictEntries[k] && dictEntries[k].b === it.barcode) continue;
      dictEntries[k] = { s: doc.supplier || '', n: it.name, b: it.barcode, t: Date.now() };
      n++;
    }
    if (n) saveDict();
    if (!silent) toast(n ? `В справочник добавлено ${n} ${plural(n, 'штрихкод', 'штрихкода', 'штрихкодов')}` : 'Все штрихкоды уже есть в справочнике');
    return n;
  }

  // ---------- настройки ----------
  const dlg = $('#dlg');
  let dlgTab = 'cols';
  let draft = null; // { slot, headerRow, cols }

  function openSettings(tab) {
    dlgTab = tab || 'cols';
    const slot = state.main ? 'main' : state.bc ? 'bc' : 'main';
    draft = makeDraft(slot);
    renderDialog();
    if (!dlg.open) dlg.showModal();
  }

  function makeDraft(slot) {
    const s = state[slot];
    const p = (s && s.profile) || {};
    return { slot, headerRow: p.headerRow || '', cols: { ...(p.cols || {}) } };
  }

  function draftResult() {
    const s = state[draft.slot];
    if (!s) return null;
    const opts = { cols: draft.cols };
    if (draft.headerRow) opts.headerRow = Number(draft.headerRow);
    if (draft.slot === 'bc') opts.barcodes = true;
    return P.parseWorkbook(s.wb, XLSX, opts);
  }

  function renderDialog() {
    $('#dlg-title').textContent = dlgTab === 'dict' ? 'Справочник штрихкодов' : 'Настройки';
    for (const b of $('#dlg-tabs').querySelectorAll('button')) b.classList.toggle('on', b.dataset.tab === dlgTab);
    const body = $('#dlg-body'), foot = $('#dlg-foot');
    if (dlgTab === 'cols') { body.innerHTML = colsTab(); foot.innerHTML = colsFoot(); }
    else if (dlgTab === 'export') { body.innerHTML = exportTab(); foot.innerHTML = `<button class="btn left" data-dact="export-reset">Вернуть по умолчанию</button><button class="btn primary" data-close>Готово</button>`; }
    else { body.innerHTML = dictTab(); foot.innerHTML = `<button class="btn left" data-dact="dict-clear">${icon('trash')}Очистить</button><button class="btn" data-dact="dict-import">Загрузить из файла</button><button class="btn" data-dact="dict-export">${icon('download')}Сохранить в файл</button><button class="btn primary" data-close>Готово</button>`; }
  }

  function colsTab() {
    const s = state[draft.slot];
    const slotSwitch = state.main && state.bc ? `<div class="seg" style="margin-bottom:14px">
        <button class="${draft.slot === 'main' ? 'on' : ''}" data-dslot="main">Накладная</button>
        <button class="${draft.slot === 'bc' ? 'on' : ''}" data-dslot="bc">Файл со штрихкодами</button></div>` : '';
    if (!s) return `<div class="note">${icon('info')}<div>Сначала загрузите накладную: здесь можно будет указать, из каких столбцов брать название, штрихкод, количество и цену. Настройки запоминаются для каждого поставщика.</div></div>`;
    const res = draftResult();
    const doc = res.docs[0];
    const sheet = res.sheets[doc ? doc.sheetIndex : 0];
    const g = sheet.grid;
    const headers = doc ? doc.headers : [];
    const auto = doc ? P.parseWorkbook(s.wb, XLSX, { headerRow: draft.headerRow ? Number(draft.headerRow) : undefined, barcodes: draft.slot === 'bc' }).docs[0] : null;
    const autoCols = auto ? auto.cols : {};
    const colOptions = (field) => {
      const opts = [];
      const a = autoCols[field];
      opts.push(`<option value="">Авто${a != null ? ' — ' + P.colLetter(a) + (headers[a] ? ': ' + esc(headers[a].slice(0, 40)) : '') : ' — не найден'}</option>`);
      opts.push(`<option value="-1" ${String(draft.cols[field]) === '-1' ? 'selected' : ''}>Не брать</option>`);
      const sampleRow = doc && doc.items[0] ? doc.items[0].row - 1 : -1;
      for (let c = 0; c < g.cols; c++) {
        const h = headers[c] || '';
        const sample = sampleRow >= 0 ? g.texts[sampleRow][c] : '';
        if (!h && !sample) continue;
        const label = P.colLetter(c) + ' — ' + (h ? h.slice(0, 40) : 'пример: ' + String(sample).slice(0, 30));
        opts.push(`<option value="${c}" ${String(draft.cols[field]) === String(c) ? 'selected' : ''}>${esc(label)}</option>`);
      }
      return opts.join('');
    };
    const fields = draft.slot === 'bc' ? P.FIELDS.filter((f) => ['name', 'barcode', 'code'].includes(f.key)) : P.FIELDS;
    const count = res.docs.reduce((a, d) => a + d.items.length, 0);
    const sum = P.round2(res.docs.reduce((a, d) => a + d.items.reduce((x, it) => x + (it.sum || 0), 0), 0));
    const supplier = doc ? doc.supplier || 'поставщик не найден' : '—';
    const from = doc ? doc.headerRows[0] - 1 : 0;
    return `${slotSwitch}
      <div class="note">${icon('info')}<div>Настройки сохраняются для поставщика <b>${esc(supplier)}</b>: в следующий раз его накладные распознаются так же. Оставьте «Авто», если столбец определён верно.</div></div>
      <div class="map-grid">
        <label><b>Строка заголовка таблицы</b><input class="inp" type="number" min="1" data-dhdr value="${esc(draft.headerRow)}" placeholder="Авто${doc ? ' — ' + doc.headerRows[0] : ''}"></label>
        ${fields.map((f) => `<label><b>${f.label}</b><select data-dcol="${f.key}">${colOptions(f.key)}</select></label>`).join('')}
      </div>
      <div class="row-inline" style="margin-bottom:10px">
        <span class="badge ${count ? 'b-ok' : 'b-bad'}">${count ? icon('check') : icon('alert')}Найдено позиций: ${count}</span>
        ${draft.slot === 'main' ? `<span class="badge b-muted">Сумма: ${fmtMoney(sum)} ₽</span>` : ''}
        ${res.docs.length > 1 ? `<span class="badge b-muted">Накладных в файле: ${res.docs.length}</span>` : ''}
      </div>
      <div class="mini-raw">${rawTable(g, doc, { from, to: from + 12 })}</div>`;
  }

  function colsFoot() {
    const s = state[draft.slot];
    if (!s) return `<button class="btn primary" data-close>Закрыть</button>`;
    return `<button class="btn left" data-dact="cols-reset">Сбросить на авто</button>
      <button class="btn" data-close>Отмена</button>
      <button class="btn primary" data-dact="cols-save">${icon('save')}Сохранить для поставщика</button>`;
  }

  function saveColsProfile(reset) {
    const s = state[draft.slot];
    const cols = {};
    for (const [f, c] of Object.entries(draft.cols)) if (c !== '' && c != null) cols[f] = Number(c);
    const profile = reset ? null : { headerRow: draft.headerRow ? Number(draft.headerRow) : undefined, cols };
    const all = LS.get(K_PROFILES, {});
    if (profile && (profile.headerRow || Object.keys(cols).length)) all[s.profileKey] = profile;
    else delete all[s.profileKey];
    LS.set(K_PROFILES, all);
    const parsed = parseSlot(draft.slot, s.wb, all[s.profileKey] || null);
    state[draft.slot] = { ...s, ...parsed, profileKey: s.profileKey };
    if (draft.slot === 'main') { state.docIdx = Math.min(state.docIdx, Math.max(0, parsed.result.docs.length - 1)); state.edits = {}; }
    else rebuildPool();
    dlg.close();
    toast(reset ? 'Настройки сброшены — столбцы определяются автоматически' : 'Настройки столбцов сохранены');
    render();
  }

  function exportTab() {
    const cfg = getExportCfg();
    return `<div class="note" style="margin-bottom:14px">${icon('info')}<div>Какие столбцы и в каком порядке попадут в итоговый Excel. Названия столбцов можно переименовать под шаблон Контур.Маркета.</div></div>
      <div class="exp-list">${cfg.map((c, i) => `<div class="exp-row ${c.on ? '' : 'off'}">
        <input type="checkbox" data-eon="${i}" ${c.on ? 'checked' : ''}>
        <span class="k">${EXPORT_LABELS[c.key]}</span>
        <input class="inp" data-etitle="${i}" value="${esc(c.title)}">
        <button class="btn ghost icon-btn" data-emove="${i}:-1" ${i === 0 ? 'disabled' : ''} title="Выше">${icon('up')}</button>
        <button class="btn ghost icon-btn" data-emove="${i}:1" ${i === cfg.length - 1 ? 'disabled' : ''} title="Ниже">${icon('down')}</button>
      </div>`).join('')}</div>`;
  }

  let dictQ = '';
  function dictTab() {
    const list = Object.entries(dictEntries).map(([k, v]) => ({ k, ...v }));
    const q = dictQ.trim().toLowerCase();
    const rows = list.filter((e) => !q || e.n.toLowerCase().includes(q) || e.b.includes(q) || (e.s || '').toLowerCase().includes(q))
      .sort((a, b) => (b.t || 0) - (a.t || 0)).slice(0, 500);
    return `<div class="note" style="margin-bottom:14px">${icon('info')}<div>Приложение запоминает, какой штрихкод у какого названия, при каждой выгрузке и по кнопке «Запомнить штрихкоды». Когда поставщик пришлёт накладную без штрихкодов, они подставятся отсюда. Сейчас в справочнике <b>${list.length}</b> ${plural(list.length, 'запись', 'записи', 'записей')}.</div></div>
      <label class="search" style="margin-bottom:12px">${icon('search')}<input data-dq placeholder="Поиск" value="${esc(dictQ)}"></label>
      ${rows.length ? `<div class="mini-raw" style="max-height:420px"><table class="dict-table"><thead><tr><th>Поставщик</th><th>Наименование</th><th>Штрихкод</th><th></th></tr></thead><tbody>
        ${rows.map((e) => `<tr><td>${esc(e.s || '—')}</td><td>${esc(e.n)}</td><td style="font-variant-numeric:tabular-nums">${esc(e.b)}</td><td><button class="btn ghost icon-btn" data-ddel="${esc(e.k)}" title="Удалить">${icon('trash')}</button></td></tr>`).join('')}
      </tbody></table></div>` : `<div class="table-empty">Справочник пока пуст</div>`}`;
  }

  // ---------- события ----------
  const fileInput = $('#file-input');
  fileInput.addEventListener('change', () => {
    const f = fileInput.files[0];
    fileInput.value = '';
    if (f) loadFile(state.pickTarget, f);
  });

  document.addEventListener('click', (e) => {
    const t = e.target.closest('[data-pick],[data-clear],[data-act],[data-doc],[data-filter],[data-view],[data-confirm]');
    if (!t || dlg.contains(t)) return;
    if (t.dataset.pick) { state.pickTarget = t.dataset.pick; fileInput.click(); }
    else if (t.dataset.clear) clearSlot(t.dataset.clear);
    else if (t.dataset.doc) { state.docIdx = Number(t.dataset.doc); render(); }
    else if (t.dataset.filter) { state.filter = t.dataset.filter; renderContent(); }
    else if (t.dataset.view) { state.view = t.dataset.view; renderContent(); }
    else if (t.dataset.confirm) {
      const i = Number(t.dataset.confirm);
      const ed = (state.edits[state.docIdx] = state.edits[state.docIdx] || {});
      ed[i] = { ...(ed[i] || {}), confirmed: true };
      render();
    } else if (t.dataset.act === 'export') exportDocs([state.docIdx]);
    else if (t.dataset.act === 'export-all') exportDocs(state.main.result.docs.map((_, i) => i));
    else if (t.dataset.act === 'learn') { learnDoc(resolvedDoc(state.docIdx)); render(); }
    else if (t.dataset.act === 'settings') openSettings('cols');
    else if (t.dataset.act === 'demo') loadDemo();
  });

  $('#btn-settings').addEventListener('click', () => openSettings('cols'));
  $('#btn-dict').addEventListener('click', () => openSettings('dict'));

  const content = $('#content');
  content.addEventListener('input', (e) => {
    if (e.target.id === 'q') {
      state.q = e.target.value;
      $('#table-wrap').innerHTML = itemsTable(resolvedDoc(state.docIdx));
    }
  });
  content.addEventListener('change', (e) => {
    const t = e.target;
    if (t.dataset.bc != null) {
      const i = Number(t.dataset.bc);
      const ed = (state.edits[state.docIdx] = state.edits[state.docIdx] || {});
      ed[i] = { barcode: t.value.replace(/\s+/g, '') };
      const next = state.focusNext ? i : null;
      state.focusNext = false;
      render();
      if (next != null) {
        const inputs = [...content.querySelectorAll('.bc-input')];
        const pos = inputs.findIndex((x) => Number(x.dataset.bc) === next);
        if (inputs[pos + 1]) inputs[pos + 1].focus();
      }
    } else if (t.id === 'supplier-input') {
      state.overrides[state.docIdx] = { supplier: t.value.trim() };
      renderSummary();
    }
  });
  content.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && e.target.dataset.bc != null) { state.focusNext = true; e.target.blur(); }
  });

  // перетаскивание файлов
  let dragDepth = 0;
  const dropSlotAt = (el) => {
    const z = el && el.closest && el.closest('[data-pick],[data-dropzone]');
    return z ? z.dataset.pick || z.dataset.dropzone : null;
  };
  document.addEventListener('dragenter', (e) => { e.preventDefault(); dragDepth++; });
  document.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; document.querySelectorAll('.drop.over').forEach((x) => x.classList.remove('over')); } });
  document.addEventListener('dragover', (e) => {
    e.preventDefault();
    document.querySelectorAll('.drop.over').forEach((x) => x.classList.remove('over'));
    const z = e.target.closest && e.target.closest('.drop');
    if (z) z.classList.add('over');
  });
  document.addEventListener('drop', (e) => {
    e.preventDefault();
    dragDepth = 0;
    document.querySelectorAll('.drop.over').forEach((x) => x.classList.remove('over'));
    const files = [...(e.dataTransfer && e.dataTransfer.files ? e.dataTransfer.files : [])];
    if (!files.length) return;
    const slot = dropSlotAt(e.target);
    if (slot) { loadFile(slot, files[0]); return; }
    // брошено мимо: первый файл — накладная, второй — штрихкоды
    loadFile('main', files[0]).then(() => { if (files[1]) loadFile('bc', files[1]); });
  });

  // диалог
  dlg.addEventListener('click', (e) => {
    if (e.target === dlg) { dlg.close(); return; }
    const t = e.target.closest('[data-close],[data-tab],[data-dslot],[data-dact],[data-emove],[data-ddel]');
    if (!t) return;
    if (t.hasAttribute('data-close')) dlg.close();
    else if (t.dataset.tab) { dlgTab = t.dataset.tab; renderDialog(); }
    else if (t.dataset.dslot) { draft = makeDraft(t.dataset.dslot); renderDialog(); }
    else if (t.dataset.emove) {
      const [i, d] = t.dataset.emove.split(':').map(Number);
      const cfg = getExportCfg();
      [cfg[i], cfg[i + d]] = [cfg[i + d], cfg[i]];
      LS.set(K_EXPORT, cfg); renderDialog();
    } else if (t.dataset.ddel) {
      delete dictEntries[t.dataset.ddel]; saveDict(); renderDialog(); render();
    } else switch (t.dataset.dact) {
      case 'cols-save': saveColsProfile(false); break;
      case 'cols-reset': saveColsProfile(true); break;
      case 'export-reset': LS.set(K_EXPORT, DEFAULT_EXPORT); renderDialog(); break;
      case 'dict-clear':
        if (confirm('Удалить все записи справочника? Это нельзя отменить.')) { dictEntries = {}; saveDict(); renderDialog(); render(); }
        break;
      case 'dict-export': {
        const blob = new Blob([JSON.stringify(dictEntries, null, 1)], { type: 'application/json' });
        const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: 'справочник-штрихкодов.json' });
        a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000);
        break;
      }
      case 'dict-import': $('#json-input').click(); break;
    }
  });
  dlg.addEventListener('change', (e) => {
    const t = e.target;
    if (t.dataset.dcol) { draft.cols[t.dataset.dcol] = t.value; renderDialog(); }
    else if (t.hasAttribute('data-dhdr')) { draft.headerRow = t.value; renderDialog(); }
    else if (t.dataset.eon != null) { const cfg = getExportCfg(); cfg[t.dataset.eon].on = t.checked; LS.set(K_EXPORT, cfg); renderDialog(); }
    else if (t.dataset.etitle != null) { const cfg = getExportCfg(); cfg[t.dataset.etitle].title = t.value.trim() || EXPORT_LABELS[cfg[t.dataset.etitle].key]; LS.set(K_EXPORT, cfg); }
  });
  dlg.addEventListener('input', (e) => {
    if (e.target.hasAttribute('data-dq')) {
      dictQ = e.target.value;
      const pos = e.target.selectionStart;
      renderDialog();
      const inp = $('[data-dq]', dlg); inp.focus(); inp.setSelectionRange(pos, pos);
    }
  });
  $('#json-input').addEventListener('change', async (e) => {
    const f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    try {
      const data = JSON.parse(await f.text());
      let n = 0;
      for (const [k, v] of Object.entries(data)) if (v && typeof v.b === 'string' && typeof v.n === 'string') { dictEntries[k] = v; n++; }
      saveDict(); renderDialog(); render();
      toast(`Загружено записей: ${n}`);
    } catch { toast('Не удалось прочитать файл справочника', true); }
  });

  // для автотестов
  window.__app = { state, loadFile, exportDocs, resolvedDoc };

  render();
})();
