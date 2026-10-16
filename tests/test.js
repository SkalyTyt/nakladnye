// Проверка распознавания на примерах: node tests/test.js
const path = require('path');
const fs = require('fs');
const XLSX = require('../vendor/xlsx.full.min.js');
const P = require('../parser.js');

const dir = path.join(__dirname, '..', 'samples');
let failed = 0;
const check = (cond, msg) => { console.log((cond ? '  ✓ ' : '  ✗ ') + msg); if (!cond) failed++; };
const read = (f) => XLSX.read(fs.readFileSync(path.join(dir, f)), { type: 'buffer' });
const letters = (cols) => Object.entries(cols).map(([f, c]) => f + '=' + P.colLetter(c)).join(' ');

for (const f of ['torg12_itkis.xlsx', 'parfum_kosmetik.xlsx'].filter((f) => fs.existsSync(path.join(dir, f)))) {
  console.log('\n=== ' + f);
  const res = P.parseWorkbook(read(f), XLSX);
  check(res.docs.length > 0, 'найдено документов: ' + res.docs.length);
  for (const d of res.docs) {
    const sum = P.round2(d.items.reduce((s, it) => s + (it.sum || 0), 0));
    console.log(`  № ${d.number} от ${d.date} | ${d.supplier} (ИНН ${d.inn || '—'}) | позиций ${d.items.length} | сумма ${sum} / итого в файле ${d.fileTotal}`);
    console.log('    столбцы: ' + letters(d.cols));
    check(d.fileTotal == null || Math.abs(sum - d.fileTotal) < 0.05, 'сумма сходится с итогом');
    check(d.items.every((it) => it.name && it.qty > 0 && it.price > 0), 'у всех позиций есть название, кол-во и цена');
    check(d.items.every((it) => /^\d{8,14}$/.test(it.barcode)), 'у всех позиций штрихкод');
  }
}

// Объединение: файл с ценами без штрихкодов + файл с названиями и штрихкодами
if (fs.existsSync(path.join(dir, 'parfum_kosmetik.xlsx'))) {
console.log('\n=== объединение двух файлов');
const src = P.parseWorkbook(read('parfum_kosmetik.xlsx'), XLSX).docs[0].items;
const priced = src.map((it) => ({ ...it, barcode: '', code: '' }));
const variants = (n, i) => (i % 3 === 0 ? n.toUpperCase() : i % 3 === 1 ? n.replace(/\s+/g, '  ') + ' ' : n.replace(/,/g, ' ,'));
const barcodes = src.map((it, i) => ({ name: variants(it.name, i), barcode: it.barcode, code: '' })).reverse();
const out = P.resolveItems(priced, { pool: P.buildPool(barcodes) });
const right = out.filter((it, i) => it.barcode === src[i].barcode).length;
const wrong = out.filter((it, i) => it.barcode && it.barcode !== src[i].barcode).length;
check(right === src.length && wrong === 0, `подставлено верно ${right} из ${src.length}, ошибок ${wrong}`);
}

// Похожие названия, но разные цифры — не путать
const pool = P.buildPool([{ name: 'Краска ФАРА-нью 501 черный', barcode: '111' }]);
check(P.matchInPool({ name: 'Краска ФАРА-нью 502 черный' }, pool) === null, '«501» не подставляется к «502»');
check(P.matchInPool({ name: 'Краска ФАРА нью 501 чёрный ' }, pool) !== null, 'небольшие отличия в написании — находится');

// Справочник
const idx = P.indexDict({ [P.dictKey('7700000000', 'Мыло ДЕГТЯРНОЕ 140гр')]: { b: '4601313011880' } });
check(P.dictLookup(idx, '7700000000', 'мыло  дегтярное 140гр') === '4601313011880', 'справочник находит штрихкод по названию');

// Счёт на оплату: номер, дата прописью, поставщик, несколько штрихкодов в ячейке
console.log('\n=== счёт на оплату (демо)');
{
  const ddir = path.join(__dirname, '..', 'demo');
  const d = P.parseWorkbook(XLSX.read(fs.readFileSync(path.join(ddir, 'schet_na_oplatu_primer.xlsx'))), XLSX).docs[0];
  check(d && d.number === '1513', 'номер счёта: ' + (d && d.number));
  check(d && d.date === '15.10.2026', 'дата: ' + (d && d.date));
  check(d && d.supplier === 'ООО "Демо Поставщик"' && d.inn === '7700000000', 'поставщик: ' + (d && d.supplier) + ', ИНН ' + (d && d.inn));
  check(d && /^\d{13}$/.test(d.items[1].barcode), 'из двух штрихкодов в ячейке взят первый: ' + (d && d.items[1].barcode));
  const sum = P.round2(d.items.reduce((s, it) => s + it.sum, 0));
  check(d.items.length === 12 && Math.abs(sum - d.fileTotal) < 0.05, `позиций ${d.items.length}, сумма ${sum} / итого ${d.fileTotal}`);
  const bc = P.parseWorkbook(XLSX.read(fs.readFileSync(path.join(ddir, 'shtrihkody_primer.xlsx'))), XLSX, { barcodes: true }).docs[0];
  const r = P.resolveItems(d.items, { pool: P.buildPool(bc.items) });
  check(r.every((it) => it.barcode), 'после второго файла штрихкоды есть у всех: ' + r.filter((it) => it.barcode).length + '/' + r.length);
}

console.log(failed ? `\nОшибок: ${failed}` : '\nВсе проверки пройдены');
process.exit(failed ? 1 : 0);
