import fs from 'node:fs/promises';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';

async function readSnapshot(path, date) {
  const buffer = await fs.readFile(path);
  const zip = await JSZip.loadAsync(buffer);
  const entry = Object.values(zip.files).find((item) => !item.dir && /\.(csv|xlsx|xls)$/i.test(item.name));
  if (!entry) throw new Error(`No spreadsheet in ${path}`);
  const data = await entry.async('uint8array');
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '', raw: true });
  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell).replace(/^\uFEFF/, '').trim() === 'ชื่อยา'));
  const headers = rows[headerIndex].map((cell) => String(cell).replace(/^\uFEFF/, '').trim());
  const col = (label) => headers.findIndex((header) => header.includes(label));
  const codeCol = col('รหัสยา'), nameCol = col('ชื่อยา'), categoryCol = col('หมวดหมู่'), subcategoryCol = col('หมวดหมู่ย่อย'), unitCol = col('หน่วย'), minimumCol = col('ยอดสต็อกขั้นต่ำ'), lotCol = col('เลขล็อตผู้ผลิต'), expiryCol = col('วันหมดอายุ'), stockCol = col('ยอดคงคลัง ณ วันที่เลือก');
  const grouped = new Map();
  for (const row of rows.slice(headerIndex + 1)) {
    const name = String(row[nameCol] || '').trim();
    if (!name) continue;
    const code = String(row[codeCol] || name).trim();
    const quantity = Number(row[stockCol]) || 0;
    const lot = String(row[lotCol] || '').trim();
    const expiry = row[expiryCol] instanceof Date ? row[expiryCol].toISOString().slice(0, 10) : String(row[expiryCol] || '').trim();
    const category = [row[categoryCol], row[subcategoryCol]].map((x) => String(x || '').trim()).filter(Boolean).join(' / ') || 'อื่น ๆ';
    const current = grouped.get(code);
    if (current) {
      current.stock += quantity;
      if (lot && !current.lots.includes(lot)) current.lots.push(lot);
      if (expiry && !current.expiries.includes(expiry)) current.expiries.push(expiry);
    } else grouped.set(code, { code, name, category, unit: String(row[unitCol] || 'อัน').trim(), minimum: Number(row[minimumCol]) || 0, stock: quantity, lots: lot ? [lot] : [], expiries: expiry ? [expiry] : [] });
  }
  const items = [...grouped.values()].map(({ lots, expiries, minimum, ...item }) => item);
  const products = [...grouped.values()].map((item) => ({ id: `drug-${item.code}`, code: item.code, name: item.name, category: item.category, unit: item.unit, minimum: item.minimum, active: true, stock: item.stock, expiry: item.expiries.join(', '), note: item.lots.length ? `Lot ${item.lots.join(', ')}` : '' }));
  return { snapshot: { id: date, date, fileName: path.split(/[\\/]/).at(-1), items }, products };
}

const [openingPath, currentPath] = process.argv.slice(2);
const opening = await readSnapshot(openingPath, '2026-08-31');
const current = await readSnapshot(currentPath, '2026-09-06');
process.stdout.write(JSON.stringify({ products: current.products, snapshots: [opening.snapshot, current.snapshot] }));
