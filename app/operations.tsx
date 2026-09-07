'use client';
import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { useSharedStored } from '@/hooks/use-shared-stored';
import { ThaiDateInput } from '@/components/thai-date-input';
import {
  ArrowLeftRight,
  Boxes,
  CheckCircle2,
  ClipboardCheck,
  ClipboardPlus,
  FileDown,
  FileUp,
  History,
  PackageMinus,
  Plus,
  Printer,
  Save,
  Search,
  Settings,
  Tags,
} from 'lucide-react';

type RecordRow = {
  id: string;
  date: string;
  branch: string;
  product: string;
  qty: number;
  unit: string;
  party: string;
  note: string;
  status: string;
  type: string;
  lot?: string;
  mfg?: string;
  exp?: string;
  attachment?: string;
};
type ReceiveDraft = { code: string; product: string; qty: number; unit: string; lot: string; mfg: string; exp: string };
type Product = {
  id: string;
  code: string;
  name: string;
  category: string;
  unit: string;
  minimum: number;
  active: boolean;
  stock: number;
  expiry?: string;
  note?: string;
};
type StockSnapshotItem = { code: string; name: string; category: string; unit: string; stock: number };
type StockSnapshot = { id: string; date: string; fileName: string; items: StockSnapshotItem[] };
const today = () => new Date().toISOString().slice(0, 10);
const id = () => crypto.randomUUID?.() || Math.random().toString(36).slice(2);
const branches = [
  'MICHIKO สาขา Emsphere',
  'MICHIKO สาขาพหลโยธิน',
];
const initialProducts: Product[] = [];
const accountingTemplate = [
  ['Restylane', 'Defyne|Volyme|Kysses|Vital|Vital light|Lyft|Classic'],
  ['Flore', 'Max|Max 1400|N|S|Aqua-S'],
  ['Hyabell', 'Basic|Lip|Deep|Ultra|Meso'],
  ['Volifil', 'Classic|Deep|Sub - Q'],
  ['Belotero', 'Intense|Volume|Sculptra (ขวด)|Radiesse (กล่อง)'],
  ['ยาสลายฟิลเลอร์', 'Liporase'],
  ['Botox', 'Dysport (S.U.)|Allergan (Unit)|Bienox (Unit)|Xeomin (Unit)'],
  ['Meso fat', 'Lipolytic (ขวด)|Red Lava (ขวด)|Neobella (ขวด)'],
  ['Meso Glitter', 'Glitter (ml.)|Derma Care (ml.)|Transamine (ml.)|Depigment (ml.)'],
  ['Cell Detox', 'Bio Centa (Amp)|REPLAY Placenta Skin (ขวด)|REJU Growth Factors (ขวด)'],
  ['วิตามิน', 'B complex (Amp.)|B6 pyridoxine (Amp.)|B12 Methylcobal (Amp.)|B12 (Amp.)|Vitamin C (Amp.)|7.5%Sodium Bicarbonate (Amp.)|Acetin 300 mg (Amp.)|50% Magnesium Sulfate (Amp.)|Lipotocin ALA 300 mg (Amp.)|Glutathione 600 mg (Amp.)|Glutathione 1200 mg (Amp.)|Calcium Gluconate(Amp.)|Ubiquinone (Amp.)|Folic (10mg/ml) (30 ml/ขวด) (ml)|Dexpanthenol (250mg/ml) (30 ml/ขวด) (ml)|Zinc sulfate (10mg/ml) (30 ml/ขวด) (ml)|Inositol (50mg/ml (30 ml/ขวด) (ml)|MTE-5(10mg/ml) (10 ml/ขวด) (ml)|Biotin inj. (Amp)|Vitamin D3 (300,000 iU/1.5ml) (Amp)|Cerebrolysin IV|Selenium (200mcg/ml) (30ml/ขวด) (ml)|Potassium Chlorride (Amp)|EDTA (3g/10ml) (ml)|InfiNADi NAD+ IV|Fluimucil 300 mg|NAD+Up 250 mg (Vial)|Polilevo|NeoAmiyu (ถุง)|L-carnitine 1000mg/amp (Amp)|Beecalfor Calcium inj. (10x3ml) (Amp)|Beecalfor Vitamin B complex inj. (10x2 ml) (Amp)|Vitamin K3 inj. (Amp)|L-Arginine inj. 0.6 mg/Vial (Vial)'],
  ['เวชภัณฑ์', 'Three way Stop Clock|Blunt Canula no.22G|Blunt Canula no.25G|Nipro 18G 1" (100 อัน/กล่อง)|Nipro 21G 1" (100 อัน/กล่อง)|Nipro 25G 1" (100 อัน/กล่อง)|Nipro 27G 1" (100 อัน/กล่อง)|Nipro 30G 1/2" (100 อัน/กล่อง)|Nipro Safelet Cath 24 G3/4" (50อัน/กล่อง)|Nipro หัวล็อค 5 ml|เข็มปีกผีเสื้อ24G (50 อัน/กล่อง)|เข็มปีกผีเสื้อ25G (50 อัน/กล่อง)|Syringe 1 CC (100 อัน/กล่อง)|Syringe 3 CC (100 อัน/กล่อง)|Syringe 5 CC (100 อัน/กล่อง)|Syringe 10 CC (100 อัน/กล่อง)|Syringe 20 CC (100 อัน/กล่อง)|Insulin Syringe 100 u|ชุดสายให้น้ำเกลือ BMI (IV Set)|ถุงมือ size S (ไม่มีแป้ง)|ถุงมือ size M|สำลีแผ่น (500g /ห่อ)|สำลีก้อน 0.35g (450g)|ไม้พันสำลี Size M|Cotton bud|ผ้าก๊อซพับ 3"x3"|ผ้าก๊อซ Sterile 4x4'],
  ['น้ำเกลือ', 'NSS 5 ml|NSS 100 ml|NSS 250 ml|NSS 500 ml|D5W 100 ml|D5W 250 ml|D5W 500 ml|Steri Water 10 ml/amp'],
  ['วิตามิน', 'Skin Glow|Liver & Kidney Detox|Stress Relax|Melatonin IX/RX|Digestzyme|RxOmega-3 Ultra Strength|Glycine|Super Synbiotic 18 strains 60 billion CFU|วิตามิน Strong Immunity Essential Formula|วิตามิน Strong Immunity Intensive Formula'],
  ['ครีม', 'Diamond White TM4 Cream|Peptide Liked Anti-Aging Serum|UV Sunscreen Gel SPF60 PA+++ (50ml)|Cleansing Water - Doctor Formula (100ml)|Cleansing Water - Doctor Formula (500ml)|Milky Soap - Doctor Formula (50ml)|Milky Soap - Doctor Formula (150ml)|Aura Skin Whitening Lotion - Doctor Formula|Aloe Vera Cream|Plankton Essence mini|Plankton Essence|Cica Moist Cream 30G|Anti Melasma Spot Cream'],
  ['ยาสิว', 'Arnicare Bruise|Arnicare 30C|Arnica Gel & Oral|Terramycin ointment|Anti P.Acne|Anti Bac & Yeast|Anti Comedones|Anti Acne Spot Gel|DARK Spot Reduction|Eczema Relief|Sebcare|DARK Spot Reduction X2'],
  ['หัวเครื่องยกกระชับ', 'Ulthera 3.0 MM|Ulthera 4.5 MM|Ultraformer III Cartridge 1.5 MM|Ultraformer III Cartridge 2.0 MM|Ultraformer III Cartridge 3.0 MM|Ultraformer III Cartridge 4.5 MM|Xerf Effector 60 (600)|Xerf Grounding Plate|Xerf Gas LUT-ICD'],
] as const;
export function OperationalPage({
  page,
  notify,
}: {
  page: string;
  notify: (s: string) => void;
}) {
  const [dailyCases] = useSharedStored<Array<{ date: string; items: Array<{ name: string; qty: number }> }>>('michiko-stock-cases', []);
  const [records, setRecords] = useSharedStored<RecordRow[]>(
    'michiko-operations',
    [],
  );
  const [products, setProducts] = useSharedStored<Product[]>(
    'michiko-products',
    initialProducts,
  );
  if (page === 'Stock คงคลัง')
    return (
      <Inventory
        products={products}
        setProducts={setProducts}
        notify={notify}
      />
    );
  if (page === 'ทะเบียนสินค้า')
    return (
      <Products products={products} setProducts={setProducts} notify={notify} />
    );
  if (page === 'รายงาน')
    return <Reports records={records} products={products} notify={notify} dailyCases={dailyCases} />;
  if (page === 'ประวัติการเคลื่อนไหว') return <Movements records={records} />;
  if (page === 'ตั้งค่า') return <Preferences notify={notify} />;
  return (
    <TransactionPage
      page={page}
      records={records}
      setRecords={setRecords}
      products={products}
      setProducts={setProducts}
      notify={notify}
    />
  );
}

const configs: Record<
  string,
  {
    icon: typeof Boxes;
    title: string;
    sub: string;
    type: string;
    party: string;
    status: string;
  }
> = {
  รับเข้า: {
    icon: ClipboardPlus,
    title: 'รับ Stock เข้า',
    sub: 'บันทึกสินค้า Lot และวันหมดอายุ พร้อมสร้างเลขเอกสารอัตโนมัติ',
    type: 'รับเข้า',
    party: 'Supplier / ที่มา',
    status: 'รับเข้าแล้ว',
  },
  'โอนย้าย / ยืมคืน': {
    icon: ArrowLeftRight,
    title: 'รายการระหว่างสาขา',
    sub: 'โอนย้าย ยืม คืน และซื้อ–ขาย Stock ระหว่าง 2 สาขาในหน้าเดียว',
    type: 'โอนย้าย',
    party: 'สาขาปลายทาง',
    status: 'รอรับ',
  },
  เบิกออก: {
    icon: PackageMinus,
    title: 'เบิก Stock ออก',
    sub: 'สำหรับใช้ภายใน Tester ของเสีย หมดอายุ และการเบิกทั่วไป',
    type: 'เบิกออก',
    party: 'ประเภทการเบิก',
    status: 'บันทึกแล้ว',
  },
  'ตรวจนับ Stock': {
    icon: ClipboardCheck,
    title: 'ตรวจนับ Stock',
    sub: 'เปรียบเทียบยอดตามระบบกับยอดนับจริงและบันทึกสาเหตุส่วนต่าง',
    type: 'ตรวจนับ',
    party: 'สาเหตุส่วนต่าง',
    status: 'ตรวจแล้ว',
  },
};
function TransactionPage({
  page,
  records,
  setRecords,
  products,
  setProducts,
  notify,
}: {
  page: string;
  records: RecordRow[];
  setRecords: (v: RecordRow[] | ((p: RecordRow[]) => RecordRow[])) => void;
  products: Product[];
  setProducts: (v: Product[] | ((p: Product[]) => Product[])) => void;
  notify: (s: string) => void;
}) {
  const cfg = configs[page] || configs['รับเข้า'];
  const Icon = cfg.icon;
  const [open, setOpen] = useState(false);
  const [form, setForm] = useState({
    date: today(),
    branch: 'MICHIKO สาขา Emsphere',
    product: products[0]?.name || '',
    qty: 1,
    unit: products[0]?.unit || 'อัน',
    party:
      page === 'โอนย้าย / ยืมคืน'
        ? 'MICHIKO สาขาพหลโยธิน'
        : '',
    note: '',
    lot: '',
    mfg: '',
    exp: '',
    attachment: '',
    movementType: cfg.type,
  });
  const [receiveDrafts, setReceiveDrafts] = useState<ReceiveDraft[]>([]);
  const receiveFileRef = useRef<HTMLInputElement>(null);
  const branchTradeTypes = ['โอนย้าย', 'ยืม', 'คืน', 'ซื้อจากสาขา', 'ขายให้สาขา'];
  const rows = records.filter((r) => page === 'โอนย้าย / ยืมคืน' ? branchTradeTypes.includes(r.type) : r.type === cfg.type);
  const save = () => {
    if (page === 'รับเข้า' && receiveDrafts.length) {
      const hasPo = /\bPO\b/i.test(form.note);
      if (hasPo) {
        setProducts((current) => {
          const next = [...current];
          receiveDrafts.forEach((item) => {
            const productIndex = next.findIndex((product) =>
              product.code === item.code || normalizeName(product.name) === normalizeName(item.product),
            );
            if (productIndex >= 0) {
              const product = next[productIndex];
              next[productIndex] = {
                ...product,
                stock: product.stock + item.qty,
                expiry: item.exp || product.expiry,
              };
            } else {
              next.push({
                id: id(),
                code: item.code || `AUTO${String(next.length + 1).padStart(4, '0')}`,
                name: item.product,
                category: inferProductCategory(item.product),
                unit: item.unit,
                minimum: 0,
                active: true,
                stock: item.qty,
                expiry: item.exp,
                note: `เพิ่มอัตโนมัติจาก ${form.note}`,
              });
            }
          });
          return next;
        });
      }
      setRecords((current) => [
        ...receiveDrafts.map((item) => ({
          id: id(), date: form.date, branch: form.branch, party: form.party,
          note: form.note, status: 'บันทึกแล้ว', type: 'รับเข้า',
          attachment: form.attachment, ...item,
        })),
        ...current,
      ]);
      setReceiveDrafts([]);
      setOpen(false);
      notify(`บันทึกรับเข้า ${receiveDrafts.length} รายการ${hasPo ? ' และอัปเดตคลังจาก PO แล้ว' : 'จากเอกสารแล้ว'}`);
      return;
    }
    if (!form.product || form.qty <= 0) return;
    if (
      page === 'โอนย้าย / ยืมคืน' &&
      form.branch === form.party
    ) {
      notify('กรุณาเลือกสาขาต้นทางและปลายทางให้ต่างกัน');
      return;
    }
    setRecords((p) => [
      { id: id(), ...form, type: form.movementType, status: form.movementType === 'โอนย้าย' ? 'รอรับ' : form.movementType === 'ยืม' ? 'รอคืน' : form.movementType === 'คืน' ? 'คืนแล้ว' : 'บันทึกแล้ว' },
      ...p,
    ]);
    setOpen(false);
    notify(`บันทึก${cfg.title}เรียบร้อยแล้ว`);
  };
  return (
    <>
      <PageHead
        icon={<Icon />}
        title={cfg.title}
        sub={cfg.sub}
        action={() => setOpen(!open)}
        actionText={open ? 'ปิดแบบฟอร์ม' : '+ ทำรายการใหม่'}
      />
      {open && (
        <section className="ops-form panel">
          {page === 'รับเข้า' && <div className="receive-document-import">
            <input ref={receiveFileRef} type="file" accept=".pdf" hidden onChange={async (event) => {
              const file = event.target.files?.[0]; if (!file) return;
              try {
                const parsed = await parseJeraReceivePdf(file, products);
                if (!parsed.items.length) throw new Error('no items');
                setReceiveDrafts(parsed.items);
                const first = parsed.items[0];
                setForm((current) => ({ ...current, product: first.product, qty: first.qty, unit: first.unit, lot: first.lot, mfg: first.mfg, exp: first.exp, party: parsed.supplier || current.party, note: parsed.po ? `PO ${parsed.po}` : current.note, attachment: file.name }));
                notify(`อ่านเอกสารสำเร็จ ${parsed.items.length} รายการ`);
              } catch { notify('อ่านเอกสารรับเข้าไม่ได้ กรุณาใช้ PDF ใบรับยาเข้าคลังจาก JERA'); }
              event.target.value = '';
            }}/>
            <button type="button" onClick={() => receiveFileRef.current?.click()}><FileUp size={16}/> แนบ PDF ใบรับยาเข้าคลัง</button>
            <span>{form.attachment ? `${form.attachment} · พบ ${receiveDrafts.length} รายการ` : 'ระบบจะกรอกสินค้า Lot วันผลิต และวันหมดอายุให้อัตโนมัติ'}</span>
          </div>}
          <div className="form-grid">
            {page === 'โอนย้าย / ยืมคืน' && <Field label="ประเภทรายการ"><select value={form.movementType} onChange={(e) => setForm({ ...form, movementType: e.target.value })}><option>โอนย้าย</option><option>ยืม</option><option>คืน</option><option>ซื้อจากสาขา</option><option>ขายให้สาขา</option></select></Field>}
            <Field label="วันที่">
              <ThaiDateInput value={form.date} onChange={(date) => setForm({ ...form, date })} />
            </Field>
            <Field
              label={
                page === 'โอนย้าย / ยืมคืน' ? 'สาขาต้นทาง' : 'สาขา'
              }
            >
              <select
                value={form.branch}
                onChange={(e) => setForm({ ...form, branch: e.target.value })}
              >
                {branches.map((branch) => (
                  <option key={branch}>{branch}</option>
                ))}
              </select>
            </Field>
            <Field label="รายการสินค้า">
              <input list="transaction-products"
                value={form.product}
                onChange={(e) => {
                  const p = products.find((x) => x.name === e.target.value || `${x.code} — ${x.name}` === e.target.value);
                  setForm({
                    ...form,
                    product: p?.name || e.target.value,
                    unit: p?.unit || form.unit,
                  });
                }}
                placeholder="พิมพ์รหัสหรือชื่อสินค้า" />
              <datalist id="transaction-products">{products.map((p) => <option key={p.id} value={`${p.code} — ${p.name}`} />)}</datalist>
            </Field>
            <Field label="จำนวน">
              <input
                type="number"
                min="0.01"
                step="0.01"
                value={form.qty}
                onChange={(e) =>
                  setForm({ ...form, qty: Number(e.target.value) })
                }
              />
            </Field>
            <Field label="หน่วย">
              <input
                value={form.unit}
                onChange={(e) => setForm({ ...form, unit: e.target.value })}
              />
            </Field>
            <Field label={cfg.party}>
              {page === 'โอนย้าย / ยืมคืน' ? (
                <select
                  value={form.party}
                  onChange={(e) => setForm({ ...form, party: e.target.value })}
                >
                  {branches.map((branch) => (
                    <option key={branch}>{branch}</option>
                  ))}
                </select>
              ) : (
                <input
                  value={form.party}
                  onChange={(e) => setForm({ ...form, party: e.target.value })}
                  placeholder={cfg.party}
                />
              )}
            </Field>
            {page === 'รับเข้า' && (
              <>
                <Field label="Lot No.">
                  <input
                    value={form.lot}
                    onChange={(e) => setForm({ ...form, lot: e.target.value })}
                  />
                </Field>
                <Field label="วันผลิต">
                  <ThaiDateInput value={form.mfg} onChange={(mfg) => setForm({ ...form, mfg })} />
                </Field>
                <Field label="วันหมดอายุ">
                  <ThaiDateInput value={form.exp} onChange={(exp) => setForm({ ...form, exp })} />
                </Field>
              </>
            )}
            <Field label="หมายเหตุ">
              <input
                value={form.note}
                onChange={(e) => setForm({ ...form, note: e.target.value })}
              />
            </Field>
          </div>
          {page === 'รับเข้า' && receiveDrafts.length > 0 && <div className="receive-import-preview"><strong>รายการจากเอกสาร</strong><div>{receiveDrafts.map((item) => <p key={`${item.product}-${item.lot}`}><span>{item.product}</span><b>{item.qty} {item.unit}</b><small>Lot {item.lot || '-'} · ผลิต {displayIsoDate(item.mfg)} · หมดอายุ {displayIsoDate(item.exp)}</small></p>)}</div></div>}
          <div className="form-actions">
            <button onClick={() => setOpen(false)}>ยกเลิก</button>
            <button className="primary" onClick={save}>
              <Save size={15} />
              บันทึกรายการ
            </button>
          </div>
        </section>
      )}
      <RecordTable rows={rows} page={page} setRecords={setRecords} />
    </>
  );
}
function RecordTable({
  rows,
  page,
  setRecords,
}: {
  rows: RecordRow[];
  page: string;
  setRecords: (v: RecordRow[] | ((p: RecordRow[]) => RecordRow[])) => void;
}) {
  const [query, setQuery] = useState('');
  const filteredRows = rows.filter((row) => `${row.product} ${row.party} ${row.note} ${row.type}`.toLowerCase().includes(query.toLowerCase()));
  const printTitle = configs[page]?.title || (page === 'ประวัติ' ? 'ประวัติการเคลื่อนไหว Stock' : page);
  const rowDates = filteredRows.map((row) => row.date).filter(Boolean).sort();
  const printPeriod = rowDates.length
    ? rowDates[0] === rowDates[rowDates.length - 1]
      ? formatShortDate(rowDates[0])
      : `${formatShortDate(rowDates[0])} ถึง ${formatShortDate(rowDates[rowDates.length - 1])}`
    : 'ไม่พบรายการ';
  const exportRows = () => {
    const data = filteredRows.map((row, index) => ({ วันที่: formatShortDate(row.date), เลขที่เอกสาร: docNo(row.type, index), ประเภท: row.type, สินค้า: row.product, จำนวน: row.qty, หน่วย: row.unit, สาขา: row.branch, รายละเอียด: row.party, หมายเหตุ: row.note, สถานะ: row.status }));
    const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(data), 'รายการ'); XLSX.writeFile(workbook, `michiko_${page}_${today()}.xlsx`);
  };
  return (
    <section className="ops-table panel">
      <div className="report-print-head">
        <span className="report-logo-crop"><img src="/michiko-logo.png" alt="MICHIKO" /></span>
        <div><h2>{printTitle}</h2><p>{printPeriod} · MICHIKO สาขา Emsphere</p></div>
      </div>
      <div className="ops-toolbar">
        <label>
          <Search size={16} />
          <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="พิมพ์ค้นหาสินค้า ประเภท หรือสาขา..." />
        </label>
        <button onClick={printOperationalReport}>
          <Printer size={15} /> พิมพ์
        </button>
        <button onClick={exportRows}>
          <FileDown size={15} /> Export Excel
        </button>
      </div>
      <table>
        <thead>
          <tr>
            <th>วันที่</th>
            <th>เลขที่เอกสาร</th>
            <th>สินค้า</th>
            <th>จำนวน</th>
            <th>รายละเอียด</th>
            <th>สถานะ</th>
          </tr>
        </thead>
        <tbody>
          {filteredRows.length ? (
            filteredRows.map((r, i) => (
              <tr key={r.id}>
                <td>{formatShortDate(r.date)}</td>
                <td>{docNo(r.type, i)}</td>
                <td>
                  <strong>{r.product}</strong>
                  <small>{r.branch}</small>
                </td>
                <td>
                  {r.qty} {r.unit}
                </td>
                <td>
                  {r.party || '-'}
                  <small>{[r.lot && `Lot ${r.lot}`, r.mfg && `ผลิต ${formatShortDate(r.mfg)}`, r.exp && `หมดอายุ ${formatShortDate(r.exp)}`, r.note, r.attachment].filter(Boolean).join(' · ')}</small>
                </td>
                <td>
                  <span
                    className={`status ${r.status.includes('รอ') ? 'pending-status' : ''}`}
                  >
                    {r.status}
                  </span>
                  {page === 'โอนย้าย / ยืมคืน' &&
                    r.status.includes('รอ') && (
                      <button
                        className="confirm-mini"
                        onClick={() =>
                          setRecords((all) =>
                            all.map((x) =>
                              x.id === r.id
                                ? {
                                    ...x,
                                    status:
                                      r.type === 'โอนย้าย' ? 'รับแล้ว' : 'คืนครบแล้ว',
                                  }
                                : x,
                            ),
                          )
                        }
                      >
                        <CheckCircle2 size={13} /> ยืนยัน
                      </button>
                    )}
                </td>
              </tr>
            ))
          ) : (
            <tr>
              <td colSpan={6}>
                <Empty text="ยังไม่มีรายการในเมนูนี้" />
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </section>
  );
}
function Inventory({
  products,
  setProducts,
  notify,
}: {
  products: Product[];
  setProducts: (v: Product[] | ((p: Product[]) => Product[])) => void;
  notify: (s: string) => void;
}) {
  const [q, setQ] = useState('');
  const [open, setOpen] = useState(false);
  const inventoryFile = useRef<HTMLInputElement>(null);
  const [form, setForm] = useState({
    branch: branches[0],
    productId: products[0]?.id || '',
    stock: products[0]?.stock || 0,
    lot: '',
    expiry: '',
    note: '',
  });
  const [productQuery, setProductQuery] = useState(products[0]?.name || '');
  const list = products.filter((p) =>
    p.name.toLowerCase().includes(q.toLowerCase()),
  );
  const selected = products.find((p) => p.id === form.productId);
  const importInventory = async (file: File) => {
    try {
      let fileData: ArrayBuffer | Uint8Array = await file.arrayBuffer();
      if (file.name.toLowerCase().endsWith('.zip')) {
        const zip = await JSZip.loadAsync(fileData);
        const entry = Object.values(zip.files).find(
          (item) => !item.dir && /\.(csv|xlsx|xls)$/i.test(item.name),
        );
        if (!entry) throw new Error('no spreadsheet in zip');
        fileData = await entry.async('uint8array');
      }
      const workbook = XLSX.read(fileData, {
        type: 'array',
        cellDates: true,
      });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
        header: 1,
        defval: '',
        raw: true,
      });
      const locationHeader = rows.findIndex((row) =>
        row.some((cell) => String(cell).trim() === 'ชื่อยา'),
      );
      if (locationHeader >= 0) {
        const headers = rows[locationHeader].map((cell) => String(cell).replace(/^\uFEFF/, '').trim());
        const column = (name: string) => headers.findIndex((header) => header.includes(name));
        const nameCol = column('ชื่อยา');
        const codeCol = column('รหัสยา');
        const categoryCol = column('หมวดหมู่');
        const unitCol = column('หน่วย');
        const minimumCol = column('ยอดสต็อกขั้นต่ำ');
        const lotCol = column('เลขล็อตผู้ผลิต');
        const expiryCol = column('วันหมดอายุ');
        const stockCol = column('ยอดคงคลัง ณ วันที่เลือก');
        const grouped = new Map<string, Product>();
        for (const row of rows.slice(locationHeader + 1)) {
          const name = String(row[nameCol] || '').trim();
          if (!name) continue;
          const key = String(row[codeCol] || name).trim();
          const quantity = Number(row[stockCol]) || 0;
          const expiry = formatExcelValue(row[expiryCol]);
          const lot = String(row[lotCol] || '').trim();
          const current = grouped.get(key);
          if (current) {
            current.stock += quantity;
            if (expiry && !current.expiry?.includes(expiry))
              current.expiry = [current.expiry, expiry].filter(Boolean).join(', ');
            if (lot && !current.note?.includes(lot))
              current.note = [current.note, `Lot ${lot}`].filter(Boolean).join(', ');
          } else {
            const existing = products.find((p) => p.code === key || normalizeName(p.name) === normalizeName(name));
            grouped.set(key, {
              id: existing?.id || id(),
              code: key,
              name,
              category: String(row[categoryCol] || existing?.category || 'อื่น ๆ'),
              unit: String(row[unitCol] || existing?.unit || 'อัน'),
              minimum: Number(row[minimumCol]) || existing?.minimum || 0,
              active: true,
              stock: quantity,
              expiry: expiry || existing?.expiry,
              note: lot ? `Lot ${lot}` : existing?.note,
            });
          }
        }
        const imported = [...grouped.values()];
        if (!imported.length) throw new Error('no inventory rows');
        if (!window.confirm(`พบสินค้า ${imported.length} รายการจากไฟล์ระบบ\nต้องการใช้เป็นยอดคงคลังปัจจุบันหรือไม่?`)) return;
        setProducts((current) => {
          const ids = new Set(imported.map((p) => p.id));
          return [...imported, ...current.filter((p) => !ids.has(p.id))];
        });
        notify(`นำเข้ายอดคงคลังจากไฟล์ระบบจำนวน ${imported.length} รายการแล้ว`);
        return;
      }
      let category = 'อื่น ๆ';
      const imported: Product[] = [];
      for (const row of rows) {
        const order = row[0];
        const name = String(row[1] || '').trim();
        if (!name) continue;
        const opening = typeof row[2] === 'number' ? row[2] : undefined;
        const closing = typeof row[7] === 'number' ? row[7] : undefined;
        if (typeof order !== 'number' && opening === undefined && closing === undefined) {
          if (!/ยอด|clinic/i.test(name)) category = name.replace(/ยอด|คงคลัง|สิ้นเดือน/g, '').trim() || category;
          continue;
        }
        if (opening === undefined && closing === undefined) continue;
        const existing = products.find((p) => normalizeName(p.name) === normalizeName(name));
        const expiry = formatExcelValue(row[8]);
        imported.push({
          id: existing?.id || id(),
          code: existing?.code || `IMP-${String(imported.length + 1).padStart(3, '0')}`,
          name,
          category: existing?.category || category,
          unit: existing?.unit || guessUnit(name),
          minimum: existing?.minimum || 0,
          active: true,
          stock: closing ?? opening ?? 0,
          expiry: expiry || existing?.expiry,
          note: String(row[9] || existing?.note || ''),
        });
      }
      if (!imported.length) throw new Error('no rows');
      if (!window.confirm(`พบสินค้า ${imported.length} รายการ\nต้องการใช้ยอดจากไฟล์นี้เป็นยอดคงคลังปัจจุบันหรือไม่?`)) return;
      setProducts((current) => {
        const ids = new Set(imported.map((p) => p.id));
        return [...imported, ...current.filter((p) => !ids.has(p.id))];
      });
      notify(`นำเข้ายอดคงคลังจาก ${file.name} จำนวน ${imported.length} รายการแล้ว`);
    } catch {
      notify('อ่านไฟล์ไม่สำเร็จ กรุณาตรวจสอบรูปแบบ Excel');
    }
  };
  const saveOpeningStock = () => {
    if (!form.productId || form.stock < 0) return;
    setProducts((all) =>
      all.map((p) =>
        p.id === form.productId
          ? {
              ...p,
              stock: form.stock,
              expiry: form.expiry || p.expiry,
              note: [form.lot ? `Lot ${form.lot}` : '', form.note]
                .filter(Boolean)
                .join(' · ') || p.note,
            }
          : p,
      ),
    );
    setOpen(false);
    notify(`บันทึกยอดคงคลังปัจจุบันของ ${selected?.name || 'สินค้า'} แล้ว`);
  };
  return (
    <>
      <PageHead
        icon={<Boxes />}
        title="Stock คงคลัง"
        sub="ยอดปัจจุบันแยกตามสินค้า หมวด และสาขา"
        action={() => setOpen(!open)}
        actionText={open ? 'ปิดแบบฟอร์ม' : '+ เพิ่มยอดคงคลังปัจจุบัน'}
      />
      <input
        ref={inventoryFile}
        hidden
        type="file"
        accept=".zip,.xlsx,.xls,.csv"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) void importInventory(file);
          e.target.value = '';
        }}
      />
      <div className="inventory-import-row">
        <button onClick={() => inventoryFile.current?.click()}>
          <FileUp size={16} /> นำเข้า Excel ยอดคงคลัง
        </button>
        <span>รองรับไฟล์ ZIP จากระบบโดยตรง รวมถึง .xlsx, .xls และ .csv</span>
      </div>
      {open && (
        <section className="ops-form panel opening-stock-form">
          <div className="opening-note">
            ใช้สำหรับบันทึกยอดที่นับได้จริงเป็นฐานเริ่มต้นของระบบ
          </div>
          <div className="form-grid">
            <Field label="สาขา">
              <select
                value={form.branch}
                onChange={(e) => setForm({ ...form, branch: e.target.value })}
              >
                {branches.map((branch) => <option key={branch}>{branch}</option>)}
              </select>
            </Field>
            <Field label="สินค้า">
              <input list="inventory-products" value={productQuery}
                onChange={(e) => {
                  const query = e.target.value; setProductQuery(query);
                  const product = products.find((p) => p.name === query || `${p.code} — ${p.name}` === query);
                  if (product) { setProductQuery(product.name); setForm({ ...form, productId: product.id, stock: product.stock }); }
                }}
                placeholder="พิมพ์รหัสหรือชื่อสินค้า" />
              <datalist id="inventory-products">{products.map((p) => <option key={p.id} value={`${p.code} — ${p.name}`} />)}</datalist>
            </Field>
            <Field label="ยอดคงเหลือจริง">
              <input type="number" min="0" step="0.01" value={form.stock} onChange={(e) => setForm({ ...form, stock: Number(e.target.value) })} />
            </Field>
            <Field label="หน่วย">
              <input value={selected?.unit || ''} readOnly />
            </Field>
            <Field label="Lot No.">
              <input value={form.lot} onChange={(e) => setForm({ ...form, lot: e.target.value })} />
            </Field>
            <Field label="Exp.">
              <ThaiDateInput value={form.expiry} onChange={(expiry) => setForm({ ...form, expiry })} />
            </Field>
            <Field label="หมายเหตุ">
              <input value={form.note} onChange={(e) => setForm({ ...form, note: e.target.value })} placeholder="เช่น ของการตลาด" />
            </Field>
          </div>
          <div className="form-actions">
            <button onClick={() => setOpen(false)}>ยกเลิก</button>
            <button className="primary" onClick={saveOpeningStock}><Save size={15} /> บันทึกยอดตั้งต้น</button>
          </div>
        </section>
      )}
      <section className="ops-table panel">
        <div className="report-print-head">
          <span className="report-logo-crop"><img src="/michiko-logo.png" alt="MICHIKO" /></span>
          <div><h2>รายงาน Stock คงคลัง</h2><p>{formatShortDate(today())} · MICHIKO สาขา Emsphere</p></div>
        </div>
        <div className="ops-toolbar">
          <label>
            <Search size={16} />
            <input
              value={q}
              onChange={(e) => setQ(e.target.value)}
              placeholder="ค้นหาสินค้า..."
            />
          </label>
          <select>
            <option>ทุกหมวดสินค้า</option>
            <option>Filler</option>
            <option>Botox</option>
            <option>อุปกรณ์</option>
          </select>
          <button onClick={printOperationalReport}>
            <Printer size={15} /> พิมพ์
          </button>
        </div>
        <table>
          <thead>
            <tr>
              <th>รหัส</th>
              <th>สินค้า</th>
              <th>หมวด</th>
              <th>คงเหลือ</th>
              <th>Minimum</th>
              <th>สถานะ</th>
            </tr>
          </thead>
          <tbody>
            {list.map((p) => (
              <tr key={p.id}>
                <td>{p.code}</td>
                <td>
                  <strong>{p.name}</strong>
                </td>
                <td>{p.category}</td>
                <td>
                  <b>{p.stock}</b> {p.unit}
                </td>
                <td>
                  {p.minimum} {p.unit}
                </td>
                <td>
                  <span
                    className={`status ${p.stock <= p.minimum ? 'low-status' : ''}`}
                  >
                    {p.stock <= p.minimum ? 'Stock ต่ำ' : 'ปกติ'}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </section>
    </>
  );
}
function Products({
  products,
  setProducts,
  notify,
}: {
  products: Product[];
  setProducts: (v: Product[] | ((p: Product[]) => Product[])) => void;
  notify: (s: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [f, setF] = useState({
    code: '',
    name: '',
    category: 'อุปกรณ์',
    unit: 'อัน',
    minimum: 0,
  });
  const save = () => {
    if (!f.name) return;
    setProducts((p) => [...p, { ...f, id: id(), active: true, stock: 0 }]);
    setOpen(false);
    notify('เพิ่มสินค้าในทะเบียนแล้ว');
  };
  return (
    <>
      <PageHead
        icon={<Tags />}
        title="ทะเบียนสินค้า"
        sub="ชื่อมาตรฐาน หน่วยหลัก Minimum Stock และสถานะสินค้า"
        action={() => setOpen(!open)}
        actionText="+ เพิ่มสินค้า"
      />
      {open && (
        <section className="ops-form panel">
          <div className="form-grid">
            <Field label="รหัสสินค้า">
              <input
                value={f.code}
                onChange={(e) => setF({ ...f, code: e.target.value })}
              />
            </Field>
            <Field label="ชื่อมาตรฐาน">
              <input
                value={f.name}
                onChange={(e) => setF({ ...f, name: e.target.value })}
              />
            </Field>
            <Field label="หมวด">
              <input
                value={f.category}
                onChange={(e) => setF({ ...f, category: e.target.value })}
              />
            </Field>
            <Field label="หน่วยหลัก">
              <input
                value={f.unit}
                onChange={(e) => setF({ ...f, unit: e.target.value })}
              />
            </Field>
            <Field label="Minimum Stock">
              <input
                type="number"
                value={f.minimum}
                onChange={(e) =>
                  setF({ ...f, minimum: Number(e.target.value) })
                }
              />
            </Field>
          </div>
          <div className="form-actions">
            <button className="primary" onClick={save}>
              <Save size={15} />
              บันทึก
            </button>
          </div>
        </section>
      )}
      <section className="ops-table panel">
        <div className="report-print-head">
          <span className="report-logo-crop"><img src="/michiko-logo.png" alt="MICHIKO" /></span>
          <div><h2>ทะเบียนสินค้า</h2><p>พิมพ์เมื่อ {formatShortDate(today())} · MICHIKO</p></div>
        </div>
        <table>
          <thead>
            <tr>
              <th>รหัส</th>
              <th>ชื่อมาตรฐาน</th>
              <th>หมวด</th>
              <th>หน่วย</th>
              <th>Minimum</th>
              <th>สถานะ</th>
            </tr>
          </thead>
          <tbody>
            {products.map((p) => (
              <tr key={p.id}>
                <td>{p.code}</td>
                <td>
                  <strong>{p.name}</strong>
                </td>
                <td>{p.category}</td>
                <td>{p.unit}</td>
                <td>{p.minimum}</td>
                <td>
                  <button
                    className={`toggle ${p.active ? 'on' : ''}`}
                    onClick={() =>
                      setProducts((all) =>
                        all.map((x) =>
                          x.id === p.id ? { ...x, active: !x.active } : x,
                        ),
                      )
                    }
                  >
                    {p.active ? 'Active' : 'Inactive'}
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </section>
    </>
  );
}
function Reports({
  records,
  products,
  notify,
  dailyCases,
}: {
  records: RecordRow[];
  products: Product[];
  notify: (s: string) => void;
  dailyCases: Array<{ date: string; items: Array<{ name: string; qty: number }> }>;
}) {
  const [view, setView] = useState<'compare' | 'monthly'>('compare');
  const [snapshots, setSnapshots] = useSharedStored<StockSnapshot[]>('michiko-stock-snapshots', []);
  const openingInput = useRef<HTMLInputElement>(null);
  const latestInput = useRef<HTMLInputElement>(null);
  const sortedSnapshots = [...snapshots].sort((a, b) => a.date.localeCompare(b.date));
  const [openingId, setOpeningId] = useState('');
  const [closingId, setClosingId] = useState('');
  useEffect(() => {
    if (!sortedSnapshots.length) return;
    const closing = sortedSnapshots.at(-1)!;
    const opening = sortedSnapshots.at(-2) || closing;
    if (!closingId || !snapshots.some((s) => s.id === closingId)) setClosingId(closing.id);
    if (!openingId || !snapshots.some((s) => s.id === openingId)) setOpeningId(opening.id);
  }, [snapshots]);
  const openingSnapshot = snapshots.find((s) => s.id === openingId);
  const closingSnapshot = snapshots.find((s) => s.id === closingId);
  const comparison = useMemo(() => compareSnapshots(openingSnapshot, closingSnapshot), [openingSnapshot, closingSnapshot]);
  const inPeriod = (date: string) => (!openingSnapshot || date >= openingSnapshot.date) && (!closingSnapshot || date <= closingSnapshot.date);
  const movement = (product: string, types: string[]) => records.filter((r) => inPeriod(r.date) && normalizeName(r.product) === normalizeName(product) && types.includes(r.type)).reduce((sum, r) => sum + r.qty, 0);
  const decreases = comparison.filter((row) => row.change < 0);
  const increases = comparison.filter((row) => row.change > 0);
  const importantRows = comparison.filter((row) => row.change !== 0).slice(0, 15).map((row) => {
    const received = movement(row.name, ['รับเข้า']);
    const issued = movement(row.name, ['เบิกออก']);
    const borrowed = movement(row.name, ['ยืม']);
    const transferred = movement(row.name, ['โอนย้าย']);
    const used = dailyCases.filter((daily) => inPeriod(daily.date)).reduce((sum, daily) => sum + daily.items.filter((item) => normalizeName(item.name) === normalizeName(row.name)).reduce((n, item) => n + Number(item.qty || 0), 0), 0) + issued;
    const documentedClosing = row.opening + received - used - borrowed - transferred;
    return { ...row, received, used, borrowed, transferred, adjustment: row.closing - documentedClosing };
  });
  const importSnapshot = async (file: File, kind: 'opening' | 'latest') => {
    try {
      const snapshot = await readStockSnapshot(file);
      setSnapshots((current) => [...current.filter((s) => s.date !== snapshot.date), snapshot].sort((a, b) => a.date.localeCompare(b.date)));
      if (kind === 'opening') setOpeningId(snapshot.id); else setClosingId(snapshot.id);
      notify(`บันทึกยอดคงคลังวันที่ ${formatSnapshotDate(snapshot.date)} แล้ว`);
    } catch { notify('อ่านไฟล์ไม่ได้ กรุณาใช้ไฟล์ Stock Location Report จากระบบ'); }
  };
  const exportComparison = () => {
    if (!openingSnapshot || !closingSnapshot) return;
    const rows = comparison.map((row, index) => ({ ลำดับ: index + 1, รหัสสินค้า: row.code, สินค้า: row.name, หมวดหมู่: row.category, ยอดต้นงวด: row.opening, ยอดล่าสุด: row.closing, เปลี่ยนแปลง: row.change, หน่วย: row.unit, สถานะ: row.change < 0 ? 'ยอดลดลง' : row.change > 0 ? 'ยอดเพิ่มขึ้น' : 'ไม่เปลี่ยนแปลง' }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), 'เปรียบเทียบยอดคงคลัง');
    XLSX.writeFile(workbook, `stock_comparison_${openingSnapshot.date}_${closingSnapshot.date}.xlsx`);
  };
  const incoming = records
    .filter((r) => r.type === 'รับเข้า')
    .reduce((n, r) => n + r.qty, 0);
  const outgoing = records
    .filter((r) => ['เบิกออก', 'โอนย้าย', 'ยืม'].includes(r.type))
    .reduce((n, r) => n + r.qty, 0);
  const stockAt = (snapshot: StockSnapshot | undefined, product: Product) => snapshot?.items.find((item) => item.code === product.code || normalizeName(item.name) === normalizeName(product.name))?.stock;
  const accountingRow = (product: Product) => {
    const opening = stockAt(openingSnapshot, product) ?? 0, closing = stockAt(closingSnapshot, product) ?? product.stock;
    const received = movement(product.name, ['รับเข้า', 'ซื้อจากสาขา']);
    const transferred = movement(product.name, ['โอนย้าย', 'ขายให้สาขา']);
    const borrowed = movement(product.name, ['ยืม']);
    const returned = movement(product.name, ['คืน']);
    const issued = movement(product.name, ['เบิกออก']);
    const used = dailyCases.filter((daily) => inPeriod(daily.date)).reduce((sum, daily) => sum + daily.items.filter((item) => normalizeName(item.name) === normalizeName(product.name)).reduce((n, item) => n + Number(item.qty || 0), 0), 0) + issued;
    const adjustment = closing - (opening - used - transferred - borrowed + received + returned);
    return { product, opening, used, transferred: transferred + borrowed, received, adjustment: adjustment + returned, closing };
  };
  const findTemplateProduct = (category: string, label: string) => {
    const aliases: Record<string, string> = { Kysses: 'Kysse' };
    const searchLabel = aliases[label] || label;
    const exact = products.find((product) => normalizeName(product.name) === normalizeName(searchLabel));
    if (exact) return exact;
    const branded = normalizeName(`${category} ${searchLabel}`);
    const brandMatch = products.find((product) => {
      const candidate = normalizeName(product.name);
      return candidate.includes(branded) || branded.includes(candidate);
    });
    if (brandMatch) return brandMatch;
    const withoutUnit = normalizeName(searchLabel.replace(/\((?:ขวด|กล่อง|amp\.?|vial|unit|s\.u\.|ml\.?)\)/gi, ''));
    return products.find((product) => {
      const candidate = normalizeName(product.name);
      return withoutUnit.length > 2 && (candidate.includes(withoutUnit) || withoutUnit.includes(candidate));
    });
  };
  const accountingRows = accountingTemplate.flatMap(([category, names], sectionIndex) => names.split('|').map((displayName, itemIndex) => {
    const matched = findTemplateProduct(category, displayName);
    const product: Product = matched || { id: `template-${sectionIndex}-${itemIndex}`, code: '', name: displayName, category, unit: guessUnit(displayName), minimum: 0, active: true, stock: 0 };
    return { ...accountingRow(product), category, displayName, sectionIndex, itemIndex };
  }));
  const exportMonthly = () => {
    const sheetRows: (string | number)[][] = [['ลำดับ','สินค้า','ยอดคงคลังเดิม','จำนวนที่ใช้/ขายไป','จำนวนที่โอนย้าย','ยอดรับเข้า','ปรับยอด/คืน','ยอดคงคลังสิ้นเดือน','Exp.','หมายเหตุ']];
    accountingTemplate.forEach(([category], sectionIndex) => {
      sheetRows.push([category,'','','','','','','','','']);
      accountingRows.filter((row) => row.sectionIndex === sectionIndex).forEach((row, index) => sheetRows.push([index + 1,row.displayName,row.opening,row.used || 0,row.transferred || 0,row.received || 0,row.adjustment || 0,row.closing,row.product.expiry || '',row.product.note || '']));
    });
    const sheet = XLSX.utils.aoa_to_sheet(sheetRows); sheet['!cols'] = [{wch:8},{wch:42},{wch:15},{wch:18},{wch:16},{wch:13},{wch:15},{wch:19},{wch:24},{wch:28}];
    const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, sheet, 'รายงานสต็อคคงคลัง'); XLSX.writeFile(workbook, `รายงานสต็อคคงคลัง_${closingSnapshot?.date || today()}.xlsx`);
  };
  return (
    <>
      <PageHead
        icon={<ClipboardCheck />}
        title="รายงานประจำเดือน"
        sub="สรุปยอดรับเข้า Stock รายวัน เบิก โอน และยอดปลายเดือน"
      />
      <div className="report-tabs"><button className={view === 'compare' ? 'active' : ''} onClick={() => setView('compare')}>เปรียบเทียบยอดคงคลัง</button><button className={view === 'monthly' ? 'active' : ''} onClick={() => setView('monthly')}>รายงานส่งบัญชีสิ้นเดือน</button></div>
      {view === 'compare' && <>
        <input ref={openingInput} hidden type="file" accept=".zip,.xlsx,.xls,.csv" onChange={(e) => { const file = e.target.files?.[0]; if (file) void importSnapshot(file, 'opening'); e.currentTarget.value = ''; }} />
        <input ref={latestInput} hidden type="file" accept=".zip,.xlsx,.xls,.csv" onChange={(e) => { const file = e.target.files?.[0]; if (file) void importSnapshot(file, 'latest'); e.currentTarget.value = ''; }} />
        <section className="snapshot-controls panel"><div className="snapshot-select"><span>ยอดต้นงวด</span><select value={openingId} onChange={(e) => setOpeningId(e.target.value)}><option value="">ยังไม่มีข้อมูล</option>{sortedSnapshots.map((s) => <option key={s.id} value={s.id}>{formatSnapshotDate(s.date)} · {s.items.length} รายการ</option>)}</select><button onClick={() => openingInput.current?.click()}><FileUp size={15}/> เปลี่ยนไฟล์ต้นงวด</button></div><div className="compare-arrow">→</div><div className="snapshot-select"><span>ยอดล่าสุด</span><select value={closingId} onChange={(e) => setClosingId(e.target.value)}><option value="">ยังไม่มีข้อมูล</option>{sortedSnapshots.map((s) => <option key={s.id} value={s.id}>{formatSnapshotDate(s.date)} · {s.items.length} รายการ</option>)}</select><button className="primary" onClick={() => latestInput.current?.click()}><FileUp size={15}/> นำเข้าไฟล์ล่าสุด</button></div></section>
        {openingSnapshot && closingSnapshot ? <><div className="report-metrics comparison-metrics"><Mini title="สินค้ายอดลดลง" value={`${decreases.length} รายการ`} /><Mini title="สินค้ายอดเพิ่มขึ้น" value={`${increases.length} รายการ`} /><Mini title="ช่วงที่เปรียบเทียบ" value={`${formatShortDate(openingSnapshot.date)}–${formatShortDate(closingSnapshot.date)}`} /><Mini title="รายการสำคัญในรายงาน" value={`${importantRows.length} รายการ`} /></div><section className="ops-table panel comparison-report"><div className="report-print-head"><span className="report-logo-crop"><img src="/michiko-logo.png" alt="MICHIKO"/></span><div><h2>รายงานสรุปการเคลื่อนไหว Stock</h2><p>{formatShortDate(openingSnapshot.date)} ถึง {formatShortDate(closingSnapshot.date)} · MICHIKO สาขา Emsphere</p></div></div><div className="ops-toolbar"><strong>รายการเคลื่อนไหวสำคัญ 15 อันดับ</strong><button onClick={printOperationalReport}><Printer size={15}/> พิมพ์</button><button onClick={exportComparison}><FileDown size={15}/> Export Excel</button></div><table className="movement-summary-table"><thead><tr><th>รหัส / สินค้า</th><th>ต้นงวด</th><th>รับเข้า</th><th>ใช้/ขาย</th><th>ยืม</th><th>โอน</th><th>ปรับยอด*</th><th>คงเหลือ</th><th>หน่วย</th></tr></thead><tbody>{importantRows.map((row) => <tr key={row.code}><td><strong>{row.code} · {row.name}</strong><small>{row.category}</small></td><td>{formatQty(row.opening)}</td><td>{row.received ? formatQty(row.received) : '-'}</td><td>{row.used ? formatQty(row.used) : '-'}</td><td>{row.borrowed ? formatQty(row.borrowed) : '-'}</td><td>{row.transferred ? formatQty(row.transferred) : '-'}</td><td className={row.adjustment < 0 ? 'change-down' : row.adjustment > 0 ? 'change-up' : ''}>{row.adjustment ? `${row.adjustment > 0 ? '+' : ''}${formatQty(row.adjustment)}` : '-'}</td><td><b>{formatQty(row.closing)}</b></td><td>{row.unit}</td></tr>)}</tbody></table></section><p className="report-note">* ปรับยอด/รอตรวจสอบ คือผลต่างที่ยังไม่มีเอกสารรับเข้า Stock รายวัน เบิก ยืม หรือโอนในระบบ เมื่อนำเข้าเอกสาร JERA ตัวเลขจะถูกแยกเข้าช่องที่ถูกต้อง</p></> : <section className="panel snapshot-empty"><Boxes/><h3>นำเข้าไฟล์ต้นงวดและไฟล์ล่าสุด</h3><p>ครั้งต่อไประบบจะจำยอดล่าสุดไว้เป็นต้นงวดให้โดยอัตโนมัติ</p></section>}
      </>}
      {view === 'monthly' && <>
      <div className="report-metrics">
        <Mini title="ยอดรับเข้า" value={`+ ${incoming}`} />
        <Mini title="Stock รายวัน" value={`- ${accountingRows.reduce((sum, row) => sum + row.used, 0)}`} />
        <Mini title="เบิก / โอนออก" value={`- ${outgoing}`} />
        <Mini
          title="Stock คงเหลือ"
          value={String(products.reduce((n, p) => n + p.stock, 0))}
        />
      </div>
      <section className="ops-table panel">
        <div className="report-print-head">
          <span className="report-logo-crop"><img src="/michiko-logo.png" alt="MICHIKO" /></span>
          <div><h2>รายงานสต็อกคงคลังประจำเดือน</h2><p>{closingSnapshot ? formatSnapshotDate(closingSnapshot.date) : formatShortDate(today())} · MICHIKO สาขา Emsphere</p></div>
        </div>
        <div className="ops-toolbar">
          <select>
            <option>กันยายน 2026</option>
          </select>
          <button onClick={printOperationalReport}>
            <Printer size={15} /> พิมพ์รายงาน
          </button>
          <button onClick={exportMonthly}>
            <FileDown size={15} /> Export Excel
          </button>
        </div>
        <table className="monthly-accounting-table">
          <thead>
            <tr>
              <th>ลำดับ</th>
              <th>สินค้า</th>
              <th>ยอดคงคลังเดิม</th>
              <th>จำนวนที่ใช้/ขายไป</th>
              <th>จำนวนที่โอนย้าย</th>
              <th>ยอดรับเข้า</th>
              <th>ปรับยอด/คืน</th>
              <th>ยอดคงคลังสิ้นเดือน</th>
              <th>Exp.</th>
              <th>หมายเหตุ</th>
            </tr>
          </thead>
          <tbody>
            {accountingTemplate.flatMap(([category], sectionIndex) => {
              const list = accountingRows.filter((row) => row.sectionIndex === sectionIndex);
              return [
                <tr className="category-row" key={`cat-${sectionIndex}-${category}`}><td colSpan={10}>{category}</td></tr>,
                ...list.map((row, index) => <tr key={`${row.sectionIndex}-${row.itemIndex}`}><td>{index + 1}</td><td><strong>{row.displayName}</strong><small>{row.product.unit}</small></td><td>{formatQty(row.opening)}</td><td>{row.used ? formatQty(row.used) : '-'}</td><td>{row.transferred ? formatQty(row.transferred) : '-'}</td><td>{row.received ? formatQty(row.received) : '-'}</td><td>{row.adjustment ? formatQty(row.adjustment) : '-'}</td><td><b>{formatQty(row.closing)}</b> {row.product.unit}</td><td>{row.product.expiry || '-'}</td><td>{row.product.note || '-'}</td></tr>),
              ];
            })}
          </tbody>
        </table>
      </section>
      </>}
    </>
  );
}
function Movements({ records }: { records: RecordRow[] }) {
  return (
    <>
      <PageHead
        icon={<History />}
        title="ประวัติการเคลื่อนไหว"
        sub="ตรวจสอบทุกการรับเข้า ใช้ โอน ยืม คืน เบิก และปรับยอด"
      />
      <RecordTable rows={records} page="ประวัติ" setRecords={() => {}} />
    </>
  );
}
function Preferences({ notify }: { notify: (s: string) => void }) {
  const [assistants, setAssistants] = useSharedStored<string[]>('michiko-assistants', []);
  const [doctors, setDoctors] = useSharedStored<string[]>('michiko-doctors', []);
  const [assistantName, setAssistantName] = useState('');
  const [doctorName, setDoctorName] = useState('');
  const addName = (value: string, list: string[], save: (value: string[]) => void, clear: (value: string) => void) => {
    const name = value.trim(); if (!name || list.includes(name)) return;
    save([...list, name]); clear(''); notify(`เพิ่ม ${name} แล้ว`);
  };
  return (
    <>
      <PageHead
        icon={<Settings />}
        title="ตั้งค่า"
        sub="ข้อมูลสาขา ผู้ใช้งาน รายชื่อพนักงาน และรูปแบบเลขเอกสาร"
      />
      <section className="settings-grid">
        <div className="panel setting-card">
          <h3>สาขาในระบบ</h3>
          <p>MICHIKO สาขาพหลโยธิน</p>
          <p>MICHIKO สาขา Emsphere</p>
          <button>
            <Plus size={14} /> เพิ่มสาขา
          </button>
        </div>
        <div className="panel setting-card">
          <h3>รายชื่อผู้ช่วย</h3>
          {assistants.length ? assistants.map((name) => <p className="managed-name" key={name}><span>{name}</span><button onClick={() => setAssistants(assistants.filter((x) => x !== name))}>ลบ</button></p>) : <p>ยังไม่มีรายชื่อผู้ช่วย</p>}
          <form className="name-entry" onSubmit={(e) => { e.preventDefault(); addName(assistantName, assistants, setAssistants, setAssistantName); }}><input value={assistantName} onChange={(e) => setAssistantName(e.target.value)} placeholder="พิมพ์ชื่อผู้ช่วย"/><button type="submit"><Plus size={14}/> เพิ่มและบันทึก</button></form>
        </div>
        <div className="panel setting-card">
          <h3>รายชื่อแพทย์</h3>
          {doctors.length ? doctors.map((name) => <p className="managed-name" key={name}><span>{name}</span><button onClick={() => setDoctors(doctors.filter((x) => x !== name))}>ลบ</button></p>) : <p>ยังไม่มีรายชื่อแพทย์</p>}
          <form className="name-entry" onSubmit={(e) => { e.preventDefault(); addName(doctorName, doctors, setDoctors, setDoctorName); }}><input value={doctorName} onChange={(e) => setDoctorName(e.target.value)} placeholder="พิมพ์ชื่อแพทย์"/><button type="submit"><Plus size={14}/> เพิ่มและบันทึก</button></form>
        </div>
        <div className="panel setting-card">
          <h3>ข้อมูลผู้ใช้งาน</h3>
          <p>ครีม · พิชญาพร · ผู้ดูแลระบบ</p>
          <button onClick={() => notify('บันทึกการตั้งค่าแล้ว')}>
            <Save size={14} /> บันทึกการตั้งค่า
          </button>
        </div>
      </section>
    </>
  );
}
function PageHead({
  icon,
  title,
  sub,
  action,
  actionText,
}: {
  icon: React.ReactNode;
  title: string;
  sub: string;
  action?: () => void;
  actionText?: string;
}) {
  return (
    <div className="ops-head">
      <div className="ops-title-icon">{icon}</div>
      <div>
        <h1>{title}</h1>
        <p>{sub}</p>
      </div>
      {action && <button onClick={action}>{actionText}</button>}
    </div>
  );
}
function Field({
  label,
  children,
}: {
  label: string;
  children: React.ReactNode;
}) {
  return (
    <label className="field">
      <span>{label}</span>
      {children}
    </label>
  );
}
function Empty({ text }: { text: string }) {
  return (
    <div className="ops-empty">
      <Boxes />
      <span>{text}</span>
    </div>
  );
}
function Mini({ title, value }: { title: string; value: string }) {
  return (
    <div className="panel report-mini">
      <span>{title}</span>
      <strong>{value}</strong>
      <small>รายการ</small>
    </div>
  );
}
function docNo(type: string, index: number) {
  const pre =
    type === 'รับเข้า'
      ? 'RC'
      : type === 'โอนย้าย'
        ? 'TF'
        : type === 'ยืม'
          ? 'LN'
          : type === 'เบิกออก'
            ? 'IS'
            : 'CT';
  return `${pre}-2609-${String(index + 1).padStart(3, '0')}`;
}

function normalizeName(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9ก-๙]/g, '');
}

async function readStockSnapshot(file: File): Promise<StockSnapshot> {
  let fileData: ArrayBuffer | Uint8Array = await file.arrayBuffer();
  if (file.name.toLowerCase().endsWith('.zip')) {
    const zip = await JSZip.loadAsync(fileData);
    const entry = Object.values(zip.files).find((item) => !item.dir && /\.(csv|xlsx|xls)$/i.test(item.name));
    if (!entry) throw new Error('missing sheet');
    fileData = await entry.async('uint8array');
  }
  const workbook = XLSX.read(fileData, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '', raw: true });
  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell).replace(/^\uFEFF/, '').trim() === 'ชื่อยา'));
  if (headerIndex < 0) throw new Error('invalid report');
  const headers = rows[headerIndex].map((cell) => String(cell).replace(/^\uFEFF/, '').trim());
  const column = (name: string) => headers.findIndex((header) => header.includes(name));
  const codeCol = column('รหัสยา'), nameCol = column('ชื่อยา'), categoryCol = column('หมวดหมู่'), unitCol = column('หน่วย'), stockCol = column('ยอดคงคลัง ณ วันที่เลือก');
  const grouped = new Map<string, StockSnapshotItem>();
  for (const row of rows.slice(headerIndex + 1)) {
    const name = String(row[nameCol] || '').trim(); if (!name) continue;
    const code = String(row[codeCol] || name).trim(), current = grouped.get(code), quantity = Number(row[stockCol]) || 0;
    if (current) current.stock += quantity; else grouped.set(code, { code, name, category: String(row[categoryCol] || 'อื่น ๆ'), unit: String(row[unitCol] || 'อัน'), stock: quantity });
  }
  const match = file.name.match(/(20\d{2})[-_](\d{2})[-_](\d{2})/), date = match ? `${match[1]}-${match[2]}-${match[3]}` : today();
  const items = [...grouped.values()]; if (!items.length) throw new Error('empty report');
  return { id: `${date}-${Date.now()}`, date, fileName: file.name, items };
}
function compareSnapshots(opening?: StockSnapshot, closing?: StockSnapshot) {
  if (!opening || !closing) return [];
  const all = new Map<string, { code: string; name: string; category: string; unit: string; opening: number; closing: number; change: number }>();
  opening.items.forEach((item) => all.set(item.code, { ...item, opening: item.stock, closing: 0, change: -item.stock }));
  closing.items.forEach((item) => { const previous = all.get(item.code); if (previous) Object.assign(previous, { closing: item.stock, change: item.stock - previous.opening }); else all.set(item.code, { ...item, opening: 0, closing: item.stock, change: item.stock }); });
  return [...all.values()].sort((a, b) => Math.abs(b.change) - Math.abs(a.change) || a.name.localeCompare(b.name, 'th'));
}
function formatSnapshotDate(date: string) { return new Intl.DateTimeFormat('th-TH', { day: 'numeric', month: 'long', year: 'numeric' }).format(new Date(`${date}T00:00:00`)); }
function formatShortDate(date: string) { const [year, month, day] = date.split('-'); return `${day}/${month}/${Number(year) + 543}`; }
function displayIsoDate(date: string) { return date ? formatShortDate(date) : '-'; }
function formatQty(value: number) { return new Intl.NumberFormat('th-TH', { maximumFractionDigits: 2 }).format(value); }

async function parseJeraReceivePdf(file: File, products: Product[]) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs');
  pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
  const document = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
  const tokens: string[] = [];
  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
    const content = await (await document.getPage(pageNumber)).getTextContent();
    for (const item of content.items) {
      if ('str' in item && item.str.trim()) tokens.push(item.str.trim());
    }
  }
  const allText = tokens.join('');
  const po = allText.match(/26854-4-\d+/)?.[0] || '';
  const supplier = tokens.find((token) => token.startsWith('บริษัท') && token.length > 12) || '';
  const codeIndexes = tokens.map((token, index) => ({ token, index })).filter(({ token }) => /^[A-Z]{1,3}\d{3}$/.test(token));
  const toIso = (date: string) => { const [day, month, year] = date.split('-'); return `${year}-${month}-${day}`; };
  const items: ReceiveDraft[] = codeIndexes.map(({ token: code, index }, position) => {
    const end = codeIndexes[position + 1]?.index ?? tokens.length;
    const segment = tokens.slice(index + 1, end);
    const segmentText = segment.join('');
    const product = products.find((candidate) => candidate.code === code);
    const fallbackName = segment.find((value) => /[A-Za-zก-๙]/.test(value) && !/^(Stock|Room|supp lot)/i.test(value))?.replace(/\s*\(1\)\s*$/, '') || code;
    const quantity = Number(segment.find((value) => /^\d+(?:\.\d+)?$/.test(value.replace(/,/g, '')))?.replace(/,/g, '')) || 1;
    const dates = [...segmentText.matchAll(/(\d{2}-\d{2}-\d{4})/g)].map((match) => match[1]);
    return {
      code,
      product: product?.name || fallbackName,
      qty: quantity,
      unit: product?.unit || segment.find((value) => ['ขวด', 'กล่อง', 'อัน', 'ชิ้น', 'cc', 'ml', 'ชุด'].includes(value)) || 'อัน',
      lot: segmentText.match(/supp lot no\.:\s*([^;]+)/i)?.[1]?.trim() || '',
      mfg: dates[0] ? toIso(dates[0]) : '',
      exp: dates[1] ? toIso(dates[1]) : '',
    };
  });
  return { po, supplier, items };
}

function inferProductCategory(name: string) {
  const normalized = normalizeName(name);
  const section = accountingTemplate.find(([, names]) =>
    names.split('|').some((item) => {
      const candidate = normalizeName(item);
      return candidate === normalized || candidate.includes(normalized) || normalized.includes(candidate);
    }),
  );
  return section?.[0] || 'รอจัดหมวดหมู่';
}
function printOperationalReport() {
  document.body.classList.add('operations-print');
  let cleaned = false;
  const cleanup = () => {
    if (cleaned) return;
    cleaned = true;
    document.body.classList.remove('operations-print');
    window.removeEventListener('afterprint', cleanup);
    printMedia.removeEventListener('change', handlePrintChange);
  };
  const handlePrintChange = (event: MediaQueryListEvent) => {
    if (!event.matches) cleanup();
  };
  const printMedia = window.matchMedia('print');
  window.addEventListener('afterprint', cleanup);
  printMedia.addEventListener('change', handlePrintChange);
  // Let the browser apply the A4 report styles before it captures the page.
  window.requestAnimationFrame(() =>
    window.requestAnimationFrame(() => window.print()),
  );
}

function guessUnit(name: string) {
  const text = name.toLowerCase();
  if (text.includes('cc') || text.includes('ml')) return text.includes('ml') ? 'ml' : 'cc';
  if (text.includes('unit') || text.includes('s.u')) return 'Unit';
  if (text.includes('ขวด') || text.includes('vial')) return 'ขวด';
  if (text.includes('กล่อง')) return 'กล่อง';
  if (text.includes('amp')) return 'Amp.';
  if (text.includes('line')) return 'line';
  return 'อัน';
}

function formatExcelValue(value: unknown) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'number' && value > 20000) {
    const date = XLSX.SSF.parse_date_code(value);
    return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
  }
  return String(value || '').trim();
}
