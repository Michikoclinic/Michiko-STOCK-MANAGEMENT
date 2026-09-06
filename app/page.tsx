'use client';
import { useEffect, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { OperationalPage } from './operations';
import {
  LayoutDashboard,
  PackageCheck,
  ClipboardPlus,
  ArrowLeftRight,
  HandCoins,
  PackageMinus,
  ScanLine,
  Boxes,
  Tags,
  ChartNoAxesCombined,
  History,
  Settings,
  Bell,
  Search,
  ChevronDown,
  CalendarDays,
  Upload,
  CircleAlert,
  ArrowUpRight,
  Plus,
  X,
  FileSpreadsheet,
  CheckCircle2,
  Users,
  Stethoscope,
  Scissors,
  Trash2,
  Pencil,
  Save,
  RotateCcw,
  Printer,
} from 'lucide-react';

type Item = {
  id: string;
  name: string;
  qty: number;
  unit: string;
  source: 'report' | 'manual';
};
type Case = {
  id: string;
  date: string;
  hn: string;
  patient: string;
  doctor: string;
  program: string;
  assistants: string[];
  items: Item[];
  backdated: boolean;
  needsReview: boolean;
};
type Preview = {
  fileName: string;
  hash: string;
  date: string;
  cases: Case[];
  rows: number;
  needsReview: number;
  duplicate: boolean;
};
type ToolRegistry = {
  registerTool: (
    tool: {
      name: string;
      title: string;
      description: string;
      inputSchema: object;
      annotations: object;
      execute: (input: unknown) => unknown;
    },
    options: { signal: AbortSignal },
  ) => void | Promise<void>;
};
const staff = ['แพรว', 'เมย์', 'น้ำ', 'ปุ้ย', 'เฟิร์น', 'ออม'];
const menu = [
  ['Dashboard', LayoutDashboard],
  ['Stock รายวัน', PackageCheck],
  ['รับเข้า', ClipboardPlus],
  ['โอนย้ายสาขา', ArrowLeftRight],
  ['ยืม / คืน', HandCoins],
  ['เบิกออก', PackageMinus],
  ['ตรวจนับ Stock', ScanLine],
  ['Stock คงคลัง', Boxes],
  ['ทะเบียนสินค้า', Tags],
  ['รายงาน', ChartNoAxesCombined],
  ['ประวัติการเคลื่อนไหว', History],
  ['ตั้งค่า', Settings],
] as const;
const seed: Case[] = [
  {
    id: 'c1',
    date: '2026-09-06',
    hn: 'EM001234',
    patient: 'คุณพิมพ์ชนก สุขใจ',
    doctor: 'หมอกิ๊ฟ',
    program: 'Filler',
    assistants: ['แพรว'],
    backdated: false,
    needsReview: false,
    items: [
      {
        id: 'i1',
        name: 'Restylane Kysse',
        qty: 1,
        unit: 'cc',
        source: 'report',
      },
      { id: 'i2', name: 'Cannula 22G', qty: 1, unit: 'อัน', source: 'manual' },
    ],
  },
  {
    id: 'c2',
    date: '2026-09-06',
    hn: 'EM001208',
    patient: 'คุณณัฐชา กิตติกุล',
    doctor: 'หมอเบนซ์',
    program: 'Botox',
    assistants: ['เมย์'],
    backdated: false,
    needsReview: false,
    items: [
      { id: 'i3', name: 'Bienox', qty: 1, unit: 'ขวด', source: 'report' },
      { id: 'i4', name: 'Syringe 1 ml', qty: 2, unit: 'อัน', source: 'manual' },
    ],
  },
];
const uid = () => crypto.randomUUID?.() || Math.random().toString(36).slice(2);

export default function Home() {
  const [page, setPage] = useState<string>(() =>
    typeof window !== 'undefined' && new URLSearchParams(window.location.search).has('print')
      ? 'Stock รายวัน'
      : 'Dashboard',
  );
  const [cases, setCases] = useState<Case[]>(seed);
  const [preview, setPreview] = useState<Preview | null>(null);
  const [toast, setToast] = useState('');
  const fileRef = useRef<HTMLInputElement>(null);
  useEffect(() => {
    const saved = localStorage.getItem('michiko-stock-cases');
    if (saved)
      try {
        setCases(JSON.parse(saved));
      } catch {}
  }, []);
  useEffect(() => {
    localStorage.setItem('michiko-stock-cases', JSON.stringify(cases));
  }, [cases]);
  useEffect(() => {
    const registry = (document as Document & { modelContext?: ToolRegistry })
      .modelContext;
    if (!registry?.registerTool) return;
    const lifecycle = new AbortController();
    void Promise.resolve(
      registry.registerTool(
        {
          name: 'add_manual_stock_to_case',
          title: 'เพิ่ม Stock เองในเคส',
          description:
            'เพิ่มรายการ Stock ที่ไม่มีในรายงานเข้าเคส Stock รายวัน โดยระบุรหัสเคส ชื่อสินค้า จำนวน และหน่วย',
          inputSchema: {
            type: 'object',
            properties: {
              caseId: { type: 'string' },
              name: { type: 'string' },
              quantity: { type: 'number', exclusiveMinimum: 0 },
              unit: { type: 'string' },
            },
            required: ['caseId', 'name', 'quantity', 'unit'],
            additionalProperties: false,
          },
          annotations: { readOnlyHint: false, untrustedContentHint: false },
          execute(input) {
            const value = input as {
              caseId?: string;
              name?: string;
              quantity?: number;
              unit?: string;
            };
            if (
              !value.caseId ||
              !value.name ||
              !value.unit ||
              !value.quantity ||
              value.quantity <= 0
            )
              throw new Error('ข้อมูลรายการ Stock ไม่ครบ');
            let found = false;
            setCases((current) =>
              current.map((c) => {
                if (c.id !== value.caseId) return c;
                found = true;
                return {
                  ...c,
                  items: [
                    ...c.items,
                    {
                      id: uid(),
                      name: value.name!,
                      qty: value.quantity!,
                      unit: value.unit!,
                      source: 'manual',
                    },
                  ],
                };
              }),
            );
            if (!found) throw new Error('ไม่พบเคสที่ระบุ');
            return {
              status: 'added',
              caseId: value.caseId,
              name: value.name,
              quantity: value.quantity,
              unit: value.unit,
            };
          },
        },
        { signal: lifecycle.signal },
      ),
    );
    return () => lifecycle.abort();
  }, []);
  const notify = (s: string) => {
    setToast(s);
    setTimeout(() => setToast(''), 2800);
  };
  const selectPage = (label: string) => {
    setPage(label);
  };
  const importFile = async (file: File) => {
    try {
      const buf = await file.arrayBuffer();
      const hash = await crypto.subtle.digest('SHA-256', buf).then((x) =>
        Array.from(new Uint8Array(x))
          .map((b) => b.toString(16).padStart(2, '0'))
          .join(''),
      );
      const workbook = XLSX.read(buf, { type: 'array', cellDates: true });
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
        workbook.Sheets[workbook.SheetNames[0]],
        { defval: '' },
      );
      const parsed = parseRows(rows);
      const imported = JSON.parse(
        localStorage.getItem('michiko-stock-imports') || '[]',
      ) as string[];
      setPreview({
        fileName: file.name,
        hash,
        date: parsed.date,
        cases: parsed.cases,
        rows: parsed.rows,
        needsReview: parsed.needsReview,
        duplicate: imported.includes(hash),
      });
    } catch {
      notify('อ่านไฟล์ไม่สำเร็จ กรุณาตรวจรูปแบบ Excel หรือ CSV');
    }
  };
  const confirmImport = () => {
    if (!preview || preview.duplicate) return;
    setCases((prev) => [...preview.cases, ...prev]);
    const hashes = JSON.parse(
      localStorage.getItem('michiko-stock-imports') || '[]',
    );
    localStorage.setItem(
      'michiko-stock-imports',
      JSON.stringify([...hashes, preview.hash]),
    );
    setPreview(null);
    setPage('Stock รายวัน');
    notify('นำเข้ารายงานเรียบร้อยแล้ว');
  };
  const update = (next: Case) =>
    setCases((x) => x.map((c) => (c.id === next.id ? next : c)));
  const removeCase = (id: string) =>
    setCases((x) => x.filter((c) => c.id !== id));
  return (
    <main className="app-shell">
      <Sidebar page={page} select={selectPage} />
      <section className="workspace">
        <Topbar />
        <div className="content">
          {page === 'Dashboard' ? (
            <Dashboard
              cases={cases}
              openImport={() => fileRef.current?.click()}
              goDaily={() => setPage('Stock รายวัน')}
            />
          ) : page === 'Stock รายวัน' ? (
            <Daily
              cases={cases}
              update={update}
              remove={removeCase}
              openImport={() => fileRef.current?.click()}
              notify={notify}
            />
          ) : (
            <OperationalPage page={page} notify={notify} />
          )}
        </div>
      </section>
      <input
        ref={fileRef}
        hidden
        type="file"
        accept=".xlsx,.xls,.csv"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) importFile(f);
          e.target.value = '';
        }}
      />
      {preview && (
        <ImportModal
          data={preview}
          close={() => setPreview(null)}
          confirm={confirmImport}
          viewExisting={() => {
            setPreview(null);
            setPage('Stock รายวัน');
          }}
        />
      )}
      {toast && (
        <div className="toast">
          <CheckCircle2 size={18} />
          {toast}
        </div>
      )}
    </main>
  );
}

function Sidebar({
  page,
  select,
}: {
  page: string;
  select: (x: string) => void;
}) {
  return (
    <aside className="sidebar">
      <div className="brand">
        <span className="brand-logo-crop"><img src="/michiko-logo.png" alt="MICHIKO Aesthetics" /></span>
        <span>STOCK MANAGEMENT</span>
      </div>
      <nav>
        {menu.map(([label, Icon]) => (
          <button
            key={label}
            onClick={() => select(label)}
            className={page === label ? 'active' : ''}
          >
            <Icon size={18} />
            <span>{label}</span>
            {label === 'ยืม / คืน' && <em>3</em>}
          </button>
        ))}
      </nav>
      <div className="sidebar-foot">
        <div className="avatar">พ</div>
        <div>
          <strong>พิชญาภรณ์</strong>
          <span>ผู้ดูแลระบบ</span>
        </div>
        <ChevronDown size={16} />
      </div>
    </aside>
  );
}
function Topbar() {
  return (
    <header className="topbar">
      <label className="search">
        <Search size={17} />
        <input placeholder="ค้นหาสินค้า ลูกค้า หรือ HN" />
      </label>
      <div className="top-actions">
        <button aria-label="การแจ้งเตือน">
          <Bell size={19} />
          <i />
        </button>
        <div className="branch">
          <span>สาขาปัจจุบัน</span>
          <strong>
            MICHIKO สาขา Emsphere <ChevronDown size={14} />
          </strong>
        </div>
      </div>
    </header>
  );
}
function Dashboard({
  cases,
  openImport,
  goDaily,
}: {
  cases: Case[];
  openImport: () => void;
  goDaily: () => void;
}) {
  const items = cases.reduce((n, c) => n + c.items.length, 0);
  return (
    <>
      <div className="page-head">
        <div>
          <p>ภาพรวมคลินิก</p>
          <h1>
            สวัสดีตอนเช้า คุณครีม <span>✦</span>
          </h1>
          <small>ติดตาม Stock และรายการที่ต้องดูแลในวันนี้</small>
        </div>
        <div className="filters">
          <button>
            <CalendarDays size={17} /> กันยายน 2026 <ChevronDown size={14} />
          </button>
          <button className="primary" onClick={openImport}>
            <Upload size={17} /> นำเข้ารายงาน
          </button>
        </div>
      </div>
      <div className="metrics">
        <Metric
          icon={<PackageCheck />}
          title="Stock รายวันวันนี้"
          value={String(items)}
          unit="รายการ"
          note="อัปเดตจากข้อมูลล่าสุด"
        />
        <Metric
          icon={<ClipboardPlus />}
          title="จำนวนเคสวันนี้"
          value={String(cases.length)}
          unit="เคส"
          note="พร้อมตรวจสอบรายละเอียด"
        />
        <Metric
          icon={<CircleAlert />}
          title="Stock ต่ำ"
          value="6"
          unit="รายการ"
          note="ควรตรวจสอบ 2 รายการ"
        />
        <Metric
          icon={<HandCoins />}
          title="ยืมค้าง"
          value="3"
          unit="รายการ"
          note="เกินกำหนด 1 รายการ"
        />
      </div>
      <div className="dashboard-grid">
        <article className="panel daily">
          <PanelHead
            title="Stock รายวันล่าสุด"
            sub={`${cases.length} เคส`}
            action="ดูทั้งหมด"
            click={goDaily}
          />
          {cases.slice(0, 3).map((c) => (
            <CaseRow key={c.id} data={c} />
          ))}
        </article>
        <article className="panel pending">
          <PanelHead title="รายการที่ต้องดูแล" sub="อัปเดตล่าสุดเมื่อสักครู่" />
          <Attention
            icon={<CircleAlert />}
            title="Stock ต่ำกว่า Minimum"
            sub="Restylane Kysse เหลือ 2 cc"
            count="ดู"
          />
          <Attention
            icon={<HandCoins />}
            title="รายการยืมค้าง"
            sub="Emsphere ยืมจากพหลโยธิน"
            count="3"
          />
          <Attention
            icon={<ScanLine />}
            title="รอตรวจสอบ"
            sub={`${cases.filter((c) => c.needsReview).length} เคสต้องจัดกลุ่ม`}
            count={String(cases.filter((c) => c.needsReview).length)}
          />
        </article>
      </div>
      <Usage />
    </>
  );
}
function Daily({
  cases,
  update,
  remove,
  openImport,
  notify,
}: {
  cases: Case[];
  update: (c: Case) => void;
  remove: (id: string) => void;
  openImport: () => void;
  notify: (s: string) => void;
}) {
  const [query, setQuery] = useState('');
  const [date, setDate] = useState('');
  const [showPrint, setShowPrint] = useState(
    () => typeof window !== 'undefined' && new URLSearchParams(window.location.search).has('print'),
  );
  const filtered = cases.filter(
    (c) =>
      (!date || c.date === date) &&
      `${c.hn} ${c.patient} ${c.doctor} ${c.program}`
        .toLowerCase()
        .includes(query.toLowerCase()),
  );
  return (
    <>
      <div className="page-head daily-head">
        <div>
          <p>งานประจำวัน</p>
          <h1>Stock รายวัน</h1>
          <small>ตรวจรายการที่ใช้กับลูกค้า แพทย์ ผู้ช่วย และโปรแกรม</small>
        </div>
        <div className="daily-head-actions">
          <button className="print-button" onClick={() => setShowPrint(true)}>
            <Printer size={17} /> พิมพ์รายงานประจำวัน
          </button>
          <button className="import-big" onClick={openImport}>
            <Upload size={17} /> นำเข้ารายงาน Stock Movement
          </button>
        </div>
      </div>
      <div className="daily-toolbar">
        <label>
          <Search size={17} />
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="ค้นหาชื่อ, HN, แพทย์ หรือโปรแกรม"
          />
        </label>
        <input
          type="date"
          value={date}
          onChange={(e) => setDate(e.target.value)}
        />
        <span>
          พบ {filtered.length} เคส ·{' '}
          {filtered.reduce((n, c) => n + c.items.length, 0)} รายการ
        </span>
      </div>
      {filtered.length ? (
        <div className="case-list">
          {filtered.map((c, i) => (
            <CaseCard
              key={c.id}
              data={c}
              number={i + 1}
              update={update}
              remove={remove}
              notify={notify}
            />
          ))}
        </div>
      ) : (
        <div className="empty">
          <PackageCheck />
          <h3>ยังไม่พบรายการ</h3>
          <p>ลองเปลี่ยนตัวกรอง หรือนำเข้ารายงาน Stock Movement</p>
          <button onClick={openImport}>นำเข้ารายงาน</button>
        </div>
      )}
      {showPrint && (
        <PrintPreview cases={filtered} close={() => setShowPrint(false)} />
      )}
    </>
  );
}

function PrintPreview({ cases, close }: { cases: Case[]; close: () => void }) {
  const reportDate = cases[0]?.date || new Date().toISOString().slice(0, 10);
  const totalItems = cases.reduce((sum, c) => sum + c.items.length, 0);
  return (
    <div className="print-preview-back">
      <div className="print-preview-shell">
        <div className="print-preview-toolbar">
          <div>
            <strong>ตัวอย่างก่อนพิมพ์</strong>
            <span>กระดาษ A4 · สามารถเลือก Save as PDF ได้</span>
          </div>
          <button onClick={close}><X size={17} /> ปิด</button>
          <button className="primary" onClick={() => window.print()}>
            <Printer size={17} /> พิมพ์ / Save as PDF
          </button>
        </div>
        <article className="daily-print-sheet">
          <header className="print-doc-head">
            <div className="print-logo">
              <span><img src="/michiko-logo.png" alt="MICHIKO Aesthetics" /></span>
            </div>
            <div className="print-title"><h1>รายงาน Stock รายวัน</h1><p>DAILY STOCK USAGE REPORT</p></div>
            <div className="print-doc-no"><span>เลขที่เอกสาร</span><strong>DS-{reportDate.replaceAll('-', '')}</strong></div>
          </header>
          <section className="print-info">
            <div><span>สาขา</span><strong>MICHIKO สาขา Emsphere</strong></div>
            <div><span>วันที่รายงาน</span><strong>{formatDate(reportDate)}</strong></div>
            <div><span>จำนวนเคส</span><strong>{cases.length} เคส</strong></div>
            <div><span>รายการ Stock</span><strong>{totalItems} รายการ</strong></div>
          </section>
          <div className="print-cases">
            {cases.map((c, index) => (
              <section className="print-case" key={c.id}>
                <div className="print-case-head">
                  <b>{index + 1}</b>
                  <div><strong>{c.patient || 'ไม่พบชื่อผู้ป่วย'}</strong><span>HN {c.hn || '-'}</span></div>
                  {c.backdated && <em>บันทึกย้อนหลัง</em>}
                </div>
                <div className="print-case-meta">
                  <div><span>แพทย์</span><strong>{c.doctor || '-'}</strong></div>
                  <div><span>ผู้ช่วย</span><strong>{c.assistants.join(', ') || '-'}</strong></div>
                  <div><span>โปรแกรม / บริการ</span><strong>{c.program || '-'}</strong></div>
                </div>
                <table><thead><tr><th>ลำดับ</th><th>รายการ Stock</th><th>จำนวน</th><th>หน่วย</th><th>แหล่งข้อมูล</th></tr></thead>
                  <tbody>{c.items.map((item, itemIndex) => <tr key={item.id}><td>{itemIndex + 1}</td><td>{item.name}</td><td>{item.qty}</td><td>{item.unit}</td><td>{item.source === 'report' ? 'จากรายงาน' : 'เพิ่มเอง'}</td></tr>)}</tbody>
                </table>
              </section>
            ))}
          </div>
          <footer className="print-signatures">
            <div><span>ผู้ทำรายการ</span><i /><strong>พิชญาภรณ์</strong><small>วันที่ ____ / ____ / ______</small></div>
            <div><span>ผู้ตรวจสอบ</span><i /><strong>________________________</strong><small>วันที่ ____ / ____ / ______</small></div>
          </footer>
          <div className="print-footer"><span>MICHIKO Stock Management</span><span>พิมพ์เมื่อ {new Intl.DateTimeFormat('th-TH',{dateStyle:'medium',timeStyle:'short'}).format(new Date())}</span></div>
        </article>
      </div>
    </div>
  );
}
function CaseCard({
  data,
  number,
  update,
  remove,
  notify,
}: {
  data: Case;
  number: number;
  update: (c: Case) => void;
  remove: (id: string) => void;
  notify: (s: string) => void;
}) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(data);
  const [adding, setAdding] = useState(false);
  const [newItem, setNewItem] = useState({ name: '', qty: 1, unit: 'อัน' });
  useEffect(() => setDraft(data), [data]);
  const save = () => {
    update(draft);
    setEditing(false);
    notify('บันทึกการแก้ไขเคสแล้ว');
  };
  const addItem = () => {
    if (!newItem.name.trim()) return;
    update({
      ...data,
      items: [...data.items, { id: uid(), ...newItem, source: 'manual' }],
    });
    setNewItem({ name: '', qty: 1, unit: 'อัน' });
    setAdding(false);
  };
  const split = (item: Item) => {
    update({ ...data, items: data.items.filter((x) => x.id !== item.id) });
    notify('แยกรายการออกจากเคสแล้ว');
  };
  return (
    <article className={`case-card ${data.needsReview ? 'review-case' : ''}`}>
      <div className="case-card-head">
        <div className="case-number">{number}</div>
        <div>
          <div className="case-title">
            <h3>{data.patient || 'ไม่พบชื่อผู้ป่วย'}</h3>
            <span>{data.hn || 'ไม่มี HN'}</span>
            {data.backdated && <b>บันทึกย้อนหลัง</b>}
            {data.needsReview && <b className="review">รอตรวจสอบ</b>}
          </div>
          <p>{formatDate(data.date)}</p>
        </div>
        <div className="case-actions">
          {editing ? (
            <>
              <button
                onClick={() => {
                  setDraft(data);
                  setEditing(false);
                }}
              >
                <X size={15} />
                ยกเลิก
              </button>
              <button className="save" onClick={save}>
                <Save size={15} />
                บันทึก
              </button>
            </>
          ) : (
            <>
              <button onClick={() => setEditing(true)}>
                <Pencil size={15} />
                แก้ไขเคส
              </button>
              <button
                className="icon-danger"
                aria-label="ลบเคส"
                onClick={() => confirm('ลบเคสนี้หรือไม่?') && remove(data.id)}
              >
                <Trash2 size={15} />
              </button>
            </>
          )}
        </div>
      </div>
      <div className="case-meta">
        <label>
          <Stethoscope size={16} />
          <span>แพทย์</span>
          {editing ? (
            <input
              value={draft.doctor}
              onChange={(e) => setDraft({ ...draft, doctor: e.target.value })}
            />
          ) : (
            <strong>{data.doctor || 'รอตรวจสอบ'}</strong>
          )}
        </label>
        <label>
          <ClipboardPlus size={16} />
          <span>โปรแกรม</span>
          {editing ? (
            <input
              value={draft.program}
              onChange={(e) => setDraft({ ...draft, program: e.target.value })}
            />
          ) : (
            <strong>{data.program || 'รอตรวจสอบ'}</strong>
          )}
        </label>
        <label className="assist">
          <Users size={16} />
          <span>ผู้ช่วย</span>
          <div>
            {(editing ? draft : data).assistants.map((a) => (
              <b key={a}>
                {a}
                {editing && (
                  <button
                    onClick={() =>
                      setDraft({
                        ...draft,
                        assistants: draft.assistants.filter((x) => x !== a),
                      })
                    }
                  >
                    <X size={11} />
                  </button>
                )}
              </b>
            ))}
            {editing && (
              <select
                value=""
                onChange={(e) =>
                  e.target.value &&
                  setDraft({
                    ...draft,
                    assistants: draft.assistants.concat(e.target.value),
                  })
                }
              >
                <option value="">+ เพิ่มผู้ช่วย</option>
                {staff.map((s) => (
                  <option key={s}>{s}</option>
                ))}
              </select>
            )}
            {!editing && !data.assistants.length && <em>ยังไม่ได้เลือก</em>}
          </div>
        </label>
      </div>
      <div className="stock-table">
        <div className="stock-head">
          <span>รายการ Stock</span>
          <span>จำนวน</span>
          <span>แหล่งข้อมูล</span>
          <span />
        </div>
        {data.items.map((item) => (
          <div className="stock-line" key={item.id}>
            <span>
              <Boxes size={15} />
              {item.name}
            </span>
            <strong>
              {item.qty} {item.unit}
            </strong>
            <em className={item.source}>
              {item.source === 'report' ? 'จากรายงาน' : 'เพิ่มเอง'}
            </em>
            <div>
              <button title="แยกเป็นเคสใหม่" onClick={() => split(item)}>
                <Scissors size={14} />
              </button>
              <button
                title="ลบรายการ"
                onClick={() =>
                  update({
                    ...data,
                    items: data.items.filter((x) => x.id !== item.id),
                  })
                }
              >
                <Trash2 size={14} />
              </button>
            </div>
          </div>
        ))}
      </div>
      {adding ? (
        <div className="add-stock">
          <input
            autoFocus
            placeholder="ชื่อสินค้า"
            value={newItem.name}
            onChange={(e) => setNewItem({ ...newItem, name: e.target.value })}
          />
          <input
            type="number"
            min="0.01"
            step="0.01"
            value={newItem.qty}
            onChange={(e) =>
              setNewItem({ ...newItem, qty: Number(e.target.value) })
            }
          />
          <select
            value={newItem.unit}
            onChange={(e) => setNewItem({ ...newItem, unit: e.target.value })}
          >
            {['อัน', 'cc', 'ขวด', 'กล่อง', 'ชิ้น', 'ml', 'เส้น'].map((x) => (
              <option key={x}>{x}</option>
            ))}
          </select>
          <button onClick={addItem}>เพิ่ม</button>
          <button onClick={() => setAdding(false)}>ยกเลิก</button>
        </div>
      ) : (
        <button className="add-stock-btn" onClick={() => setAdding(true)}>
          <Plus size={15} /> เพิ่มรายการ Stock
        </button>
      )}
    </article>
  );
}

function ImportModal({
  data,
  close,
  confirm,
  viewExisting,
}: {
  data: Preview;
  close: () => void;
  confirm: () => void;
  viewExisting: () => void;
}) {
  return (
    <div className="modal-back">
      <section className="import-modal">
        <button className="modal-x" onClick={close}>
          <X />
        </button>
        {data.duplicate ? (
          <>
            <div className="modal-icon danger">
              <RotateCcw />
            </div>
            <h2>รายงานนี้อาจเคยนำเข้าแล้ว</h2>
            <p className="modal-desc">ระบบตรวจพบไฟล์เดียวกัน จึงยังไม่ตัด Stock ซ้ำ</p>
            <div className="duplicate-box">
              <FileSpreadsheet />
              <div>
                <strong>{data.fileName}</strong>
                <span>ตรวจจากเนื้อหาไฟล์ ไม่ใช่เพียงชื่อไฟล์</span>
              </div>
            </div>
            <div className="modal-actions">
              <button onClick={close}>ยกเลิก</button>
              <button className="primary" onClick={viewExisting}>
                ดูข้อมูลเดิม
              </button>
            </div>
          </>
        ) : (
          <>
            <div className="modal-icon">
              <FileSpreadsheet />
            </div>
            <h2>ตรวจสอบข้อมูลก่อนนำเข้า</h2>
            <p className="modal-desc">{data.fileName}</p>
            <div className="preview-date">
              <CalendarDays />
              <div>
                <span>วันที่รายงาน</span>
                <strong>{formatDate(data.date)}</strong>
              </div>
              {data.cases.some((c) => c.backdated) && <b>บันทึกย้อนหลัง</b>}
            </div>
            <div className="preview-stats">
              <div>
                <strong>{new Set(data.cases.map((c) => c.hn)).size}</strong>
                <span>ลูกค้า</span>
              </div>
              <div>
                <strong>{data.rows}</strong>
                <span>รายการ Stock</span>
              </div>
              <div className={data.needsReview ? 'warn' : ''}>
                <strong>{data.needsReview}</strong>
                <span>รอตรวจสอบ</span>
              </div>
            </div>
            {data.needsReview > 0 && (
              <p className="review-note">
                <CircleAlert /> ข้อมูลที่ไม่มั่นใจจะถูกทำเครื่องหมาย “รอตรวจสอบ”
                และสามารถแก้ได้หลังนำเข้า
              </p>
            )}
            <div className="modal-actions">
              <button onClick={close}>ยกเลิก</button>
              <button className="primary" onClick={confirm}>
                <CheckCircle2 /> ยืนยันนำเข้า
              </button>
            </div>
          </>
        )}
      </section>
    </div>
  );
}

function parseRows(rows: Record<string, unknown>[]) {
  const val = (r: Record<string, unknown>, names: string[]) => {
    const key = Object.keys(r).find((k) =>
      names.some((n) => k.toLowerCase().replace(/\s/g, '').includes(n)),
    );
    return key ? String(r[key] ?? '').trim() : '';
  };
  const valid = rows
    .map((r) => ({
      patient: val(r, ['ชื่อผู้ป่วย', 'ชื่อลูกค้า', 'patient', 'customer']),
      hn: val(r, ['hn', 'หมายเลขผู้ป่วย']),
      doctor: val(r, ['แพทย์', 'doctor']),
      stock: val(r, ['ชื่อยา', 'ชื่อstock', 'สินค้า', 'product', 'item']),
      qty: val(r, ['จำนวน', 'qty', 'quantity']),
      unit: val(r, ['หน่วย', 'unit']),
      program: val(r, ['รายละเอียดบริการ', 'บริการ', 'program', 'service']),
      date: val(r, ['วันที่', 'date']),
    }))
    .filter((r) => r.stock && (r.hn || r.patient));
  const date =
    normalizeDate(valid.find((r) => r.date)?.date) ||
    new Date().toISOString().slice(0, 10);
  const groups = new Map<string, Case>();
  for (const r of valid) {
    const key = [r.hn || r.patient, r.doctor, r.program].join('|');
    if (!groups.has(key))
      groups.set(key, {
        id: uid(),
        date,
        hn: r.hn,
        patient: r.patient,
        doctor: r.doctor,
        program: r.program,
        assistants: [],
        items: [],
        backdated: date < new Date().toISOString().slice(0, 10),
        needsReview: !r.hn || !r.doctor || !r.program,
      });
    groups
      .get(key)!
      .items.push({
        id: uid(),
        name: r.stock,
        qty: Math.abs(Number(String(r.qty).replace(/,/g, ''))) || 1,
        unit: r.unit || 'อัน',
        source: 'report',
      });
  }
  const cases = [...groups.values()];
  return {
    date,
    cases,
    rows: valid.length,
    needsReview: cases.filter((c) => c.needsReview).length,
  };
}
function normalizeDate(v: string) {
  if (!v) return '';
  const n = Number(v);
  if (Number.isFinite(n) && n > 20000) {
    const d = XLSX.SSF.parse_date_code(n);
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  const d = new Date(v);
  return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
}
function formatDate(s: string) {
  if (!s) return 'ไม่พบวันที่';
  return new Intl.DateTimeFormat('th-TH', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  }).format(new Date(`${s}T00:00:00`));
}
function Metric({
  icon,
  title,
  value,
  unit,
  note,
}: {
  icon: React.ReactNode;
  title: string;
  value: string;
  unit: string;
  note: string;
}) {
  return (
    <article className="metric">
      <div className="metric-icon">{icon}</div>
      <span>{title}</span>
      <strong>
        {value} <small>{unit}</small>
      </strong>
      <p>{note}</p>
    </article>
  );
}
function PanelHead({
  title,
  sub,
  action,
  click,
}: {
  title: string;
  sub: string;
  action?: string;
  click?: () => void;
}) {
  return (
    <div className="panel-head">
      <div>
        <h2>{title}</h2>
        <p>{sub}</p>
      </div>
      {action && (
        <button onClick={click}>
          {action}
          <ArrowUpRight size={15} />
        </button>
      )}
    </div>
  );
}
function CaseRow({ data }: { data: Case }) {
  return (
    <div className="case-row">
      <div className="patient">
        {data.patient.replace(/คุณ/g, '').slice(0, 2)}
      </div>
      <div className="case-info">
        <strong>{data.patient}</strong>
        <span>
          HN {data.hn} · {data.program}
        </span>
      </div>
      <div className="doctor">
        <span>แพทย์</span>
        <strong>{data.doctor}</strong>
      </div>
      <div className="stock-pills">
        {data.items.slice(0, 2).map((x) => (
          <span key={x.id}>
            {x.name}{' '}
            <b>
              {x.qty} {x.unit}
            </b>
          </span>
        ))}
      </div>
    </div>
  );
}
function Attention({
  icon,
  title,
  sub,
  count,
}: {
  icon: React.ReactNode;
  title: string;
  sub: string;
  count: string;
}) {
  return (
    <div className="attention">
      <div className="attention-icon">{icon}</div>
      <div>
        <strong>{title}</strong>
        <span>{sub}</span>
      </div>
      <button>{count}</button>
    </div>
  );
}
function Usage() {
  return (
    <article className="panel usage">
      <PanelHead
        title="สรุปการใช้ Stock เดือนนี้"
        sub="1–6 กันยายน 2026"
        action="ดูรายงาน"
      />
      <div className="usage-body">
        <div className="bars">
          {[
            ['Restylane Kysse', 76, '14 cc'],
            ['Bienox', 59, '8 ขวด'],
            ['Syringe 1 ml', 45, '26 อัน'],
            ['Cannula 22G', 34, '18 อัน'],
          ].map(([n, w, v]) => (
            <div className="bar-row" key={n}>
              <span>{n}</span>
              <div>
                <i style={{ width: `${w}%` }} />
              </div>
              <strong>{v}</strong>
            </div>
          ))}
        </div>
        <div className="donut">
          <div>
            <strong>168</strong>
            <span>รายการที่ใช้</span>
          </div>
        </div>
        <div className="legend">
          <p>
            <i /> Stock รายวัน <b>142</b>
          </p>
          <p>
            <i className="l2" /> เบิกทั่วไป <b>18</b>
          </p>
          <p>
            <i className="l3" /> ของเสีย <b>8</b>
          </p>
        </div>
      </div>
    </article>
  );
}
