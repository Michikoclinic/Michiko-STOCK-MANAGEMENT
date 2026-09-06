'use client';
import { useEffect, useState } from 'react';
export function ThaiDateInput({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  const [year, month, day] = value ? value.split('-') : ['', '', ''];
  const [parts, setParts] = useState({ day, month, year });
  useEffect(() => setParts({ day, month, year }), [value]);
  const update = (key: 'day' | 'month' | 'year', next: string) => { const result = { ...parts, [key]: next }; setParts(result); if (result.day && result.month && result.year) onChange(`${result.year}-${result.month}-${result.day}`); };
  const years = Array.from({ length: 12 }, (_, index) => String(new Date().getFullYear() - 2 + index));
  return <div className="thai-date-input" aria-label="วัน เดือน ปี"><select value={parts.day} onChange={(e) => update('day', e.target.value)}><option value="">วัน</option>{Array.from({ length: 31 }, (_, i) => String(i + 1).padStart(2, '0')).map((x) => <option key={x}>{x}</option>)}</select><select value={parts.month} onChange={(e) => update('month', e.target.value)}><option value="">เดือน</option>{Array.from({ length: 12 }, (_, i) => String(i + 1).padStart(2, '0')).map((x) => <option key={x}>{x}</option>)}</select><select value={parts.year} onChange={(e) => update('year', e.target.value)}><option value="">ปี</option>{years.map((x) => <option key={x} value={x}>{Number(x) + 543}</option>)}</select></div>;
}
