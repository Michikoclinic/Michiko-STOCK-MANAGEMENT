'use client';
import { useRef } from 'react';
import { CalendarDays, X } from 'lucide-react';

function displayDate(value: string) {
  if (!value) return 'วัน/เดือน/ปี';
  const [year, month, day] = value.split('-');
  return `${day}/${month}/${Number(year) + 543}`;
}

export function ThaiDateInput({ value, onChange, allowClear = true }: { value: string; onChange: (value: string) => void; allowClear?: boolean }) {
  const inputRef = useRef<HTMLInputElement>(null);
  return <div className="thai-date-input calendar-date-input" aria-label="เลือกวันที่จากปฏิทิน">
    <button className="calendar-trigger" type="button" onClick={() => inputRef.current?.showPicker()} aria-label="เปิดปฏิทิน"><CalendarDays size={16} /></button>
    <span className="thai-date-value">{displayDate(value)}</span>
    <input ref={inputRef} type="date" lang="th-TH" value={value} onChange={(e) => onChange(e.target.value)} aria-label="วัน เดือน ปี" />
    {allowClear && value && <button type="button" onClick={() => onChange('')} aria-label="ล้างวันที่"><X size={14} /></button>}
  </div>;
}
