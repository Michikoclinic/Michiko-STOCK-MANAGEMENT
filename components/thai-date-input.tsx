'use client';
import { CalendarDays, X } from 'lucide-react';
export function ThaiDateInput({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  return <div className="thai-date-input calendar-date-input" aria-label="เลือกวันที่จากปฏิทิน">
    <CalendarDays size={16} />
    <input type="date" value={value} onChange={(e) => onChange(e.target.value)} aria-label="วัน เดือน ปี" />
    {value && <button type="button" onClick={() => onChange('')} aria-label="แสดงทุกวัน"><X size={14} /></button>}
  </div>;
}
