'use client';

import { useEffect, useRef, useState } from 'react';
import { supabase } from '@/lib/supabase';

export function useSharedStored<T>(key: string, initial: T) {
  const storageKey = `stock:${key}`;
  const [value, setValue] = useState<T>(initial);
  const ready = useRef(false);
  const userId = useRef<string | null>(null);
  const lastSaved = useRef('');

  useEffect(() => {
    let active = true;
    const load = async () => {
      const localRaw = localStorage.getItem(key);
      const hasPendingLocalChange = localStorage.getItem(`${key}:pending`) === '1';
      let localValue = initial;
      if (localRaw) try { localValue = JSON.parse(localRaw) as T; } catch {}
      const { data: { session } } = await supabase.auth.getSession();
      if (!active || !session?.user) return;
      userId.current = session.user.id;
      if (hasPendingLocalChange && localRaw) {
        const serialized = JSON.stringify(localValue);
        lastSaved.current = serialized;
        setValue(localValue);
        ready.current = true;
        const { error } = await supabase.from('smart_ledger_state').upsert({ storage_key: storageKey, payload: localValue, updated_by: session.user.id, updated_at: new Date().toISOString() });
        if (!error && localStorage.getItem(key) === serialized) localStorage.removeItem(`${key}:pending`);
        return;
      }
      const { data } = await supabase.from('smart_ledger_state').select('payload').eq('storage_key', storageKey).maybeSingle();
      if (!active) return;
      if (data?.payload !== undefined) {
        lastSaved.current = JSON.stringify(data.payload);
        localStorage.setItem(key, lastSaved.current);
        setValue(data.payload as T);
      }
      else {
        lastSaved.current = JSON.stringify(localValue);
        setValue(localValue);
        await supabase.from('smart_ledger_state').upsert({ storage_key: storageKey, payload: localValue, updated_by: session.user.id, updated_at: new Date().toISOString() });
      }
      ready.current = true;
    };
    void load();
    const { data: authListener } = supabase.auth.onAuthStateChange((event) => {
      if (event === 'SIGNED_IN') void load();
      if (event === 'SIGNED_OUT') { ready.current = false; userId.current = null; }
    });
    const channel = supabase.channel(`sync-${storageKey}`).on('postgres_changes', { event: '*', schema: 'public', table: 'smart_ledger_state', filter: `storage_key=eq.${storageKey}` }, (event) => {
      const row = event.new as { payload?: T };
      if (active && row?.payload !== undefined) {
        lastSaved.current = JSON.stringify(row.payload);
        localStorage.setItem(key, lastSaved.current);
        setValue(row.payload);
      }
    }).subscribe();
    return () => { active = false; authListener.subscription.unsubscribe(); void supabase.removeChannel(channel); };
  }, [key]);

  useEffect(() => {
    const serialized = JSON.stringify(value);
    localStorage.setItem(key, serialized);
    if (!ready.current || !userId.current) return;
    if (serialized === lastSaved.current) return;
    lastSaved.current = serialized;
    localStorage.setItem(`${key}:pending`, '1');
    void supabase.from('smart_ledger_state').upsert({ storage_key: storageKey, payload: value, updated_by: userId.current, updated_at: new Date().toISOString() }).then(({ error }) => {
      if (!error && localStorage.getItem(key) === serialized) localStorage.removeItem(`${key}:pending`);
    });
  }, [key, storageKey, value]);

  return [value, setValue] as const;
}
