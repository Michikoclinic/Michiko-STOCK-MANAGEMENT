'use client';

import { useEffect, useRef, useState } from 'react';
import { supabase } from '@/lib/supabase';

export function useSharedStored<T>(key: string, initial: T) {
  const storageKey = `stock:${key}`;
  const [value, setValue] = useState<T>(initial);
  const ready = useRef(false);
  const userId = useRef<string | null>(null);

  useEffect(() => {
    let active = true;
    const load = async () => {
      const localRaw = localStorage.getItem(key);
      let localValue = initial;
      if (localRaw) try { localValue = JSON.parse(localRaw) as T; } catch {}
      const { data: { session } } = await supabase.auth.getSession();
      if (!active || !session?.user) return;
      userId.current = session.user.id;
      const { data } = await supabase.from('smart_ledger_state').select('payload').eq('storage_key', storageKey).maybeSingle();
      if (!active) return;
      if (data?.payload !== undefined) setValue(data.payload as T);
      else {
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
      if (active && row?.payload !== undefined) setValue(row.payload);
    }).subscribe();
    return () => { active = false; authListener.subscription.unsubscribe(); void supabase.removeChannel(channel); };
  }, [key]);

  useEffect(() => {
    localStorage.setItem(key, JSON.stringify(value));
    if (!ready.current || !userId.current) return;
    const timer = window.setTimeout(() => {
      void supabase.from('smart_ledger_state').upsert({ storage_key: storageKey, payload: value, updated_by: userId.current, updated_at: new Date().toISOString() });
    }, 350);
    return () => window.clearTimeout(timer);
  }, [key, storageKey, value]);

  return [value, setValue] as const;
}
