import { createClient } from "@supabase/supabase-js";

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!url || !anonKey) {
  // eslint-disable-next-line no-console
  console.warn(
    "ยังไม่ได้ตั้งค่า VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY ใน .env — ดูวิธีตั้งค่าในไฟล์ README.md"
  );
}

export const supabase = createClient(url, anonKey);

/**
 * AccountingApp.jsx was originally built for Claude's artifact environment,
 * which exposes a global `window.storage` key/value API (get/set/delete/list).
 * This polyfill re-implements the same interface on top of a single Supabase
 * table ("app_data") so the rest of the app's code did not need to change.
 *
 * Every signed-in family member reads/writes the same rows (see
 * supabase/schema.sql for the table + row-level-security policies), so
 * everyone shares one live dataset.
 */
window.storage = {
  async get(key) {
    const { data, error } = await supabase
      .from("app_data")
      .select("value")
      .eq("key", key)
      .maybeSingle();
    if (error) throw error;
    if (!data) throw new Error(`key not found: ${key}`);
    return { key, value: JSON.stringify(data.value) };
  },

  async set(key, value) {
    const { error } = await supabase
      .from("app_data")
      .upsert({ key, value: JSON.parse(value) }, { onConflict: "key" });
    if (error) throw error;
    return { key, value };
  },

  async delete(key) {
    const { error } = await supabase.from("app_data").delete().eq("key", key);
    if (error) throw error;
    return { key, deleted: true };
  },

  async list(prefix) {
    let query = supabase.from("app_data").select("key");
    if (prefix) query = query.like("key", `${prefix}%`);
    const { data, error } = await query;
    if (error) throw error;
    return { keys: (data || []).map((r) => r.key) };
  },
};
