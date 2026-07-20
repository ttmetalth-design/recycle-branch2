import { useEffect, useRef, useState } from 'react'
import { supabase, isSupabaseReady } from './supabase'

const DEVICE_ID = (() => {
  try {
    const stored = localStorage.getItem('_device_id')
    if (stored) return stored
    const newId = Math.random().toString(36).slice(2)
    localStorage.setItem('_device_id', newId)
    return newId
  } catch {
    return Math.random().toString(36).slice(2)
  }
})()

const ARRAY_TABLES = {
  purchases: 'purchases',
  sales: 'sales',
  customers: 'customers',
  expenses: 'expenses',
  withdrawals: 'withdrawals',
  deposits: 'deposits',
  prepayments: 'prepayments',
  deliveries: 'deliveries',
  bankTransfers: 'bank_transfers',
  assets: 'assets',
  loans: 'loans',
  dividendPayments: 'dividend_payments',
  storeBankAccounts: 'store_bank_accounts',
  shareholders: 'shareholders',
}

// ตาราง → stateKey (reverse map)
const TABLE_TO_KEY = Object.fromEntries(
  Object.entries(ARRAY_TABLES).map(([k, v]) => [v, k])
)

const SETTINGS_KEYS = [
  'shopProfile', 'companySettings', 'unitOptions',
  'expenseCategories', 'productCategories', 'payFlags',
]

// ---------- Global sync status ----------
let globalStatus = 'synced'
let pendingCount = 0
const statusListeners = new Set()
const pendingSaveKeys = new Set() // keys ที่กำลัง pending save อยู่

function setGlobalStatus(status) {
  globalStatus = status
  statusListeners.forEach(fn => fn(status))
}
function incrementPending() { pendingCount++; setGlobalStatus('saving') }
function decrementPending(success) {
  pendingCount = Math.max(0, pendingCount - 1)
  if (pendingCount === 0) setGlobalStatus(success ? 'synced' : 'error')
}

export function useSyncStatus() {
  const [status, setStatus] = useState(globalStatus)
  useEffect(() => {
    statusListeners.add(setStatus)
    return () => statusListeners.delete(setStatus)
  }, [])
  return status
}

// ---------- Array table helpers ----------
async function saveArrayTable(tableName, items) {
  if (!isSupabaseReady || !Array.isArray(items) || items.length === 0) return true
  const rows = items.map(item => ({
    id: item.id,
    data: { ...item, _updated_by: DEVICE_ID },
    updated_at: new Date().toISOString(),
  }))
  const { error } = await supabase.from(tableName).upsert(rows, { onConflict: 'id' })
  return !error
}

async function deleteArrayRow(tableName, id) {
  if (!isSupabaseReady) return true
  const { error } = await supabase.from(tableName).delete().eq('id', id)
  return !error
}

async function loadArrayTable(tableName) {
  if (!isSupabaseReady) return []
  const PAGE = 1000
  let all = []
  let lastUpdatedAt = null
  let lastId = null
  while (true) {
    let query = supabase
      .from(tableName)
      .select('id, data, updated_at')
      .order('updated_at', { ascending: true })
      .order('id', { ascending: true })
      .limit(PAGE)

    // ใช้ cursor-based pagination แทน offset
    if (lastUpdatedAt && lastId) {
      query = query.or(`updated_at.gt.${lastUpdatedAt},and(updated_at.eq.${lastUpdatedAt},id.gt.${lastId})`)
    }

    const { data, error } = await query
    if (error || !data || data.length === 0) break
    all = all.concat(data.map(row => row.data))
    if (data.length < PAGE) break
    const last = data[data.length - 1]
    lastUpdatedAt = last.updated_at
    lastId = last.id
  }
  return all
}

// ---------- Settings helpers ----------
async function saveSettings(key, value) {
  if (!isSupabaseReady) return true
  const { error } = await supabase.from('app_settings').upsert(
    { key, data: { value, _updated_by: DEVICE_ID }, updated_at: new Date().toISOString() },
    { onConflict: 'key' }
  )
  return !error
}

async function loadSettings(key) {
  if (!isSupabaseReady) return null
  const { data, error } = await supabase
    .from('app_settings')
    .select('data')
    .eq('key', key)
    .single()
  if (error || !data) return null
  return data.data?.value ?? null
}

// ---------- loadAllFromSupabase ----------
export async function loadAllFromSupabase() {
  if (!isSupabaseReady) return null
  const result = {}
  await Promise.all(
    Object.entries(ARRAY_TABLES).map(async ([stateKey, tableName]) => {
      result[stateKey] = await loadArrayTable(tableName)
    })
  )
  await Promise.all(
    SETTINGS_KEYS.map(async (key) => {
      const val = await loadSettings(key)
      if (val !== null) result[key] = val
    })
  )
  return result
}

// ---------- saveToSupabase (เรียกตรงๆ สำหรับกรณีพิเศษ) ----------
export async function saveToSupabase(key, items) {
  const tableName = ARRAY_TABLES[key]
  if (tableName) return await saveArrayTable(tableName, items)
  if (SETTINGS_KEYS.includes(key)) return await saveSettings(key, items)
}

// ============================================================
// Global Realtime Manager
// รวม subscriptions ทั้งหมดเป็น 2 channels เท่านั้น
// (แทนที่จะเป็น 19 channels แยกกัน)
// ============================================================
const arraySetters = new Map()   // tableName → setValue fn
const settingsSetters = new Map() // key → setValue fn
let arrayChannel = null
let settingsChannel = null
let arrayChannelLoaded = false
let settingsChannelLoaded = false

function ensureArrayChannel() {
  if (arrayChannel) return
  arrayChannel = supabase
    .channel(`rt-all-arrays-${DEVICE_ID}`)
    .on('postgres_changes', { event: 'INSERT', schema: 'public', table: '*' }, (payload) => {
      const tableName = payload.table
      const stateKey = TABLE_TO_KEY[tableName]
      const setter = stateKey && arraySetters.get(stateKey)
      if (!setter) return
      const item = payload.new?.data
      if (!item || item._updated_by === DEVICE_ID) return
      setter(prev => prev.some(x => x.id === item.id) ? prev : [...prev, item])
    })
    .on('postgres_changes', { event: 'UPDATE', schema: 'public', table: '*' }, (payload) => {
      const tableName = payload.table
      const stateKey = TABLE_TO_KEY[tableName]
      const setter = stateKey && arraySetters.get(stateKey)
      if (!setter) return
      const item = payload.new?.data
      if (!item || item._updated_by === DEVICE_ID) return
      setter(prev => prev.map(x => x.id === item.id ? item : x))
    })
    .on('postgres_changes', { event: 'DELETE', schema: 'public', table: '*' }, (payload) => {
      const tableName = payload.table
      const stateKey = TABLE_TO_KEY[tableName]
      const setter = stateKey && arraySetters.get(stateKey)
      if (!setter) return
      const id = payload.old?.id
      if (!id) return
      setter(prev => prev.filter(x => x.id !== id))
    })
    .subscribe()
}

function ensureSettingsChannel() {
  if (settingsChannel) return
  settingsChannel = supabase
    .channel(`rt-all-settings-${DEVICE_ID}`)
    .on('postgres_changes', { event: 'UPDATE', schema: 'public', table: 'app_settings' }, (payload) => {
      const key = payload.new?.key
      const setter = key && settingsSetters.get(key)
      if (!setter) return
      const updatedBy = payload.new?.data?._updated_by
      // ถ้ามาจาก device เดียวกัน ข้าม (เราอัปเดต state เองแล้ว)
      if (updatedBy === DEVICE_ID) return
      const newValue = payload.new?.data?.value
      if (newValue === undefined) return
      // สำหรับ payFlags: merge แทน overwrite เพื่อไม่ให้ข้อมูลหาย
      if (key === 'payFlags' && typeof newValue === 'object' && newValue !== null) {
        setter(prev => {
          if (typeof prev !== 'object' || prev === null) return newValue
          // merge: ค่า true จากทั้งสองฝั่ง จะ win (ติ๊กแล้วไม่หาย)
          return { ...newValue, ...prev }
        })
        return
      }
      setter(newValue)
    })
    .subscribe()
}

// ---------- useSupabaseSync ----------
export function useSupabaseSync(key, value, setValue, loaded) {
  const valueRef = useRef(value)
  useEffect(() => { valueRef.current = value }, [value])

  const prevValueRef = useRef(null)
  const saveTimer = useRef(null)
  const maxWaitTimer = useRef(null)
  const isFirstRender = useRef(true)

  const tableName = ARRAY_TABLES[key]
  const isArrayTable = !!tableName
  const isSettingsKey = SETTINGS_KEYS.includes(key)

  // ---------- SAVE ----------
  useEffect(() => {
    if (!loaded || !isSupabaseReady) return
    if (isFirstRender.current) {
      isFirstRender.current = false
      prevValueRef.current = value
      return
    }

    const doSave = async () => {
      clearTimeout(saveTimer.current)
      clearTimeout(maxWaitTimer.current)
      saveTimer.current = null
      maxWaitTimer.current = null

      if (isSettingsKey) pendingSaveKeys.add(key)
      incrementPending()
      let success = false
      try {
        if (isArrayTable) {
          const current = valueRef.current
          const prev = prevValueRef.current || []
          const prevMap = new Map(prev.filter(x => x.id).map(x => [x.id, JSON.stringify(x)]))
          const changed = current.filter(item => {
            if (!item.id) return true
            return prevMap.get(item.id) !== JSON.stringify(item)
          })
          const currentIds = new Set(current.filter(x => x.id).map(x => x.id))
          const deleted = prev.filter(x => x.id && !currentIds.has(x.id))

          let ok = true
          if (changed.length > 0) ok = await saveArrayTable(tableName, changed)
          for (const item of deleted) {
            const r = await deleteArrayRow(tableName, item.id)
            if (!r) ok = false
          }
          success = ok
          prevValueRef.current = [...current]
        } else if (isSettingsKey) {
          success = await saveSettings(key, valueRef.current)
          prevValueRef.current = valueRef.current
        }
      } catch {
        success = false
      } finally {
        if (isSettingsKey) {
          // ปลด lock หลัง save เสร็จ (500ms พอ — แค่กัน echo ทันที)
          setTimeout(() => pendingSaveKeys.delete(key), 500)
        }
        decrementPending(success)
      }
    }

    clearTimeout(saveTimer.current)
    saveTimer.current = setTimeout(doSave, 2000)
    if (!maxWaitTimer.current) {
      maxWaitTimer.current = setTimeout(doSave, 6000)
    }

    return () => {
      clearTimeout(saveTimer.current)
      clearTimeout(maxWaitTimer.current)
    }
  }, [key, value, loaded])

  // ---------- REALTIME (shared channels) ----------
  useEffect(() => {
    if (!isSupabaseReady || !loaded) return

    if (isArrayTable) {
      // ลงทะเบียน setter ใน global map
      arraySetters.set(key, setValue)
      // สร้าง channel รวม (ถ้ายังไม่มี)
      ensureArrayChannel()
      return () => {
        arraySetters.delete(key)
        // ถ้าไม่มี setter เหลือแล้ว ปิด channel
        if (arraySetters.size === 0 && arrayChannel) {
          supabase.removeChannel(arrayChannel)
          arrayChannel = null
        }
      }
    }

    if (isSettingsKey) {
      settingsSetters.set(key, setValue)
      ensureSettingsChannel()
      return () => {
        settingsSetters.delete(key)
        if (settingsSetters.size === 0 && settingsChannel) {
          supabase.removeChannel(settingsChannel)
          settingsChannel = null
        }
      }
    }
  }, [key, setValue, loaded, tableName, isArrayTable, isSettingsKey])
}
