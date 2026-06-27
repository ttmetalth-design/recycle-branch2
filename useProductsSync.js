import { useEffect, useRef, useCallback } from 'react'
import { supabase, isSupabaseReady } from './supabase'

// แปลงรูปแบบจาก Supabase (snake_case) -> รูปแบบที่แอปใช้ (camelCase)
function fromRow(row) {
  return {
    id: row.id,
    name: row.name,
    type: row.type,
    unit: row.unit,
    openingQty: row.opening_qty,
    openingCost: row.opening_cost,
    openingMonth: row.opening_month || "",
    buyPrice: row.buy_price || 0,
    vipPrice: row.vip_price || 0,
  }
}

// แปลงรูปแบบจากแอป (camelCase) -> Supabase (snake_case)
function toRow(product) {
  return {
    id: product.id,
    name: product.name,
    type: product.type,
    unit: product.unit,
    opening_qty: product.openingQty,
    opening_cost: product.openingCost,
    opening_month: product.openingMonth || "",
    buy_price: Number(product.buyPrice) || 0,
    vip_price: Number(product.vipPrice) || 0,
    updated_at: new Date().toISOString(),
  }
}

// โหลดสินค้าทั้งหมดจากตาราง products (เรียกครั้งเดียวตอนเปิดแอป)
export async function loadProducts() {
  if (!isSupabaseReady) return []
  const { data, error } = await supabase.from('products').select('*').order('id')
  if (error || !data) return []
  return data.map(fromRow)
}

// เพิ่มสินค้าใหม่ 1 รายการ — เขียนตรงไป Supabase ทันที ไม่ต้องรอ debounce
export async function insertProduct(product) {
  if (!isSupabaseReady) return { error: 'not ready' }
  const { error } = await supabase.from('products').insert(toRow(product))
  return { error }
}

// แก้ไขสินค้า 1 รายการ
export async function updateProduct(product) {
  if (!isSupabaseReady) return { error: 'not ready' }
  const { error } = await supabase.from('products').update(toRow(product)).eq('id', product.id)
  return { error }
}

// ลบสินค้า 1 รายการ
export async function deleteProduct(id) {
  if (!isSupabaseReady) return { error: 'not ready' }
  const { error } = await supabase.from('products').delete().eq('id', id)
  return { error }
}

// Hook: subscribe การเปลี่ยนแปลงของตาราง products แบบ realtime
// เมื่อเครื่องอื่นเพิ่ม/แก้/ลบสินค้า เครื่องนี้จะเห็นการเปลี่ยนแปลงทันทีโดยไม่ต้องรอ polling
export function useProductsRealtime(setProducts, loaded) {
  const channelRef = useRef(null)

  useEffect(() => {
    if (!loaded || !isSupabaseReady) return

    const channel = supabase
      .channel('products-realtime')
      .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'products' }, (payload) => {
        setProducts((prev) => {
          if (prev.some((p) => p.id === payload.new.id)) return prev
          return [...prev, fromRow(payload.new)]
        })
      })
      .on('postgres_changes', { event: 'UPDATE', schema: 'public', table: 'products' }, (payload) => {
        setProducts((prev) => prev.map((p) => (p.id === payload.new.id ? fromRow(payload.new) : p)))
      })
      .on('postgres_changes', { event: 'DELETE', schema: 'public', table: 'products' }, (payload) => {
        setProducts((prev) => prev.filter((p) => p.id !== payload.old.id))
      })
      .subscribe()

    channelRef.current = channel

    return () => {
      supabase.removeChannel(channel)
    }
  }, [loaded, setProducts])
}
