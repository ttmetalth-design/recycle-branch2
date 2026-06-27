import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://fzutuwlsbayatkcmjbmo.supabase.co'
const supabaseKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ6dXR1d2xzYmF5YXRrY21qYm1vIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI1MTc1MjQsImV4cCI6MjA5ODA5MzUyNH0.O3dcSUeAPKv87GaDhKhz3SwCLeIHThriWIo5SuzoEqk'

export const supabase = supabaseKey ? createClient(supabaseUrl, supabaseKey) : null
export const isSupabaseReady = !!supabaseKey
