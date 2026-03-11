import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = "https://uozeeclvruonpdaznhod.supabase.co"
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVvemVlY2x2cnVvbnBkYXpuaG9kIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzI4MjgwNDMsImV4cCI6MjA4ODQwNDA0M30.fDNk57tSvVDB7vY--LE4SaQGLAZtCZ0l89-we9vuv8Y"

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY)
