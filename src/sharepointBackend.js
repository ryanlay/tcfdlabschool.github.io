import { createClient } from '@supabase/supabase-js'

const rawSupabaseUrl = import.meta.env.VITE_SUPABASE_URL || 'https://uqedhpsjugpnlzohearq.supabase.co/rest/v1/'
const SUPABASE_URL = rawSupabaseUrl.replace(/\/rest\/v1\/?$/, '')
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || 'sb_publishable_OvVPFSDDRNxQGnw8lIyXOA_KWB0G6GP'
const SUPABASE_TABLE = import.meta.env.VITE_SUPABASE_TABLE || 'lab_school_state'
const SUPABASE_STATE_KEY = 'shared'

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
  auth: {
    persistSession: false,
    autoRefreshToken: false,
  },
})

export function isSharePointConfigured() {
  return Boolean(SUPABASE_URL && SUPABASE_ANON_KEY)
}

function cloneDefaultState(defaultState) {
  return {
    subjects: Array.isArray(defaultState.subjects) ? defaultState.subjects : [],
    behaviors: Array.isArray(defaultState.behaviors) ? defaultState.behaviors : [],
    videos: Array.isArray(defaultState.videos) ? defaultState.videos : [],
  }
}

function normalizeState(state, fallback) {
  return {
    subjects: Array.isArray(state?.subjects) ? state.subjects : fallback.subjects,
    behaviors: Array.isArray(state?.behaviors) ? state.behaviors : fallback.behaviors,
    videos: Array.isArray(state?.videos) ? state.videos : fallback.videos,
  }
}

export async function loadSharedState(defaultState) {
  const fallback = cloneDefaultState(defaultState)
  if (!isSharePointConfigured()) return fallback

  const { data, error } = await supabase
    .from(SUPABASE_TABLE)
    .select('subjects, behaviors, videos')
    .eq('state_key', SUPABASE_STATE_KEY)
    .maybeSingle()

  if (error) {
    throw new Error(`Supabase load failed: ${error.message}`)
  }

  if (!data) {
    await saveSharedState(fallback)
    return fallback
  }

  return normalizeState(data, fallback)
}

export async function saveSharedState(state) {
  if (!isSharePointConfigured()) return

  const payload = {
    state_key: SUPABASE_STATE_KEY,
    subjects: Array.isArray(state?.subjects) ? state.subjects : [],
    behaviors: Array.isArray(state?.behaviors) ? state.behaviors : [],
    videos: Array.isArray(state?.videos) ? state.videos : [],
    updated_at: new Date().toISOString(),
  }

  const { error } = await supabase
    .from(SUPABASE_TABLE)
    .upsert(payload, { onConflict: 'state_key' })

  if (error) {
    throw new Error(`Supabase save failed: ${error.message}`)
  }
}
