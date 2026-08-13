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

// Thrown by saveSharedState when another device/tab has saved newer changes than the ones
// this client last loaded. Callers must NOT retry-overwrite on this error - the caller should
// tell the user to reload and redo their change, otherwise their save would silently clobber
// someone else's more recent edit (multiple tabs/devices share one row with no realtime sync).
export class StaleWriteError extends Error {
  constructor(message) {
    super(message)
    this.name = 'StaleWriteError'
  }
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
  if (!isSharePointConfigured()) return { ...fallback, updatedAt: null }

  const { data, error } = await supabase
    .from(SUPABASE_TABLE)
    .select('subjects, behaviors, videos, updated_at')
    .eq('state_key', SUPABASE_STATE_KEY)
    .maybeSingle()

  if (error) {
    throw new Error(`Supabase load failed: ${error.message}`)
  }

  if (!data) {
    const saved = await saveSharedState(fallback, { force: true })
    return { ...fallback, updatedAt: saved.updatedAt }
  }

  return { ...normalizeState(data, fallback), updatedAt: data.updated_at ?? null }
}

// `expectedUpdatedAt` should be the `updatedAt` this client last saw (from loadSharedState or a
// prior saveSharedState result). If the row's current updated_at no longer matches, someone else
// has saved since - we refuse to overwrite it and throw StaleWriteError instead. Pass
// `{ force: true }` only for the one-time "create the row if it doesn't exist yet" bootstrap.
export async function saveSharedState(state, { expectedUpdatedAt, force = false } = {}) {
  if (!isSharePointConfigured()) return { updatedAt: null }

  if (!force) {
    const { data: current, error: checkError } = await supabase
      .from(SUPABASE_TABLE)
      .select('updated_at')
      .eq('state_key', SUPABASE_STATE_KEY)
      .maybeSingle()

    if (checkError) {
      throw new Error(`Supabase save failed: ${checkError.message}`)
    }

    const currentUpdatedAt = current?.updated_at ?? null
    if (currentUpdatedAt !== (expectedUpdatedAt ?? null)) {
      throw new StaleWriteError(
        'Someone else saved changes to the shared data since this device last loaded it. Reload to get the latest version, then redo your change.',
      )
    }
  }

  const payload = {
    state_key: SUPABASE_STATE_KEY,
    subjects: Array.isArray(state?.subjects) ? state.subjects : [],
    behaviors: Array.isArray(state?.behaviors) ? state.behaviors : [],
    videos: Array.isArray(state?.videos) ? state.videos : [],
    updated_at: new Date().toISOString(),
  }

  // Read back the value Postgres actually stored (rather than trusting the ISO string we sent)
  // so future comparisons in the pre-save check above use the exact same serialization that a
  // plain `.select('updated_at')` would return - otherwise formatting differences (e.g. Postgres
  // returning microsecond precision / a "+00:00" offset instead of our "Z"-suffixed string) would
  // make every subsequent save from this same tab look like a false-positive conflict.
  const { data: savedRow, error } = await supabase
    .from(SUPABASE_TABLE)
    .upsert(payload, { onConflict: 'state_key' })
    .select('updated_at')
    .single()

  if (error) {
    throw new Error(`Supabase save failed: ${error.message}`)
  }

  return { updatedAt: savedRow.updated_at }
}
