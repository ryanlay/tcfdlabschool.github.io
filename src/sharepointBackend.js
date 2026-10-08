import { initializeApp } from 'firebase/app'
import { doc, getDoc, getFirestore, runTransaction } from 'firebase/firestore'

const firebaseConfig = {
  apiKey: import.meta.env.VITE_FIREBASE_API_KEY,
  authDomain: import.meta.env.VITE_FIREBASE_AUTH_DOMAIN,
  projectId: import.meta.env.VITE_FIREBASE_PROJECT_ID,
  storageBucket: import.meta.env.VITE_FIREBASE_STORAGE_BUCKET,
  messagingSenderId: import.meta.env.VITE_FIREBASE_MESSAGING_SENDER_ID,
  appId: import.meta.env.VITE_FIREBASE_APP_ID,
}
const FIREBASE_COLLECTION = import.meta.env.VITE_FIREBASE_COLLECTION || 'lab_school_state'
const FIREBASE_STATE_ID = 'shared'
const REQUEST_TIMEOUT_MS = 10000
const isConfigured = Object.values(firebaseConfig).every((value) => typeof value === 'string' && value.trim())
const db = isConfigured ? getFirestore(initializeApp(firebaseConfig)) : null
const sharedStateRef = db ? doc(db, FIREBASE_COLLECTION, FIREBASE_STATE_ID) : null

// Firestore's web SDK does not expose an AbortSignal for individual operations. Bound the
// wait at the call site so a blackholed connection fails promptly instead of stalling retries.
function withTimeout(promise, timeoutMs = REQUEST_TIMEOUT_MS) {
  let timeoutId
  const timeout = new Promise((_, reject) => {
    timeoutId = setTimeout(() => reject(new Error('Firebase request timed out. Check your network/firewall and retry.')), timeoutMs)
  })

  return Promise.race([promise, timeout]).finally(() => clearTimeout(timeoutId))
}

export function isSharePointConfigured() {
  return Boolean(db && sharedStateRef)
}

// Thrown when another device/tab has saved newer changes than the ones this client last loaded.
// Callers must not retry-overwrite on this error; they should ask the user to reload and redo
// their change, otherwise their save could silently clobber someone else's edit.
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
    devices: Array.isArray(defaultState.devices) ? defaultState.devices : [],
    videos: Array.isArray(defaultState.videos) ? defaultState.videos : [],
  }
}

function normalizeState(state, fallback) {
  return {
    subjects: Array.isArray(state?.subjects) ? state.subjects : fallback.subjects,
    behaviors: Array.isArray(state?.behaviors) ? state.behaviors : fallback.behaviors,
    devices: Array.isArray(state?.devices) ? state.devices : fallback.devices,
    videos: Array.isArray(state?.videos) ? state.videos : fallback.videos,
  }
}

export async function loadSharedState(defaultState) {
  const fallback = cloneDefaultState(defaultState)
  if (!isSharePointConfigured()) return { ...fallback, updatedAt: null }

  let snapshot
  try {
    snapshot = await withTimeout(getDoc(sharedStateRef))
  } catch (error) {
    throw new Error(`Firebase load failed: ${error instanceof Error ? error.message : 'Unknown error'}`)
  }

  if (!snapshot.exists()) {
    await saveSharedState(fallback, { force: true })
    try {
      snapshot = await withTimeout(getDoc(sharedStateRef))
    } catch (error) {
      throw new Error(`Firebase load failed: ${error instanceof Error ? error.message : 'Unknown error'}`)
    }
    if (!snapshot.exists()) return { ...fallback, updatedAt: null }
  }

  const data = snapshot.data()
  return { ...normalizeState(data, fallback), updatedAt: data.updatedAt ?? null }
}

// `expectedUpdatedAt` should be the `updatedAt` this client last saw. The transaction compares
// and writes atomically, so concurrent devices cannot both pass the stale-write check.
// `{ force: true }` is reserved for bootstrapping the shared document when it does not exist.
export async function saveSharedState(state, { expectedUpdatedAt, force = false } = {}) {
  if (!isSharePointConfigured()) return { updatedAt: null }

  const payload = {
    subjects: Array.isArray(state?.subjects) ? state.subjects : [],
    behaviors: Array.isArray(state?.behaviors) ? state.behaviors : [],
    devices: Array.isArray(state?.devices) ? state.devices : [],
    videos: Array.isArray(state?.videos) ? state.videos : [],
  }

  try {
    return await withTimeout(runTransaction(db, async (transaction) => {
      const current = await transaction.get(sharedStateRef)
      const currentUpdatedAt = current.exists() ? current.data().updatedAt ?? null : null
      if (force && current.exists()) return { updatedAt: currentUpdatedAt }
      if (!force && currentUpdatedAt !== (expectedUpdatedAt ?? null)) {
        throw new StaleWriteError(
          'Someone else saved changes to the shared data since this device last loaded it. Reload to get the latest version, then redo your change.',
        )
      }

      const previousMillis = Date.parse(currentUpdatedAt)
      const updatedAt = new Date(Math.max(Date.now(), Number.isFinite(previousMillis) ? previousMillis + 1 : 0)).toISOString()
      transaction.set(sharedStateRef, { ...payload, updatedAt })
      return { updatedAt }
    }))
  } catch (error) {
    if (error instanceof StaleWriteError) throw error
    throw new Error(`Firebase save failed: ${error instanceof Error ? error.message : 'Unknown error'}`)
  }
}
