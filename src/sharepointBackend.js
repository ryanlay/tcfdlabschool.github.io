const STORAGE_KEY = 'lab-school-db-state'

export function isSharePointConfigured() {
  return true
}

function cloneDefaultState(defaultState) {
  return {
    subjects: Array.isArray(defaultState.subjects) ? defaultState.subjects : [],
    behaviors: Array.isArray(defaultState.behaviors) ? defaultState.behaviors : [],
    videos: Array.isArray(defaultState.videos) ? defaultState.videos : [],
  }
}

export async function loadSharedState(defaultState) {
  const fallback = cloneDefaultState(defaultState)

  try {
    const stored = localStorage.getItem(STORAGE_KEY)
    if (!stored) return fallback

    const parsed = JSON.parse(stored)
    return {
      subjects: Array.isArray(parsed.subjects) ? parsed.subjects : fallback.subjects,
      behaviors: Array.isArray(parsed.behaviors) ? parsed.behaviors : fallback.behaviors,
      videos: Array.isArray(parsed.videos) ? parsed.videos : fallback.videos,
    }
  } catch {
    return fallback
  }
}

export async function saveSharedState(state) {
  const payload = {
    subjects: Array.isArray(state.subjects) ? state.subjects : [],
    behaviors: Array.isArray(state.behaviors) ? state.behaviors : [],
    videos: Array.isArray(state.videos) ? state.videos : [],
  }

  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload))
}
