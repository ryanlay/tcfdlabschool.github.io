/**
 * SharePoint backend using REST API with browser authentication.
 * Primary: SharePoint (cross-device/cross-browser sync)
 * Fallback: localStorage (offline only)
 */

const SHAREPOINT_HOSTNAME = import.meta.env.VITE_SHAREPOINT_HOSTNAME
const SHAREPOINT_SITE_PATH = import.meta.env.VITE_SHAREPOINT_SITE_PATH
const SHAREPOINT_LIST_NAME = import.meta.env.VITE_SHAREPOINT_LIST_NAME || 'LabSchoolAppState'

const STORAGE_KEY = 'lab-school-db-state'

// Cache for list/item IDs (per session)
let listIdCache = null
let itemIdCache = null
let formDigestCache = null
let formDigestExpiry = 0

export function isSharePointConfigured() {
  return Boolean(
    SHAREPOINT_HOSTNAME
      && SHAREPOINT_SITE_PATH
      && SHAREPOINT_LIST_NAME,
  )
}

/**
 * Build the base SharePoint URL
 */
function getBaseUrl() {
  return `https://${SHAREPOINT_HOSTNAME}/${SHAREPOINT_SITE_PATH.replace(/^\/+/, '')}`
}

/**
 * Direct SharePoint REST API call with browser authentication.
 */
async function sharepointRequest(relativePath, init = {}) {
  const url = `${getBaseUrl()}${relativePath.startsWith('/') ? '' : '/'}${relativePath}`
  const options = {
    ...init,
    headers: {
      'Content-Type': 'application/json',
      Accept: 'application/json',
      ...init.headers,
    },
    credentials: 'include', // Include cookies for browser auth
  }

  console.log(`[SP] ${(init.method || 'GET').padEnd(6)} ${url}`)

  const response = await fetch(url, options)

  if (!response.ok) {
    const text = await response.text().catch(() => '')
    const errorMsg = `${response.status}: ${text || response.statusText}`.substring(0, 100)
    console.error(`[SP] ✗ ${errorMsg}`)
    throw new Error(`SharePoint API error ${response.status}`)
  }

  if (response.status === 204) return null
  const json = await response.json()
  console.log(`[SP] ✓ Success`)
  return json
}

/**
 * Get or refresh form digest (CSRF token for POST/PATCH)
 * Form digest expires after ~30 minutes
 */
async function getFormDigest() {
  const now = Date.now()
  
  // Return cached digest if still valid
  if (formDigestCache && formDigestExpiry > now) {
    console.log(`[SP] Using cached form digest`)
    return formDigestCache
  }

  console.log(`[SP] Fetching new form digest...`)
  const response = await sharepointRequest('/_api/contextinfo', {
    method: 'POST',
  })
  
  formDigestCache = response.FormDigestValue
  formDigestExpiry = now + (28 * 60 * 1000) // Cache for 28 minutes
  
  return formDigestCache
}

/**
 * Find or create the list
 */
async function getListId() {
  if (listIdCache) {
    console.log(`[SP] Using cached list ID: ${listIdCache}`)
    return listIdCache
  }
  
  console.log(`[SP] Searching for list: ${SHAREPOINT_LIST_NAME}`)
  
  // Try to find existing list
  try {
    const data = await sharepointRequest(
      `/_api/web/lists?$filter=Title eq '${SHAREPOINT_LIST_NAME.replace(/'/g, "''")}'&$select=Id`
    )

    if (Array.isArray(data?.value) && data.value.length > 0) {
      listIdCache = data.value[0].Id
      console.log(`[SP] Found existing list: ${listIdCache}`)
      return listIdCache
    }
  } catch (err) {
    console.error(`[SP] Error searching for list:`, err.message)
    throw err
  }

  // Create new list if not found
  console.log(`[SP] Creating new list: ${SHAREPOINT_LIST_NAME}`)
  const digest = await getFormDigest()
  
  const created = await sharepointRequest('/_api/web/lists', {
    method: 'POST',
    headers: { 'X-RequestDigest': digest },
    body: JSON.stringify({
      __metadata: { type: 'SP.List' },
      Title: SHAREPOINT_LIST_NAME,
      BaseTemplate: 100, // Generic list
      Description: 'Shared state for Lab School Database',
    }),
  })
  
  listIdCache = created.Id
  console.log(`[SP] ✓ Created list: ${listIdCache}`)

  // Add columns
  const columns = [
    { Title: 'subjectsJson', FieldTypeKind: 3 },
    { Title: 'behaviorsJson', FieldTypeKind: 3 },
    { Title: 'videosJson', FieldTypeKind: 3 },
  ]
  
  for (const col of columns) {
    try {
      const colDigest = await getFormDigest()
      await sharepointRequest(
        `/_api/web/lists(guid'${listIdCache}')/fields`,
        {
          method: 'POST',
          headers: { 'X-RequestDigest': colDigest },
          body: JSON.stringify({
            __metadata: { type: 'SP.Field' },
            FieldTypeKind: col.FieldTypeKind,
            Title: col.Title,
          }),
        }
      )
      console.log(`[SP] ✓ Added column: ${col.Title}`)
    } catch (err) {
      console.warn(`[SP] Column may already exist: ${col.Title}`)
    }
  }

  return listIdCache
}

/**
 * Find or create the SharedState item
 */
async function getSharedStateItem(listId) {
  if (itemIdCache) {
    console.log(`[SP] Using cached item ID: ${itemIdCache}`)
    return itemIdCache
  }

  console.log(`[SP] Searching for SharedState item...`)
  
  // Find existing item
  try {
    const existing = await sharepointRequest(
      `/_api/web/lists(guid'${listId}')/items?$filter=Title eq 'SharedState'&$select=Id`
    )
    
    if (Array.isArray(existing?.value) && existing.value.length > 0) {
      itemIdCache = existing.value[0].Id
      console.log(`[SP] Found existing item: ${itemIdCache}`)
      return itemIdCache
    }
  } catch (err) {
    console.error(`[SP] Error searching for item:`, err.message)
    throw err
  }

  // Create new item
  console.log(`[SP] Creating new SharedState item...`)
  const digest = await getFormDigest()
  
  const created = await sharepointRequest(
    `/_api/web/lists(guid'${listId}')/items`,
    {
      method: 'POST',
      headers: { 'X-RequestDigest': digest },
      body: JSON.stringify({
        __metadata: { type: 'SP.ListItem' },
        Title: 'SharedState',
        subjectsJson: '[]',
        behaviorsJson: '[]',
        videosJson: '[]',
      }),
    }
  )

  itemIdCache = created.Id
  console.log(`[SP] ✓ Created item: ${itemIdCache}`)
  return itemIdCache
}

function parseJsonArray(value, fallback) {
  try {
    const parsed = JSON.parse(value)
    return Array.isArray(parsed) ? parsed : fallback
  } catch {
    return fallback
  }
}

export async function loadSharedState(defaultState) {
  const fallback = {
    subjects: defaultState.subjects,
    behaviors: defaultState.behaviors,
    videos: defaultState.videos,
  }

  if (!isSharePointConfigured()) {
    console.log('[Local] SharePoint not configured, using localStorage only')
    try {
      const stored = localStorage.getItem(STORAGE_KEY)
      if (stored) {
        return JSON.parse(stored)
      }
    } catch (err) {
      console.warn('[Local] Error loading from localStorage:', err)
    }
    return fallback
  }

  // Try SharePoint FIRST (not as fallback)
  try {
    console.log('[SP] ===== LOAD FROM SHAREPOINT =====')
    const listId = await getListId()
    const itemId = await getSharedStateItem(listId)

    const item = await sharepointRequest(
      `/_api/web/lists(guid'${listId}')/items(${itemId})?$select=subjectsJson,behaviorsJson,videosJson`
    )

    const result = {
      subjects: parseJsonArray(item.subjectsJson, fallback.subjects),
      behaviors: parseJsonArray(item.behaviorsJson, fallback.behaviors),
      videos: parseJsonArray(item.videosJson, fallback.videos),
    }
    console.log('[SP] ✓✓✓ LOADED FROM SHAREPOINT ✓✓✓', result)
    return result
  } catch (err) {
    console.error('[SP] ✗✗✗ FAILED TO LOAD FROM SHAREPOINT ✗✗✗', err.message)
    console.log('[Local] Falling back to localStorage')
    
    try {
      const stored = localStorage.getItem(STORAGE_KEY)
      if (stored) {
        const parsed = JSON.parse(stored)
        console.log('[Local] Loaded from localStorage:', parsed)
        return parsed
      }
    } catch (innerErr) {
      console.warn('[Local] Error loading from localStorage:', innerErr)
    }
  }

  return fallback
}

export async function saveSharedState(state) {
  // ALWAYS try SharePoint FIRST
  if (isSharePointConfigured()) {
    try {
      console.log('[SP] ===== SAVE TO SHAREPOINT =====')
      
      const listId = await getListId()
      const itemId = await getSharedStateItem(listId)
      const digest = await getFormDigest()

      await sharepointRequest(
        `/_api/web/lists(guid'${listId}')/items(${itemId})`,
        {
          method: 'PATCH',
          headers: {
            'X-RequestDigest': digest,
            'If-Match': '*',
          },
          body: JSON.stringify({
            __metadata: { type: 'SP.ListItem' },
            subjectsJson: JSON.stringify(state.subjects || []),
            behaviorsJson: JSON.stringify(state.behaviors || []),
            videosJson: JSON.stringify(state.videos || []),
          }),
        }
      )
      console.log('[SP] ✓✓✓ SAVED TO SHAREPOINT ✓✓✓')
      return // Success - don't need fallback
    } catch (err) {
      console.error('[SP] ✗✗✗ FAILED TO SAVE TO SHAREPOINT ✗✗✗', err.message)
    }
  }

  // Fallback to localStorage only if SharePoint failed
  console.log('[Local] Saving to localStorage as fallback...')
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state))
    console.log('[Local] ✓ Saved to localStorage')
  } catch (err) {
    console.error('[Local] ✗ Failed to save to localStorage:', err)
  }
}
