/**
 * SharePoint backend using localStorage fallback for local dev.
 * In production on GitHub Pages, user must have SharePoint site access.
 */

const SHAREPOINT_HOSTNAME = import.meta.env.VITE_SHAREPOINT_HOSTNAME
const SHAREPOINT_SITE_PATH = import.meta.env.VITE_SHAREPOINT_SITE_PATH
const SHAREPOINT_LIST_NAME = import.meta.env.VITE_SHAREPOINT_LIST_NAME || 'LabSchoolAppState'

const STORAGE_KEY = 'lab-school-db-state'
let hasTriedSharePoint = false

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
 * Browser will handle OAuth popup if needed on first request.
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

  console.log(`[SharePoint] ${init.method || 'GET'} ${url}`)

  const response = await fetch(url, options)

  if (!response.ok) {
    // 401/403 typically means auth issue
    if (response.status === 401 || response.status === 403) {
      throw new Error(
        `SharePoint authentication required. Status: ${response.status}. `
        + `You may need to sign in at ${getBaseUrl()} first.`
      )
    }
    const text = await response.text().catch(() => '')
    throw new Error(`SharePoint API error ${response.status}: ${text || response.statusText}`)
  }

  if (response.status === 204) return null
  return response.json()
}

/**
 * Get form digest for POST/PATCH operations (CSRF protection)
 */
async function getFormDigest() {
  const response = await sharepointRequest('/_api/contextinfo', {
    method: 'POST',
  })
  return response.FormDigestValue
}

/**
 * Find or create list
 */
async function getListId() {
  // Try to find existing list
  try {
    const data = await sharepointRequest(
      `/_api/web/lists?$filter=Title eq '${SHAREPOINT_LIST_NAME.replace(/'/g, "''")}'`
    )

    const existing = Array.isArray(data?.value) ? data.value[0] : null
    if (existing?.Id) {
      console.log(`[SharePoint] Found existing list: ${existing.Id}`)
      return existing.Id
    }
  } catch (err) {
    console.warn('[SharePoint] Error searching for existing list:', err.message)
    throw err
  }

  // Create new list if not found
  console.log(`[SharePoint] Creating new list: ${SHAREPOINT_LIST_NAME}`)
  const digest = await getFormDigest()
  
  const created = await sharepointRequest('/_api/web/lists', {
    method: 'POST',
    headers: { 'X-RequestDigest': digest },
    body: JSON.stringify({
      __metadata: { type: 'SP.List' },
      Title: SHAREPOINT_LIST_NAME,
      BaseTemplate: 100,
      Description: 'Shared state for Lab School Database',
    }),
  })
  
  console.log(`[SharePoint] Created new list: ${created.Id}`)

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
        `/_api/web/lists(guid'${created.Id}')/fields`,
        {
          method: 'POST',
          headers: { 'X-RequestDigest': colDigest },
          body: JSON.stringify({
            __metadata: { type: 'SP.Field' },
            ...col,
          }),
        }
      )
      console.log(`[SharePoint] Added column: ${col.Title}`)
    } catch (err) {
      console.warn(`[SharePoint] Could not add column ${col.Title}:`, err.message)
    }
  }

  return created.Id
}

/**
 * Find or create the SharedState item
 */
async function getSharedStateItem(listId) {
  // Find existing item with Title 'SharedState'
  try {
    const existing = await sharepointRequest(
      `/_api/web/lists(guid'${listId}')/items?$filter=Title eq 'SharedState'`
    )
    const item = Array.isArray(existing?.value) ? existing.value[0] : null

    if (item?.Id) {
      console.log(`[SharePoint] Found existing item: ${item.Id}`)
      return item
    }
  } catch (err) {
    console.warn('[SharePoint] Error searching for existing item:', err.message)
    throw err
  }

  // Create new item if not found
  console.log('[SharePoint] Creating new SharedState item')
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

  console.log(`[SharePoint] Created new item: ${created.Id}`)
  return sharepointRequest(`/_api/web/lists(guid'${listId}')/items(${created.Id})`)
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
    console.log('[Storage] SharePoint not configured, using localStorage')
    try {
      const stored = localStorage.getItem(STORAGE_KEY)
      if (stored) {
        const parsed = JSON.parse(stored)
        console.log('[Storage] Loaded from localStorage:', parsed)
        return parsed
      }
    } catch (err) {
      console.warn('[Storage] Error loading from localStorage:', err)
    }
    return fallback
  }

  // Try SharePoint first
  if (!hasTriedSharePoint) {
    try {
      console.log('[SharePoint] Attempting to load from SharePoint...')
      hasTriedSharePoint = true
      
      const listId = await getListId()
      const item = await getSharedStateItem(listId)

      const result = {
        subjects: parseJsonArray(item.subjectsJson, fallback.subjects),
        behaviors: parseJsonArray(item.behaviorsJson, fallback.behaviors),
        videos: parseJsonArray(item.videosJson, fallback.videos),
      }
      console.log('[SharePoint] ✓ Loaded from SharePoint:', result)
      return result
    } catch (err) {
      console.error('[SharePoint] ✗ Failed to load:', err.message)
      console.log('[Storage] Falling back to localStorage')
      
      try {
        const stored = localStorage.getItem(STORAGE_KEY)
        if (stored) {
          const parsed = JSON.parse(stored)
          console.log('[Storage] Loaded from localStorage:', parsed)
          return parsed
        }
      } catch (innerErr) {
        console.warn('[Storage] Error loading from localStorage:', innerErr)
      }
    }
  }

  return fallback
}

export async function saveSharedState(state) {
  // Always save to localStorage for local dev/backup
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state))
    console.log('[Storage] ✓ Saved to localStorage')
  } catch (err) {
    console.warn('[Storage] Could not save to localStorage:', err)
  }

  if (!isSharePointConfigured()) {
    return
  }

  // Try to also save to SharePoint if configured
  try {
    console.log('[SharePoint] Saving to SharePoint...')
    
    const listId = await getListId()
    const item = await getSharedStateItem(listId)
    const digest = await getFormDigest()

    await sharepointRequest(
      `/_api/web/lists(guid'${listId}')/items(${item.Id})`,
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
    console.log('[SharePoint] ✓ Saved to SharePoint')
  } catch (err) {
    console.error('[SharePoint] ✗ Failed to save:', err.message)
    console.log('[Storage] Data is safe in localStorage')
  }
}
