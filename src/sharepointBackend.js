/**
 * SharePoint backend using REST API with browser authentication.
 * This module is SharePoint-first (no silent localStorage fallback) so
 * cross-browser/device sync behavior is explicit and predictable.
 */

const SHAREPOINT_HOSTNAME = import.meta.env.VITE_SHAREPOINT_HOSTNAME || 'thecenterfordiscovery.sharepoint.com'
const SHAREPOINT_SITE_PATH = import.meta.env.VITE_SHAREPOINT_SITE_PATH || 'sites/LabSchool'
const SHAREPOINT_LIST_NAME = import.meta.env.VITE_SHAREPOINT_LIST_NAME || 'LabSchoolAppState'

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

  const response = await fetch(url, options)

  if (!response.ok) {
    const text = await response.text().catch(() => '')
    const summary = `${response.status}: ${text || response.statusText}`.slice(0, 240)
    throw new Error(`SharePoint API error ${summary}`)
  }

  if (response.status === 204) return null
  return response.json()
}

/**
 * Get or refresh form digest (CSRF token for POST/PATCH)
 * Form digest expires after ~30 minutes
 */
async function getFormDigest() {
  const now = Date.now()
  
  // Return cached digest if still valid
  if (formDigestCache && formDigestExpiry > now) return formDigestCache

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
  if (listIdCache) return listIdCache
  
  // Try to find existing list
  const data = await sharepointRequest(
    `/_api/web/lists?$filter=Title eq '${SHAREPOINT_LIST_NAME.replace(/'/g, "''")}'&$select=Id`
  )

  if (Array.isArray(data?.value) && data.value.length > 0) {
    listIdCache = data.value[0].Id
    return listIdCache
  }

  // Create new list if not found
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
    } catch (err) {
      // no-op when field exists
    }
  }

  return listIdCache
}

/**
 * Find or create the SharedState item
 */
async function getSharedStateItem(listId) {
  if (itemIdCache) return itemIdCache
  
  // Find existing item
  const existing = await sharepointRequest(
    `/_api/web/lists(guid'${listId}')/items?$filter=Title eq 'SharedState'&$select=Id`
  )

  if (Array.isArray(existing?.value) && existing.value.length > 0) {
    itemIdCache = existing.value[0].Id
    return itemIdCache
  }

  // Create new item
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

  if (!isSharePointConfigured()) return fallback

  const listId = await getListId()
  const itemId = await getSharedStateItem(listId)

  const item = await sharepointRequest(
    `/_api/web/lists(guid'${listId}')/items(${itemId})?$select=subjectsJson,behaviorsJson,videosJson`
  )

  return {
    subjects: parseJsonArray(item.subjectsJson, fallback.subjects),
    behaviors: parseJsonArray(item.behaviorsJson, fallback.behaviors),
    videos: parseJsonArray(item.videosJson, fallback.videos),
  }
}

export async function saveSharedState(state) {
  if (!isSharePointConfigured()) return

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
}
