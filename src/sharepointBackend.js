/**
 * SharePoint backend using direct REST API with browser native authentication.
 * No Entra app registration required—browser handles OAuth directly with SharePoint.
 * First call may prompt user to sign in to SharePoint.
 */

const SHAREPOINT_HOSTNAME = import.meta.env.VITE_SHAREPOINT_HOSTNAME
const SHAREPOINT_SITE_PATH = import.meta.env.VITE_SHAREPOINT_SITE_PATH
const SHAREPOINT_LIST_NAME = import.meta.env.VITE_SHAREPOINT_LIST_NAME || 'LabSchoolAppState'

let siteIdCache = null
let listIdCache = null
let itemIdCache = null

export function isSharePointConfigured() {
  return Boolean(
    SHAREPOINT_HOSTNAME
      && SHAREPOINT_SITE_PATH
      && SHAREPOINT_LIST_NAME,
  )
}

/**
 * Direct SharePoint REST API call with browser authentication.
 * Browser will handle OAuth popup if needed on first request.
 */
async function sharepointRequest(path, init = {}) {
  const url = `https://${SHAREPOINT_HOSTNAME}${path}`
  const options = {
    ...init,
    headers: {
      'Content-Type': 'application/json',
      ...init.headers,
    },
    credentials: 'include', // Include cookies for browser auth
  }

  const response = await fetch(url, options)

  if (!response.ok) {
    // 401/403 typically means auth issue
    if (response.status === 401 || response.status === 403) {
      throw new Error(
        `SharePoint authentication required. Status: ${response.status}. `
        + `You may need to sign in at https://${SHAREPOINT_HOSTNAME} first.`
      )
    }
    const text = await response.text().catch(() => '')
    throw new Error(`SharePoint API error ${response.status}: ${text || response.statusText}`)
  }

  if (response.status === 204) return null
  return response.json()
}

async function getSiteId() {
  if (siteIdCache) return siteIdCache
  
  const siteUrl = `${SHAREPOINT_HOSTNAME}/${String(SHAREPOINT_SITE_PATH).replace(/^\/+/, '')}`
  const data = await sharepointRequest(`/_api/site`)
  
  siteIdCache = data.Id
  return siteIdCache
}

async function getListId(siteId) {
  if (listIdCache) return listIdCache
  
  // Try to find existing list
  const data = await sharepointRequest(
    `/_api/web/lists?$filter=Title eq '${SHAREPOINT_LIST_NAME.replace(/'/g, "''")}'`
  )

  const existing = Array.isArray(data?.value) ? data.value[0] : null
  if (existing?.Id) {
    listIdCache = existing.Id
    return listIdCache
  }

  // Create new list if not found
  const created = await sharepointRequest(`/_api/web/lists`, {
    method: 'POST',
    headers: { 'X-RequestDigest': await getFormDigest() },
    body: JSON.stringify({
      Title: SHAREPOINT_LIST_NAME,
      BaseTemplate: 100,
      Description: 'Shared state for Lab School Database',
    }),
  })
  
  listIdCache = created.Id

  // Add columns if they don't exist
  const columns = [
    { Title: 'subjectsJson', FieldTypeKind: 3 },
    { Title: 'behaviorsJson', FieldTypeKind: 3 },
    { Title: 'videosJson', FieldTypeKind: 3 },
  ]
  
  for (const col of columns) {
    await sharepointRequest(
      `/_api/web/lists(guid'${listIdCache}')/fields`,
      {
        method: 'POST',
        headers: { 'X-RequestDigest': await getFormDigest() },
        body: JSON.stringify(col),
      }
    ).catch(() => null) // Ignore if columns already exist
  }

  return listIdCache
}

async function getFormDigest() {
  const response = await sharepointRequest(`/_api/contextinfo`, {
    method: 'POST',
  })
  return response.FormDigestValue
}

async function getSharedStateItem(listId) {
  if (itemIdCache) {
    const item = await sharepointRequest(
      `/_api/web/lists(guid'${listId}')/items(${itemIdCache})`
    ).catch(() => null)
    if (item) return item
    itemIdCache = null
  }

  // Find existing item with Title 'SharedState'
  const existing = await sharepointRequest(
    `/_api/web/lists(guid'${listId}')/items?$filter=Title eq 'SharedState'`
  )
  const item = Array.isArray(existing?.value) ? existing.value[0] : null

  if (item?.Id) {
    itemIdCache = item.Id
    return item
  }

  // Create new item if not found
  const digest = await getFormDigest()
  const created = await sharepointRequest(
    `/_api/web/lists(guid'${listId}')/items`,
    {
      method: 'POST',
      headers: { 'X-RequestDigest': digest },
      body: JSON.stringify({
        Title: 'SharedState',
        subjectsJson: '[]',
        behaviorsJson: '[]',
        videosJson: '[]',
      }),
    }
  )

  itemIdCache = created.Id
  return sharepointRequest(`/_api/web/lists(guid'${listId}')/items(${itemIdCache})`)
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

  try {
    const listId = await getListId()
    const item = await getSharedStateItem(listId)

    return {
      subjects: parseJsonArray(item.subjectsJson, fallback.subjects),
      behaviors: parseJsonArray(item.behaviorsJson, fallback.behaviors),
      videos: parseJsonArray(item.videosJson, fallback.videos),
    }
  } catch (err) {
    console.error('Failed to load from SharePoint:', err)
    return fallback
  }
}

export async function saveSharedState(state) {
  if (!isSharePointConfigured()) return

  try {
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
          subjectsJson: JSON.stringify(state.subjects || []),
          behaviorsJson: JSON.stringify(state.behaviors || []),
          videosJson: JSON.stringify(state.videos || []),
        }),
      }
    )
  } catch (err) {
    console.error('Failed to save to SharePoint:', err)
    throw err
  }
}
