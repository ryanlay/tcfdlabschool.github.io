import { PublicClientApplication } from '@azure/msal-browser'

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0'
const SHARED_STATE_TITLE = 'SharedState'

const AAD_CLIENT_ID = import.meta.env.VITE_AAD_CLIENT_ID
const AAD_TENANT_ID = import.meta.env.VITE_AAD_TENANT_ID
const SHAREPOINT_HOSTNAME = import.meta.env.VITE_SHAREPOINT_HOSTNAME
const SHAREPOINT_SITE_PATH = import.meta.env.VITE_SHAREPOINT_SITE_PATH
const SHAREPOINT_LIST_NAME = import.meta.env.VITE_SHAREPOINT_LIST_NAME || 'LabSchoolAppState'

let siteIdCache = null
let listIdCache = null
let itemIdCache = null

const msalApp = (AAD_CLIENT_ID && AAD_TENANT_ID)
  ? new PublicClientApplication({
      auth: {
        clientId: AAD_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${AAD_TENANT_ID}`,
      },
      cache: {
        cacheLocation: 'localStorage',
      },
    })
  : null

const loginRequest = {
  scopes: ['User.Read', 'Sites.ReadWrite.All'],
}

export function isSharePointConfigured() {
  return Boolean(
    AAD_CLIENT_ID
      && AAD_TENANT_ID
      && SHAREPOINT_HOSTNAME
      && SHAREPOINT_SITE_PATH
      && SHAREPOINT_LIST_NAME,
  )
}

async function getAccessToken() {
  if (!msalApp) throw new Error('Microsoft auth is not configured.')

  const accounts = msalApp.getAllAccounts()
  if (accounts.length === 0) {
    await msalApp.loginPopup(loginRequest)
  }

  const activeAccount = msalApp.getAllAccounts()[0]
  if (!activeAccount) throw new Error('No Microsoft account is signed in.')

  const tokenResult = await msalApp.acquireTokenSilent({
    ...loginRequest,
    account: activeAccount,
  }).catch(async () => {
    return msalApp.acquireTokenPopup({
      ...loginRequest,
      account: activeAccount,
    })
  })

  return tokenResult.accessToken
}

async function graphRequest(path, init = {}) {
  const token = await getAccessToken()
  const response = await fetch(`${GRAPH_BASE}${path}`, {
    ...init,
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${token}`,
      ...(init.headers || {}),
    },
  })

  if (!response.ok) {
    const errorText = await response.text().catch(() => '')
    throw new Error(`Graph request failed (${response.status}): ${errorText || response.statusText}`)
  }

  if (response.status === 204) return null
  return response.json()
}

async function getSiteId() {
  if (siteIdCache) return siteIdCache
  const trimmedPath = String(SHAREPOINT_SITE_PATH || '').replace(/^\/+/, '')
  const data = await graphRequest(`/sites/${SHAREPOINT_HOSTNAME}:/${trimmedPath}`)
  siteIdCache = data.id
  return siteIdCache
}

async function getListId(siteId) {
  if (listIdCache) return listIdCache
  const data = await graphRequest(`/sites/${siteId}/lists?$filter=displayName eq '${SHAREPOINT_LIST_NAME.replace(/'/g, "''")}'`)

  const existing = Array.isArray(data?.value) ? data.value[0] : null
  if (existing?.id) {
    listIdCache = existing.id
    return listIdCache
  }

  const created = await graphRequest(`/sites/${siteId}/lists`, {
    method: 'POST',
    body: JSON.stringify({
      displayName: SHAREPOINT_LIST_NAME,
      list: { template: 'genericList' },
      columns: [
        { name: 'subjectsJson', text: { allowMultipleLines: true } },
        { name: 'behaviorsJson', text: { allowMultipleLines: true } },
        { name: 'videosJson', text: { allowMultipleLines: true } },
      ],
    }),
  })
  listIdCache = created.id

  return listIdCache
}

async function getSharedStateItem(siteId, listId) {
  if (itemIdCache) {
    const item = await graphRequest(`/sites/${siteId}/lists/${listId}/items/${itemIdCache}?expand=fields`).catch(() => null)
    if (item) return item
    itemIdCache = null
  }

  const query = encodeURIComponent(`fields/Title eq '${SHARED_STATE_TITLE}'`)
  const existing = await graphRequest(`/sites/${siteId}/lists/${listId}/items?$expand=fields&$filter=${query}`)
  const item = Array.isArray(existing?.value) ? existing.value[0] : null

  if (item?.id) {
    itemIdCache = item.id
    return item
  }

  const created = await graphRequest(`/sites/${siteId}/lists/${listId}/items`, {
    method: 'POST',
    body: JSON.stringify({
      fields: {
        Title: SHARED_STATE_TITLE,
        subjectsJson: '[]',
        behaviorsJson: '[]',
        videosJson: '[]',
      },
    }),
  })

  itemIdCache = created.id
  return graphRequest(`/sites/${siteId}/lists/${listId}/items/${itemIdCache}?expand=fields`)
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

  const siteId = await getSiteId()
  const listId = await getListId(siteId)
  const item = await getSharedStateItem(siteId, listId)
  const fields = item?.fields || {}

  return {
    subjects: parseJsonArray(fields.subjectsJson, fallback.subjects),
    behaviors: parseJsonArray(fields.behaviorsJson, fallback.behaviors),
    videos: parseJsonArray(fields.videosJson, fallback.videos),
  }
}

export async function saveSharedState(state) {
  if (!isSharePointConfigured()) return

  const siteId = await getSiteId()
  const listId = await getListId(siteId)
  const item = await getSharedStateItem(siteId, listId)

  await graphRequest(`/sites/${siteId}/lists/${listId}/items/${item.id}/fields`, {
    method: 'PATCH',
    body: JSON.stringify({
      subjectsJson: JSON.stringify(state.subjects || []),
      behaviorsJson: JSON.stringify(state.behaviors || []),
      videosJson: JSON.stringify(state.videos || []),
    }),
  })
}
