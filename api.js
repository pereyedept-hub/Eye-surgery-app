/**
 * Talks to the Google Apps Script Web App backing this app.
 *
 * After deploying Code.gs as a Web App (see the instructions at
 * the top of that file), paste the URL you're given below.
 */
const API_URL = 'https://script.google.com/macros/s/AKfycbzFNXAl-hGEK_mWj1JbbsH2vyJeanu1iOQFJkFsSIq3fNsc7_DVeNLaRhypBwZ_GDWG1w/exec'

export const isApiConfigured = () =>
  Boolean(API_URL) && !API_URL.includes('PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE')

async function apiGet(action, params = {}) {
  const query = new URLSearchParams({ action, ...params }).toString()
  const res = await fetch(`${API_URL}?${query}`)

  if (!res.ok) {
    throw new Error(`Request failed (${res.status})`)
  }

  const data = await res.json()

  if (data.error) {
    throw new Error(data.error)
  }

  return data
}

async function apiPost(action, payload = {}) {
  // Content-Type must stay "text/plain" (not "application/json") so the
  // browser treats this as a simple request and skips the CORS preflight,
  // which Apps Script web apps don't handle.
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, payload })
  })

  if (!res.ok) {
    throw new Error(`Request failed (${res.status})`)
  }

  const data = await res.json()

  if (data.error) {
    throw new Error(data.error)
  }

  return data
}

export const api = {
  getRecords: () => apiGet('getRecords').then((d) => d.records || []),

  getCompletedRecords: () =>
    apiGet('getCompletedRecords').then((d) => d.completed || {}),

  getChecklist: (patientId) =>
    apiGet('getChecklist', { patientId }).then((d) => d.checklist),

  getTheatreList: (date) =>
    apiGet('getTheatreList', { date }).then((d) => d.list || []),

  uploadRecords: (records) => apiPost('uploadRecords', { records }),

  syncFromSource: () => apiPost('syncFromSource'),

  getCaseSheet: (patientId) =>
    apiGet('getCaseSheet', { patientId }).then((d) => d.caseSheet),

  saveCaseSheet: (patientId, caseSheet) =>
    apiPost('saveCaseSheet', { patientId, caseSheet }),

  saveChecklist: (patientId, checklist) =>
    apiPost('saveChecklist', { patientId, checklist }),

  saveTheatreList: (date, list) => apiPost('saveTheatreList', { date, list })
}
