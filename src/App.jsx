import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import { isSharePointConfigured, loadSharedState, saveSharedState, StaleWriteError } from './sharepointBackend'

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

// Transient network blips (Wi-Fi reconnecting after sleep, a momentary drop, etc.) surface as
// "TypeError: Failed to fetch" and shouldn't require the user to notice an error banner and click
// Retry themselves. Give the initial load a few quick automatic attempts before giving up.
async function loadSharedStateWithRetry(defaultState, attempts = 2, delayMs = 500) {
  let lastError
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      return await loadSharedState(defaultState)
    } catch (error) {
      lastError = error
      if (attempt < attempts) await wait(delayMs * attempt)
    }
  }
  throw lastError
}

const defaultSubjects = [
  { id: 1, subjectCode: 'CSH01', displayName: 'Classroom Student 01', isActive: true, dateOfBirth: null, labSchoolStartDate: null, labSchoolEndDate: null, personId: null, targetBehaviors: [] },
  { id: 2, subjectCode: 'CSH02', displayName: 'Classroom Student 02', isActive: true, dateOfBirth: null, labSchoolStartDate: null, labSchoolEndDate: null, personId: null, targetBehaviors: [] },
  { id: 3, subjectCode: 'CSH03', displayName: 'Classroom Student 03', isActive: true, dateOfBirth: null, labSchoolStartDate: null, labSchoolEndDate: null, personId: null, targetBehaviors: [] },
]

// Validated list of target behaviors used on Subject historical profiles.
// Kept separate from `behaviors` (used for per-video behavior-occurrence logging).
const targetBehaviorOptions = [
  'Aggression',
  'Self-injury',
  'Elopement',
  'Pica',
  'Disruptive Behaviors',
  'Mouthing',
  'Impulsivity',
  'Agitation',
  'Refusal Behavior',
  'Off-Task Behavior',
  'Property Destruction',
  'Ritualistic Behavior',
  'Unsanitary Behavior',
  'Sensory Stimulation',
  'Stereotypy/Repetitive Behavior',
  'Inappropriate Social Behavior',
  'Inappropriate Touch',
  'Sexually Inappropriate Behavior',
  'Disrobing',
  'Food Stealing',
]

// Historical subject roster to make available for profile lookup/entry alongside any existing subjects.
const historicalSubjectCodes = [
  'S01', 'S02', 'S03', 'S04', 'S05', 'S06', 'S07', 'S08', 'S09', 'S10',
  'S11', 'S12', 'S13', 'S14', 'S15', 'S16', 'S17', 'S18', 'S19', 'S20',
  'S21', 'S22', 'S23',
  'AS01', 'AS02', 'AS03', 'AS04', 'AS05', 'AS06', 'AS07', 'AS08', 'AS09', 'AS10',
  'AS11', 'AS12', 'AS13', 'AS14', 'AS15',
  'RS01', 'RS02', 'RS03', 'RS04',
]

// Known Person_IDs for historical subjects (S04, S08, S18 have no known Person_ID).
const historicalPersonIds = {
  S01: 'P55943', S02: 'P55659', S03: 'P54788', S05: 'P52801', S06: 'P56145',
  S07: 'P52164', S09: 'P55914', S10: 'P52233', S11: 'P59798', S12: 'P51352',
  S13: 'P52709', S14: 'P57933', S15: 'P51538', S16: 'P55520', S17: 'P59495',
  S19: 'P56833', S20: 'P51762', S21: 'P51948', S22: 'P53890', S23: 'P54824',
  AS01: 'P53960', AS02: 'P57072', AS03: 'P59825', AS04: 'P54467', AS05: 'P54111',
  AS06: 'P58432', AS07: 'P54002', AS08: 'P51327', AS09: 'P50217', AS10: 'P53202',
  AS11: 'P58716', AS12: 'P57335', AS13: 'P51947', AS14: 'P59079', AS15: 'P50587',
  RS01: 'P56631', RS02: 'P58976', RS03: 'P58719', RS04: 'P59081',
}

// Additively merges the historical roster into an existing subjects list, without touching
// or removing any subjects/data that are already there. Returns the same array reference
// when nothing needs to change, so it doesn't trigger a needless re-save on every load.
function withHistoricalRoster(existingSubjects) {
  const existingCodes = new Set(existingSubjects.map((subject) => subject.subjectCode.toUpperCase()))
  const missing = historicalSubjectCodes.filter((code) => !existingCodes.has(code.toUpperCase()))

  const needsPersonIdBackfill = existingSubjects.some(
    (subject) => !subject.personId && historicalPersonIds[subject.subjectCode.toUpperCase()],
  )

  if (missing.length === 0 && !needsPersonIdBackfill) {
    return existingSubjects
  }

  let nextId = existingSubjects.length ? Math.max(...existingSubjects.map((subject) => subject.id)) + 1 : 1
  const additions = missing.map((code) => ({
    id: nextId++,
    subjectCode: code,
    displayName: code,
    isActive: true,
    dateOfBirth: null,
    labSchoolStartDate: null,
    labSchoolEndDate: null,
    personId: historicalPersonIds[code] || null,
    targetBehaviors: [],
  }))
  const merged = missing.length ? [...existingSubjects, ...additions] : existingSubjects

  // Additively backfill Person_ID for known historical subjects that don't have one yet.
  // Never overwrites a personId that's already set (e.g. manually edited).
  return merged.map((subject) => (
    !subject.personId && historicalPersonIds[subject.subjectCode.toUpperCase()]
      ? { ...subject, personId: historicalPersonIds[subject.subjectCode.toUpperCase()] }
      : subject
  ))
}

const defaultBehaviors = [
  { id: 1, name: 'Aggression', isActive: true },
  { id: 2, name: 'Self-injury', isActive: true },
  { id: 3, name: 'Motor Disruption', isActive: true },
]

const defaultVideos = [
  {
    id: 1,
    recordStartTime: '2026-05-20T09:00:00.000Z',
    durationSeconds: 1200,
    notes: 'Morning classroom session',
    uploadedToSharePoint: false,
    subjectCodes: ['CSH01', 'CSH02'],
    occurrences: [
      { subjectCode: 'CSH01', behaviorTypeName: 'Aggression', notes: null },
      { subjectCode: 'CSH02', behaviorTypeName: 'Motor Disruption', notes: 'brief' },
    ],
    createdAt: '2026-05-20T10:00:00.000Z',
  },
  {
    id: 2,
    recordStartTime: '2026-05-21T13:15:00.000Z',
    durationSeconds: 900,
    notes: 'Afternoon group work',
    uploadedToSharePoint: true,
    subjectCodes: ['CSH01', 'CSH03'],
    occurrences: [
      { subjectCode: 'CSH03', behaviorTypeName: 'Self-injury', notes: null },
    ],
    createdAt: '2026-05-21T13:20:00.000Z',
  },
]

const views = {
  home: 'home',
  intake: 'intake',
  review: 'review',
  data: 'data',
  admin: 'admin',
}


function localInputValue(date = new Date()) {
  const copy = new Date(date)
  copy.setMinutes(copy.getMinutes() - copy.getTimezoneOffset())
  return copy.toISOString().slice(0, 19)
}

function fmtDate(value) {
  if (!value) return '—'
  const d = new Date(value)
  const month = String(d.getMonth() + 1).padStart(2, '0')
  const day = String(d.getDate()).padStart(2, '0')
  const year = d.getFullYear()
  let hours = d.getHours()
  const ampm = hours >= 12 ? 'PM' : 'AM'
  hours = hours % 12 || 12
  const mins = String(d.getMinutes()).padStart(2, '0')
  const secs = String(d.getSeconds()).padStart(2, '0')
  return `${month}/${day}/${year} ${hours}:${mins}:${secs} ${ampm}`
}

function fmtDuration(seconds) {
  const totalMinutes = Math.round(Number(seconds || 0) / 60)
  return `${totalMinutes} min`
}

function esc(value) {
  return String(value ?? '')
}

function sharePointFolderUrl(isoDate) {
  const date = new Date(isoDate)
  const yyyy = date.getFullYear()
  const mm = String(date.getMonth() + 1).padStart(2, '0')
  const monthName = date.toLocaleString('en-US', { month: 'long' })
  const monthFolder = `${mm} ${monthName}`
  const dd = String(date.getDate()).padStart(2, '0')
  return `https://thecenterfordiscovery.sharepoint.com/:f:/r/sites/LabSchool/Shared%20Documents/Videos/${yyyy}/${encodeURIComponent(monthFolder)}/${dd}?csf=1&web=1`
}

function sortRows(rows, sortCol, sortDir, getter) {
  return [...rows].sort((left, right) => {
    const a = getter(left, sortCol)
    const b = getter(right, sortCol)
    const comparison = typeof a === 'number' && typeof b === 'number'
      ? a - b
      : String(a ?? '').localeCompare(String(b ?? ''), undefined, { numeric: true, sensitivity: 'base' })
    return comparison * sortDir
  })
}

function Table({ columns, rows, sortState, onSort, emptyMessage }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {columns.map((column) => (
              <th key={column.key}>
                {column.sortable ? (
                  <button className="th-button" type="button" onClick={() => onSort(column.key)}>
                    {column.label}
                    {sortState?.sortCol === column.key ? (sortState.sortDir === 1 ? ' ▲' : ' ▼') : ''}
                  </button>
                ) : (
                  column.label
                )}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 ? (
            <tr>
              <td colSpan={columns.length} className="muted center">{emptyMessage}</td>
            </tr>
          ) : (
            rows.map((row) => (
              <tr key={row.__key}>
                {columns.map((column) => (
                  <td key={column.key}>{column.render(row)}</td>
                ))}
              </tr>
            ))
          )}
        </tbody>
      </table>
    </div>
  )
}

function EditVideoPanel({
  editingVideoId,
  videos,
  editVideoStart,
  setEditVideoStart,
  editVideoDuration,
  setEditVideoDuration,
  editVideoNotes,
  setEditVideoNotes,
  editVideoSubjects,
  subjects,
  toggleEditSubject,
  behaviors,
  editVideoOccurrenceMap,
  toggleEditOccurrence,
  setVideos,
  cancelEditVideo,
  saveEditVideo,
}) {
  if (editingVideoId === null) return null

  const editing = videos.find((video) => video.id === editingVideoId)
  if (!editing) return null

  return (
    <div className="edit-panel">
      <h3>Editing Video {editingVideoId}</h3>
      <div className="form-grid">
        <label>
          Start time
          <input type="datetime-local" step="1" value={editVideoStart} onChange={(event) => setEditVideoStart(event.target.value)} />
        </label>
        <label>
          Duration (minutes)
          <input type="number" min="1" step="1" value={editVideoDuration} onChange={(event) => setEditVideoDuration(event.target.value)} />
        </label>
        <label className="full">
          Notes
          <textarea rows="3" value={editVideoNotes} onChange={(event) => setEditVideoNotes(event.target.value)} placeholder="Optional notes" />
        </label>
      </div>
      <label style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
        <input
          type="checkbox"
          checked={editing.uploadedToSharePoint || false}
          onChange={(event) =>
            setVideos((current) =>
              current.map((video) =>
                video.id === editingVideoId
                  ? { ...video, uploadedToSharePoint: event.target.checked }
                  : video,
              ),
            )
          }
        />
        <span>Uploaded to SharePoint</span>
      </label>

      <p className="muted" style={{ marginTop: '12px', marginBottom: '4px' }}>Subjects present</p>
      <div className="chip-grid">
        {subjects.map((subject) => (
          <label key={subject.subjectCode} className={`chip ${editVideoSubjects.includes(subject.subjectCode) ? 'selected' : ''}`}>
            <input
              type="checkbox"
              checked={editVideoSubjects.includes(subject.subjectCode)}
              onChange={() => toggleEditSubject(subject.subjectCode)}
            />
            <span>{subject.subjectCode}</span>
          </label>
        ))}
      </div>

      {editVideoSubjects.length > 0 && (
        <>
          <p className="muted" style={{ marginTop: '12px', marginBottom: '4px' }}>Behavior occurrences</p>
          <div className="stack">
            {editVideoSubjects.map((subjectCode) => (
              <div key={subjectCode} className="card inset">
                <h4 style={{ margin: '0 0 8px' }}>{subjectCode}</h4>
                <div className="chip-grid">
                  {behaviors.map((behavior) => {
                    const checked = Boolean(editVideoOccurrenceMap[`${subjectCode}::${behavior.name}`])
                    const label = behavior.isActive ? behavior.name : `${behavior.name} (inactive)`
                    return (
                      <label key={behavior.name} className={`chip ${checked ? 'selected' : ''}`}>
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() => toggleEditOccurrence(subjectCode, behavior.name)}
                        />
                        <span>{label}</span>
                      </label>
                    )
                  })}
                </div>
              </div>
            ))}
          </div>
        </>
      )}

      <div className="button-row">
        <button type="button" className="secondary" onClick={cancelEditVideo}>Cancel</button>
        <button type="button" className="primary" onClick={saveEditVideo}>Save Changes</button>
      </div>
    </div>
  )
}

function App() {
  const [ready, setReady] = useState(false)
  const [view, setView] = useState(views.home)
  const [subjects, setSubjects] = useState(defaultSubjects)
  const [behaviors, setBehaviors] = useState(defaultBehaviors)
  const [videos, setVideos] = useState(defaultVideos)
  const [status, setStatus] = useState({ home: '', recording: '', review: '', data: '', admin: '', profile: '', conflict: '' })

  const [intakeStep, setIntakeStep] = useState(1)
  const [recordStartTime, setRecordStartTime] = useState(localInputValue())
  const [durationMinutes, setDurationMinutes] = useState('')
  const [videoNotes, setVideoNotes] = useState('')
  const [selectedSubjects, setSelectedSubjects] = useState([])
  const [occurrenceMap, setOccurrenceMap] = useState({})

  const [dataSearch, setDataSearch] = useState('')
  const [dataSort, setDataSort] = useState({ sortCol: 'recordStartTime', sortDir: -1 })

  const [q1Sort, setQ1Sort] = useState({ sortCol: 'recordStartTime', sortDir: -1 })
  const [q2Sort, setQ2Sort] = useState({ sortCol: 'recordStartTime', sortDir: -1 })
  const [q3Subject, setQ3Subject] = useState('CSH01')
  const [q4Behavior, setQ4Behavior] = useState('Aggression')
  const [q5Sort, setQ5Sort] = useState({ sortCol: 'recordStartTime', sortDir: -1 })
  const [q6Sort, setQ6Sort] = useState({ sortCol: 'recordStartTime', sortDir: -1 })

  const [newSubjectCode, setNewSubjectCode] = useState('')
  const [newSubjectName, setNewSubjectName] = useState('')
  const [newBehaviorName, setNewBehaviorName] = useState('')

  // Subject historical profile
  const [profileSubjectCode, setProfileSubjectCode] = useState('')
  const [profileLoaded, setProfileLoaded] = useState(false)
  const [profilePersonId, setProfilePersonId] = useState('')
  const [profileDob, setProfileDob] = useState('')
  const [profileStartDate, setProfileStartDate] = useState('')
  const [profileEndDate, setProfileEndDate] = useState('')
  const [profileTargetBehaviors, setProfileTargetBehaviors] = useState([])
  const [newProfileSubjectCode, setNewProfileSubjectCode] = useState('')

  // Log entry editing
  const [editingVideoId, setEditingVideoId] = useState(null)
  const [editVideoStart, setEditVideoStart] = useState('')
  const [editVideoDuration, setEditVideoDuration] = useState('')
  const [editVideoNotes, setEditVideoNotes] = useState('')
  const [editVideoSubjects, setEditVideoSubjects] = useState([])
  const [editVideoOccurrenceMap, setEditVideoOccurrenceMap] = useState({})
  const [adminLogSort, setAdminLogSort] = useState({ sortCol: 'recordStartTime', sortDir: -1 })
  const importFileInputRef = useRef(null)
  const hasHydratedFromBackendRef = useRef(false)
  const saveTimerRef = useRef(null)
  // Tracks the `updated_at` this tab last saw from Supabase. Every save is checked against the
  // server's CURRENT updated_at first - if it no longer matches (another tab/device saved since),
  // the save is refused (StaleWriteError) instead of silently clobbering that newer data.
  const updatedAtRef = useRef(null)
  const [loading, setLoading] = useState(false)

  const initialize = useCallback(async () => {
    setLoading(true)
    show('conflict', '', false)
    try {
      if (!isSharePointConfigured()) {
        show('home', 'Supabase is not configured. Add Supabase environment settings.', true)
        setSubjects(withHistoricalRoster(defaultSubjects))
        setBehaviors(defaultBehaviors)
        setVideos(defaultVideos)
        setReady(true)
        return
      }

      const state = await loadSharedStateWithRetry({
        subjects: defaultSubjects,
        behaviors: defaultBehaviors,
        videos: defaultVideos,
      })

      setSubjects(withHistoricalRoster(Array.isArray(state.subjects) && state.subjects.length ? state.subjects : defaultSubjects))
      setBehaviors(Array.isArray(state.behaviors) && state.behaviors.length ? state.behaviors : defaultBehaviors)
      setVideos(Array.isArray(state.videos) ? state.videos : defaultVideos)
      updatedAtRef.current = state.updatedAt ?? null
      hasHydratedFromBackendRef.current = true
      setReady(true)
      show('home', '', false)
    } catch (error) {
      // IMPORTANT: a failed load must NEVER be mistaken for "my data is gone". It isn't -
      // hasHydratedFromBackendRef stays false, so the auto-save effect below refuses to run
      // and can't overwrite the real Supabase row with this placeholder data. This banner is
      // the only thing telling the user what actually happened, so it must stay visible
      // (not auto-dismiss) until a retry succeeds.
      const detail = error instanceof Error ? error.message : 'Unknown error'
      show(
        'home',
        `Could not reach Supabase to load your saved data (showing local placeholder data — nothing on the server was changed or deleted). Check your network/firewall, then Retry. Error: ${detail}`,
        true,
      )
      setSubjects(withHistoricalRoster(defaultSubjects))
      setBehaviors(defaultBehaviors)
      setVideos(defaultVideos)
      setReady(true)
    } finally {
      setLoading(false)
    }
  }, [])

  useEffect(() => {
    initialize()
  }, [initialize])

  useEffect(() => {
    if (!ready) return
    if (!hasHydratedFromBackendRef.current) return
    if (!isSharePointConfigured()) return

    if (saveTimerRef.current) {
      clearTimeout(saveTimerRef.current)
    }

    saveTimerRef.current = setTimeout(async () => {
      try {
        const saved = await saveSharedState({ subjects, behaviors, videos }, { expectedUpdatedAt: updatedAtRef.current })
        updatedAtRef.current = saved.updatedAt
        show('admin', 'Saved to Supabase.', false)
      } catch (error) {
        if (error instanceof StaleWriteError) {
          show('conflict', `${error.message} (Your most recent change on this device was NOT saved.)`, true)
          return
        }
        const detail = error instanceof Error ? error.message : 'Unknown error'
        show('admin', `Save to Supabase failed: ${detail}`, true)
      }
    }, 500)

    return () => {
      if (saveTimerRef.current) {
        clearTimeout(saveTimerRef.current)
      }
    }
  }, [ready, subjects, behaviors, videos])

  const activeSubjects = useMemo(() => subjects.filter((subject) => subject.isActive), [subjects])
  const activeBehaviors = useMemo(() => behaviors.filter((behavior) => behavior.isActive), [behaviors])
  const activeSubjectCodeSet = useMemo(() => new Set(activeSubjects.map((subject) => subject.subjectCode)), [activeSubjects])
  const selectedActiveSubjects = useMemo(
    () => selectedSubjects.filter((subjectCode) => activeSubjectCodeSet.has(subjectCode)),
    [selectedSubjects, activeSubjectCodeSet],
  )
  const q3SubjectOptions = useMemo(() => {
    const codes = new Set()
    for (const subject of subjects) codes.add(subject.subjectCode)
    for (const video of videos) {
      for (const code of video.subjectCodes || []) codes.add(code)
    }
    return [...codes].sort((left, right) => left.localeCompare(right, undefined, { numeric: true, sensitivity: 'base' }))
  }, [subjects, videos])

  useEffect(() => {
    if (!q3SubjectOptions.length) return
    if (!q3SubjectOptions.includes(q3Subject)) {
      setQ3Subject(q3SubjectOptions[0])
    }
  }, [q3SubjectOptions, q3Subject])

  // Historical review isn't limited to currently-active subjects, so this list is unfiltered.
  const profileSubjectOptions = useMemo(
    () => [...subjects].sort((left, right) => left.subjectCode.localeCompare(right.subjectCode, undefined, { numeric: true, sensitivity: 'base' })),
    [subjects],
  )
  const profileVideoCount = useMemo(
    () => videos.filter((video) => video.subjectCodes.includes(profileSubjectCode)).length,
    [videos, profileSubjectCode],
  )
  const profileOccurrenceCount = useMemo(
    () => videos.reduce((total, video) => total + video.occurrences.filter((occurrence) => occurrence.subjectCode === profileSubjectCode).length, 0),
    [videos, profileSubjectCode],
  )

  useEffect(() => {
    setSelectedSubjects((current) => current.filter((subjectCode) => activeSubjectCodeSet.has(subjectCode)))
    setOccurrenceMap((current) => {
      const next = {}
      for (const [key, value] of Object.entries(current)) {
        const [subjectCode] = key.split('::')
        if (activeSubjectCodeSet.has(subjectCode)) {
          next[key] = value
        }
      }
      return next
    })
  }, [activeSubjectCodeSet])

  const q1Rows = useMemo(() => {
    const rows = videos.map((video) => ({
      __key: `q1-${video.id}`,
      ...video,
      subjectCount: video.subjectCodes.length,
      behaviorCount: video.occurrences.length,
    }))
    return sortRows(rows, q1Sort.sortCol, q1Sort.sortDir, (row, key) => row[key])
  }, [videos, q1Sort])

  const q2Rows = useMemo(() => {
    const rows = videos.flatMap((video) =>
      video.occurrences.map((occurrence, index) => ({
        __key: `q2-${video.id}-${index}`,
        ...occurrence,
        recordStartTime: video.recordStartTime,
        videoId: video.id,
        uploadedToSharePoint: video.uploadedToSharePoint,
        notes: occurrence.notes,
      })),
    )
    return sortRows(rows, q2Sort.sortCol, q2Sort.sortDir, (row, key) => row[key])
  }, [videos, q2Sort])

  const q3Rows = useMemo(() => {
    const rows = videos.filter((video) => video.subjectCodes.includes(q3Subject)).map((video) => ({
      __key: `q3-${video.id}`,
      ...video,
    }))
    return sortRows(rows, dataSort.sortCol, dataSort.sortDir, (row, key) => row[key])
  }, [videos, q3Subject, dataSort])

  const q4Rows = useMemo(() => {
    const rows = videos.filter((video) => video.occurrences.some((occurrence) => occurrence.behaviorTypeName === q4Behavior)).map((video) => ({
      __key: `q4-${video.id}`,
      ...video,
    }))
    return sortRows(rows, dataSort.sortCol, dataSort.sortDir, (row, key) => row[key])
  }, [videos, q4Behavior, dataSort])

  const q5Rows = useMemo(() => {
    const rows = videos.filter((video) => !video.uploadedToSharePoint).map((video) => ({
      __key: `q5-${video.id}`,
      ...video,
    }))
    return sortRows(rows, q5Sort.sortCol, q5Sort.sortDir, (row, key) => row[key])
  }, [videos, q5Sort])

  const q6Rows = useMemo(() => {
    const rows = videos.filter((video) => video.uploadedToSharePoint).map((video) => ({
      __key: `q6-${video.id}`,
      ...video,
    }))
    return sortRows(rows, q6Sort.sortCol, q6Sort.sortDir, (row, key) => row[key])
  }, [videos, q6Sort])

  const adminLogRows = useMemo(() => {
    const rows = videos.map((video) => ({ __key: `admin-log-${video.id}`, ...video }))
    return sortRows(rows, adminLogSort.sortCol, adminLogSort.sortDir, (row, key) => row[key])
  }, [videos, adminLogSort])

  const dataRows = useMemo(() => {
    const filtered = videos.filter((video) => {
      const haystack = `${video.id} ${video.notes || ''} ${video.subjectCodes.join(' ')} ${video.occurrences.map((x) => x.behaviorTypeName).join(' ')}`.toLowerCase()
      return haystack.includes(dataSearch.toLowerCase())
    })
    return sortRows(filtered.map((video) => ({ __key: `data-${video.id}`, ...video })), dataSort.sortCol, dataSort.sortDir, (row, key) => row[key])
  }, [videos, dataSearch, dataSort])

  function show(msgKey, message, isError = false) {
    setStatus((current) => ({ ...current, [msgKey]: message ? { text: message, isError } : '' }))
  }

  function resetIntake() {
    setIntakeStep(1)
    setRecordStartTime(localInputValue())
    setDurationMinutes('')
    setVideoNotes('')
    setSelectedSubjects([])
    setOccurrenceMap({})
    show('recording', '', false)
  }

  function toggleSubject(code) {
    setSelectedSubjects((current) => {
      const next = current.includes(code) ? current.filter((value) => value !== code) : [...current, code]
      // Clean up occurrence map when subject is removed
      if (!next.includes(code)) {
        setOccurrenceMap((currentMap) => {
          const updated = { ...currentMap }
          for (const key in updated) {
            if (key.startsWith(`${code}::`)) {
              delete updated[key]
            }
          }
          return updated
        })
      }
      return next
    })
  }

  function toggleOccurrence(subjectCode, behaviorName) {
    const key = `${subjectCode}::${behaviorName}`
    setOccurrenceMap((current) => ({ ...current, [key]: !current[key] }))
  }

  function collectOccurrences() {
    const rows = []
    for (const subjectCode of selectedSubjects) {
      for (const behavior of activeBehaviors) {
        const key = `${subjectCode}::${behavior.name}`
        if (occurrenceMap[key]) {
          rows.push({ subjectCode, behaviorTypeName: behavior.name, notes: null })
        }
      }
    }
    return rows
  }

  function handleSave(addAnother = false) {
    const minutes = Number(durationMinutes)
    if (!recordStartTime || !Number.isFinite(minutes) || minutes < 1 || selectedSubjects.length === 0) {
      show('recording', 'Enter a start time, a duration, and at least one subject.', true)
      return
    }

    const nextVideo = {
      id: videos.length ? Math.max(...videos.map((video) => video.id)) + 1 : 1,
      recordStartTime: new Date(recordStartTime).toISOString(),
      durationSeconds: Math.round(minutes * 60),
      notes: videoNotes.trim() || null,
      uploadedToSharePoint: false,
      subjectCodes: selectedSubjects,
      occurrences: collectOccurrences(),
      createdAt: new Date().toISOString(),
    }

    setVideos((current) => [nextVideo, ...current])
    show('recording', `Saved video ${nextVideo.id}.`, false)

    if (addAnother) {
      setRecordStartTime(localInputValue())
      setDurationMinutes('')
      setVideoNotes('')
      setSelectedSubjects([])
      setOccurrenceMap({})
      setIntakeStep(1)
    } else {
      resetIntake()
      setView(views.home)
    }
  }

  function toggleUpload(videoId, nextValue) {
    setVideos((current) => current.map((video) => (video.id === videoId ? { ...video, uploadedToSharePoint: nextValue } : video)))
    show('review', `Video ${videoId} marked ${nextValue ? 'uploaded' : 'pending'}.`, false)
  }

  function addSubject() {
    const code = newSubjectCode.trim().toUpperCase()
    if (!code) return show('admin', 'Enter a subject code.', true)

    if (subjects.some((subject) => subject.subjectCode.toUpperCase() === code)) {
      return show('admin', 'That subject already exists.', true)
    }

    const nextSubject = {
      id: subjects.length ? Math.max(...subjects.map((subject) => subject.id)) + 1 : 1,
      subjectCode: code,
      displayName: code,
      isActive: true,
      dateOfBirth: null,
      labSchoolStartDate: null,
      labSchoolEndDate: null,
      personId: null,
      targetBehaviors: [],
    }

    setSubjects((current) => [nextSubject, ...current])
    setNewSubjectCode('')
    show('admin', `Added subject ${code}.`, false)
  }

  function loadSubjectProfile() {
    if (!profileSubjectCode) return show('profile', 'Select a subject first.', true)

    const subject = subjects.find((entry) => entry.subjectCode === profileSubjectCode)
    if (!subject) return show('profile', 'Subject not found.', true)

    setProfilePersonId(subject.personId || '')
    setProfileDob(subject.dateOfBirth ? subject.dateOfBirth.slice(0, 10) : '')
    setProfileStartDate(subject.labSchoolStartDate ? subject.labSchoolStartDate.slice(0, 10) : '')
    setProfileEndDate(subject.labSchoolEndDate ? subject.labSchoolEndDate.slice(0, 10) : '')
    setProfileTargetBehaviors(Array.isArray(subject.targetBehaviors) ? subject.targetBehaviors : [])
    setProfileLoaded(true)
    show('profile', '', false)
  }

  function toggleProfileTargetBehavior(name) {
    setProfileTargetBehaviors((current) => (current.includes(name) ? current.filter((item) => item !== name) : [...current, name]))
  }

  function saveSubjectProfile() {
    if (!profileSubjectCode) return show('profile', 'Select a subject first.', true)
    if (profileDob && profileStartDate && profileDob > profileStartDate) {
      return show('profile', 'Date of birth cannot be after the Lab School start date.', true)
    }
    if (profileStartDate && profileEndDate && profileEndDate < profileStartDate) {
      return show('profile', 'Lab School end date cannot be before the start date.', true)
    }

    setSubjects((current) => current.map((entry) => (
      entry.subjectCode === profileSubjectCode
        ? {
            ...entry,
            personId: profilePersonId.trim() || null,
            dateOfBirth: profileDob || null,
            labSchoolStartDate: profileStartDate || null,
            labSchoolEndDate: profileEndDate || null,
            targetBehaviors: profileTargetBehaviors,
          }
        : entry
    )))
    show('profile', `Profile saved for ${profileSubjectCode}.`, false)
  }

  function addProfileSubject() {
    const code = newProfileSubjectCode.trim().toUpperCase()
    if (!code) return show('profile', 'Enter a subject ID.', true)

    if (subjects.some((subject) => subject.subjectCode.toUpperCase() === code)) {
      return show('profile', 'That subject already exists.', true)
    }

    const nextSubject = {
      id: subjects.length ? Math.max(...subjects.map((subject) => subject.id)) + 1 : 1,
      subjectCode: code,
      displayName: code,
      isActive: true,
      dateOfBirth: null,
      labSchoolStartDate: null,
      labSchoolEndDate: null,
      personId: null,
      targetBehaviors: [],
    }

    setSubjects((current) => [nextSubject, ...current])
    setNewProfileSubjectCode('')
    setProfileSubjectCode(code)
    show('profile', `Added subject ${code}.`, false)
  }

  function addBehavior() {
    const name = newBehaviorName.trim()
    if (!name) return show('admin', 'Enter a behavior name.', true)

    if (behaviors.some((behavior) => behavior.name.toLowerCase() === name.toLowerCase())) {
      return show('admin', 'That behavior already exists.', true)
    }

    const nextBehavior = {
      id: behaviors.length ? Math.max(...behaviors.map((behavior) => behavior.id)) + 1 : 1,
      name,
      isActive: true,
    }

    setBehaviors((current) => [nextBehavior, ...current])
    setNewBehaviorName('')
    show('admin', `Added behavior ${name}.`, false)
  }

  function startEditVideo(video) {
    setEditingVideoId(video.id)
    setEditVideoStart(localInputValue(new Date(video.recordStartTime)))
    setEditVideoDuration(String(Math.round(video.durationSeconds / 60)))
    setEditVideoNotes(video.notes || '')
    setEditVideoSubjects([...video.subjectCodes])
    const map = {}
    for (const occ of video.occurrences) {
      map[`${occ.subjectCode}::${occ.behaviorTypeName}`] = true
    }
    setEditVideoOccurrenceMap(map)
    show('admin', '', false)
  }

  function cancelEditVideo() {
    setEditingVideoId(null)
    show('admin', '', false)
  }

  function saveEditVideo() {
    const minutes = Number(editVideoDuration)
    if (!editVideoStart || !Number.isFinite(minutes) || minutes < 1 || editVideoSubjects.length === 0) {
      show('admin', 'Enter a start time, a duration, and at least one subject.', true)
      return
    }
    const occurrences = []
    for (const subjectCode of editVideoSubjects) {
      for (const behavior of behaviors) {
        const key = `${subjectCode}::${behavior.name}`
        if (editVideoOccurrenceMap[key]) {
          occurrences.push({ subjectCode, behaviorTypeName: behavior.name, notes: null })
        }
      }
    }
    setVideos((current) =>
      current.map((video) =>
        video.id === editingVideoId
          ? {
              ...video,
              recordStartTime: new Date(editVideoStart).toISOString(),
              durationSeconds: Math.round(minutes * 60),
              notes: editVideoNotes.trim() || null,
              subjectCodes: editVideoSubjects,
              occurrences,
            }
          : video,
      ),
    )
    setEditingVideoId(null)
    show('admin', `Video ${editingVideoId} updated.`, false)
  }

  function deleteVideo(videoId) {
    if (!window.confirm(`Delete video ${videoId}? This cannot be undone.`)) return
    setVideos((current) => current.filter((video) => video.id !== videoId))
    if (editingVideoId === videoId) setEditingVideoId(null)
    show('admin', `Video ${videoId} deleted.`, false)
  }

  function exportAllData() {
    const payload = {
      exportedAt: new Date().toISOString(),
      subjects,
      behaviors,
      videos,
    }
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const anchor = document.createElement('a')
    const stamp = new Date().toISOString().replace(/[:.]/g, '-')
    anchor.href = url
    anchor.download = `lab-school-recordings-backup-${stamp}.json`
    document.body.appendChild(anchor)
    anchor.click()
    document.body.removeChild(anchor)
    URL.revokeObjectURL(url)
    show('admin', 'Backup file downloaded.', false)
  }

  function openImportDialog() {
    importFileInputRef.current?.click()
  }

  async function importAllData(event) {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      const text = await file.text()
      const parsed = JSON.parse(text)
      const nextSubjects = Array.isArray(parsed?.subjects) ? parsed.subjects : null
      const nextBehaviors = Array.isArray(parsed?.behaviors) ? parsed.behaviors : null
      const nextVideos = Array.isArray(parsed?.videos) ? parsed.videos : null

      if (!nextSubjects || !nextBehaviors || !nextVideos) {
        show('admin', 'Invalid backup file format.', true)
        return
      }

      if (!window.confirm(`Import backup from ${file.name}? This will replace current on-screen data.`)) {
        return
      }

      const rosteredSubjects = withHistoricalRoster(nextSubjects)
      setSubjects(rosteredSubjects)
      setBehaviors(nextBehaviors)
      setVideos(nextVideos)
      setEditingVideoId(null)

      if (isSharePointConfigured()) {
        if (saveTimerRef.current) {
          clearTimeout(saveTimerRef.current)
        }
        // Save immediately (awaited) instead of relying on the debounced auto-save effect,
        // so the import can't be lost if the tab is closed right after confirming. Still checked
        // against the server's current updated_at - if someone else saved more recent changes
        // since this tab last loaded, we refuse to blindly overwrite them.
        const saved = await saveSharedState(
          { subjects: rosteredSubjects, behaviors: nextBehaviors, videos: nextVideos },
          { expectedUpdatedAt: updatedAtRef.current },
        )
        updatedAtRef.current = saved.updatedAt
        show('admin', `Imported and saved to Supabase: ${nextVideos.length} videos loaded.`, false)
      } else {
        show('admin', `Imported backup: ${nextVideos.length} videos loaded (not saved — Supabase not configured).`, false)
      }
    } catch (error) {
      if (error instanceof StaleWriteError) {
        show(
          'admin',
          `Import NOT saved: someone else saved changes to the shared data since this page loaded. Click Retry on the home screen to reload the latest data, then try the import again.`,
          true,
        )
        return
      }
      const detail = error instanceof Error ? error.message : 'Unknown error'
      show('admin', `Could not import backup file: ${detail}`, true)
    } finally {
      event.target.value = ''
    }
  }

  function toggleEditSubject(code) {
    setEditVideoSubjects((current) =>
      current.includes(code) ? current.filter((value) => value !== code) : [...current, code],
    )
  }

  function toggleEditOccurrence(subjectCode, behaviorName) {
    const key = `${subjectCode}::${behaviorName}`
    setEditVideoOccurrenceMap((current) => ({ ...current, [key]: !current[key] }))
  }

  if (!ready) {
    return (
      <div className="app-shell loading">
        <div className="card">
          <h1>Lab School Video Behavior Database</h1>
          <p className="muted">Loading data…</p>
        </div>
      </div>
    )
  }

  return (
    <div className="app-shell">
      <header className="brand-bar">
        <button type="button" className="brand-logo-button" onClick={() => setView(views.home)}>
          <img
            className="brand-logo"
            src={`${import.meta.env.BASE_URL}tcfdlogo.png`}
            alt="The Center For Discovery"
          />
        </button>
      </header>

      {status.conflict && (
        <div className="card" style={{ borderColor: '#c0392b', background: '#fdecea', marginBottom: '16px' }}>
          <span className="error">{status.conflict.text}</span>
          <button type="button" onClick={initialize} disabled={loading} style={{ marginLeft: '12px' }}>
            {loading ? 'Reloading…' : 'Reload latest data'}
          </button>
        </div>
      )}

      <header className="hero card">
        <div>
          <p className="eyebrow">TCFD Lab School</p>
          <h1>Recording Database</h1>
          <p className="muted">What would you like to do?</p>
        </div>
        <div className="hero-actions">
          <button type="button" className="primary" onClick={() => setView(views.intake)}>Add New Recording</button>
          <button type="button" onClick={() => setView(views.data)}>Review Data</button>
          <button type="button" onClick={() => setView(views.review)}>Run Queries</button>
        </div>
      </header>

      {view === views.home && (
        <section className="card">
          {status.home && (
            <div className="status" aria-live="polite" style={{ marginBottom: '16px' }}>
              <span className={status.home.isError ? 'error' : 'success'}>{status.home.text}</span>
              {status.home.isError && (
                <button type="button" onClick={initialize} disabled={loading} style={{ marginLeft: '12px' }}>
                  {loading ? 'Retrying…' : 'Retry'}
                </button>
              )}
            </div>
          )}
          <div className="home-cards">
            <button type="button" className="home-card" onClick={() => setView(views.intake)}>
              <span className="home-card-icon">🎥</span>
              <span className="home-card-title">Add New Recording</span>
              <span className="home-card-sub">Log a video and mark behaviors</span>
            </button>
            <button type="button" className="home-card" onClick={() => setView(views.data)}>
              <span className="home-card-icon">📄</span>
              <span className="home-card-title">Review Data</span>
              <span className="home-card-sub">Browse all recordings in one sortable table</span>
            </button>
            <button type="button" className="home-card" onClick={() => setView(views.review)}>
              <span className="home-card-icon">🔎</span>
              <span className="home-card-title">Run Queries</span>
              <span className="home-card-sub">Search by subject, behavior, or upload status</span>
            </button>
          </div>
          <div className="home-admin-link">
            <button type="button" className="admin-quiet-link" onClick={() => setView(views.admin)}>⚙ Admin</button>
          </div>
        </section>
      )}

      {view === views.intake && (
        <section className="card">
          <div className="section-heading">
            <div>
              <h2>Video Recording</h2>
                <p className="muted">Step {intakeStep} of 3</p>
            </div>
            <button type="button" onClick={resetIntake}>Reset</button>
          </div>

          {intakeStep === 1 && (
            <div className="form-grid">
              <label>
                Start time
                <input type="datetime-local" step="1" value={recordStartTime} onChange={(event) => setRecordStartTime(event.target.value)} />
              </label>
              <label>
                Duration (minutes)
                <input type="number" min="1" step="1" value={durationMinutes} onChange={(event) => setDurationMinutes(event.target.value)} />
              </label>
              <label className="full">
                Notes
                <textarea rows="4" value={videoNotes} onChange={(event) => setVideoNotes(event.target.value)} placeholder="Optional recording notes" />
              </label>
            </div>
          )}

          {intakeStep === 2 && (
            <div>
              <p className="muted">Choose who was present in the video.</p>
              <div className="chip-grid">
                {activeSubjects.map((subject) => (
                  <label key={subject.subjectCode} className={`chip ${selectedSubjects.includes(subject.subjectCode) ? 'selected' : ''}`}>
                    <input
                      type="checkbox"
                      checked={selectedSubjects.includes(subject.subjectCode)}
                      onChange={() => toggleSubject(subject.subjectCode)}
                    />
                    <span>{subject.subjectCode}</span>
                  </label>
                ))}
              </div>
            </div>
          )}

          {intakeStep === 3 && (
            <div className="stack">
              {selectedActiveSubjects.map((subjectCode) => (
                <div key={subjectCode} className="card inset">
                  <h3>{subjectCode}</h3>
                  <div className="chip-grid">
                    {activeBehaviors.map((behavior) => {
                      const checked = Boolean(occurrenceMap[`${subjectCode}::${behavior.name}`])
                      return (
                        <label key={behavior.name} className={`chip ${checked ? 'selected' : ''}`}>
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => toggleOccurrence(subjectCode, behavior.name)}
                          />
                          <span>{behavior.name}</span>
                        </label>
                      )
                    })}
                  </div>
                </div>
              ))}
            </div>
          )}

          <div className="status" aria-live="polite">
            {status.recording && <span className={status.recording.isError ? 'error' : 'success'}>{status.recording.text}</span>}
          </div>

          <div className="button-row">
            {intakeStep > 1 ? (
              <button type="button" onClick={() => setIntakeStep(intakeStep - 1)}>Back</button>
            ) : <span />}
            {intakeStep < 3 ? (
              <button type="button" className="primary" onClick={() => setIntakeStep(intakeStep + 1)}>Next</button>
            ) : (
              <div className="button-row">
                <button type="button" onClick={() => handleSave(true)}>Save and Add Another</button>
                <button type="button" className="primary" onClick={() => handleSave(false)}>Save</button>
              </div>
            )}
          </div>
        </section>
      )}

      {view === views.review && (
        <section className="stack">
          <div className="card">
            <div className="section-heading">
              <div>
                <h2>Queries</h2>
                <p className="muted">Fast reports for reviews and upload tracking.</p>
              </div>
            </div>
            <div className="query-grid">
              <article className="card inset">
                <h3>Subject Profile</h3>
                <p className="muted">Look up a subject's historical profile, or add a new subject by ID.</p>
                <div className="section-heading tight">
                  <select value={profileSubjectCode} onChange={(event) => { setProfileSubjectCode(event.target.value); setProfileLoaded(false) }}>
                    <option value="">Select a subject…</option>
                    {profileSubjectOptions.map((subject) => (
                      <option key={subject.subjectCode} value={subject.subjectCode}>
                        {subject.subjectCode}{subject.isActive ? '' : ' (inactive)'}
                      </option>
                    ))}
                  </select>
                  <button type="button" onClick={loadSubjectProfile}>Load Profile</button>
                </div>

                {profileLoaded && (
                  <div className="stack">
                    <p className="muted">
                      {profileVideoCount} video(s) logged, {profileOccurrenceCount} behavior occurrence(s) recorded.
                    </p>
                    <div className="form-grid">
                      <label>
                        Person ID
                        <input value={profilePersonId} onChange={(event) => setProfilePersonId(event.target.value)} placeholder="e.g. P55943" />
                      </label>
                      <label>
                        Date of birth
                        <input type="date" value={profileDob} onChange={(event) => setProfileDob(event.target.value)} />
                      </label>
                      <label>
                        Lab School start date
                        <input type="date" value={profileStartDate} onChange={(event) => setProfileStartDate(event.target.value)} />
                      </label>
                      <label>
                        Lab School end date
                        <input type="date" value={profileEndDate} onChange={(event) => setProfileEndDate(event.target.value)} />
                      </label>
                    </div>
                    <p className="muted">Target behaviors (validated list)</p>
                    <div className="chip-grid">
                      {targetBehaviorOptions.map((name) => {
                        const checked = profileTargetBehaviors.includes(name)
                        return (
                          <label key={name} className={`chip ${checked ? 'selected' : ''}`}>
                            <input type="checkbox" checked={checked} onChange={() => toggleProfileTargetBehavior(name)} />
                            <span>{name}</span>
                          </label>
                        )
                      })}
                    </div>
                    <div className="button-row">
                      <button type="button" className="primary" onClick={saveSubjectProfile}>Save Profile</button>
                    </div>
                  </div>
                )}

                <div className="status" aria-live="polite">
                  {status.profile && <span className={status.profile.isError ? 'error' : 'success'}>{status.profile.text}</span>}
                </div>

                <div className="form-grid single">
                  <label>
                    Add a new subject
                    <input value={newProfileSubjectCode} onChange={(event) => setNewProfileSubjectCode(event.target.value)} placeholder="S24" />
                  </label>
                </div>
                <button type="button" onClick={addProfileSubject}>Add Subject</button>
              </article>

              <article className="card inset">
                <h3>Q1 — All Videos</h3>
                <Table
                  columns={[
                    { key: 'id', label: 'ID', sortable: true, render: (row) => esc(row.id) },
                    { key: 'recordStartTime', label: 'Start', sortable: true, render: (row) => esc(fmtDate(row.recordStartTime)) },
                    { key: 'durationSeconds', label: 'Duration', sortable: true, render: (row) => esc(fmtDuration(row.durationSeconds)) },
                    { key: 'subjectCount', label: 'Subjects', sortable: true, render: (row) => esc(row.subjectCodes.length) },
                    { key: 'uploadedToSharePoint', label: 'Uploaded', sortable: true, render: (row) => (row.uploadedToSharePoint ? 'Yes' : 'No') },
                    { key: 'folder', label: 'Folder', sortable: false, render: (row) => <a href={sharePointFolderUrl(row.recordStartTime)} target="_blank" rel="noopener">View Folder</a> },
                  ]}
                  rows={q1Rows}
                  sortState={q1Sort}
                  onSort={(key) => setQ1Sort((current) => ({ sortCol: key, sortDir: current.sortCol === key ? current.sortDir * -1 : 1 }))}
                  emptyMessage="No videos yet."
                />
              </article>

              <article className="card inset">
                <h3>Q2 — Behavior Occurrences</h3>
                <Table
                  columns={[
                    { key: 'videoId', label: 'Video', sortable: true, render: (row) => esc(row.videoId) },
                    { key: 'recordStartTime', label: 'Start', sortable: true, render: (row) => esc(fmtDate(row.recordStartTime)) },
                    { key: 'subjectCode', label: 'Subject', sortable: true, render: (row) => esc(row.subjectCode) },
                    { key: 'behaviorTypeName', label: 'Behavior', sortable: true, render: (row) => esc(row.behaviorTypeName) },
                    { key: 'notes', label: 'Notes', sortable: true, render: (row) => esc(row.notes || '—') },
                  ]}
                  rows={q2Rows}
                  sortState={q2Sort}
                  onSort={(key) => setQ2Sort((current) => ({ sortCol: key, sortDir: current.sortCol === key ? current.sortDir * -1 : 1 }))}
                  emptyMessage="No behavior occurrences yet."
                />
              </article>

              <article className="card inset">
                <div className="section-heading tight">
                  <h3>Q3 — Videos for Subject</h3>
                  <select value={q3Subject} onChange={(event) => setQ3Subject(event.target.value)}>
                    {q3SubjectOptions.map((subjectCode) => <option key={subjectCode} value={subjectCode}>{subjectCode}</option>)}
                  </select>
                </div>
                <Table
                  columns={[
                    { key: 'id', label: 'ID', sortable: true, render: (row) => esc(row.id) },
                    { key: 'recordStartTime', label: 'Start', sortable: true, render: (row) => esc(fmtDate(row.recordStartTime)) },
                    { key: 'durationSeconds', label: 'Duration', sortable: true, render: (row) => esc(fmtDuration(row.durationSeconds)) },
                    { key: 'folder', label: 'Folder', sortable: false, render: (row) => <a href={sharePointFolderUrl(row.recordStartTime)} target="_blank" rel="noopener">View Folder</a> },
                  ]}
                  rows={q3Rows}
                  sortState={dataSort}
                  onSort={(key) => setDataSort((current) => ({ sortCol: key, sortDir: current.sortCol === key ? current.sortDir * -1 : 1 }))}
                  emptyMessage="No videos for this subject."
                />
              </article>

              <article className="card inset">
                <div className="section-heading tight">
                  <h3>Q4 — Videos with Behavior</h3>
                  <select value={q4Behavior} onChange={(event) => setQ4Behavior(event.target.value)}>
                    {activeBehaviors.map((behavior) => <option key={behavior.name} value={behavior.name}>{behavior.name}</option>)}
                  </select>
                </div>
                <Table
                  columns={[
                    { key: 'id', label: 'ID', sortable: true, render: (row) => esc(row.id) },
                    { key: 'recordStartTime', label: 'Start', sortable: true, render: (row) => esc(fmtDate(row.recordStartTime)) },
                    { key: 'durationSeconds', label: 'Duration', sortable: true, render: (row) => esc(fmtDuration(row.durationSeconds)) },
                    { key: 'folder', label: 'Folder', sortable: false, render: (row) => <a href={sharePointFolderUrl(row.recordStartTime)} target="_blank" rel="noopener">View Folder</a> },
                  ]}
                  rows={q4Rows}
                  sortState={dataSort}
                  onSort={(key) => setDataSort((current) => ({ sortCol: key, sortDir: current.sortCol === key ? current.sortDir * -1 : 1 }))}
                  emptyMessage="No videos with that behavior."
                />
              </article>

              <article className="card inset">
                <div className="section-heading tight">
                  <h3>Q5 — Pending Upload</h3>
                  <span className="muted">{q5Rows.length} rows</span>
                </div>
                <Table
                  columns={[
                    { key: 'id', label: 'ID', sortable: true, render: (row) => esc(row.id) },
                    { key: 'recordStartTime', label: 'Start', sortable: true, render: (row) => esc(fmtDate(row.recordStartTime)) },
                    { key: 'uploadedToSharePoint', label: 'Uploaded', sortable: true, render: (row) => (
                      <label className="toggle">
                        <input type="checkbox" checked={row.uploadedToSharePoint} onChange={(event) => toggleUpload(row.id, event.target.checked)} />
                        <span>Uploaded to SharePoint</span>
                      </label>
                    ) },
                    { key: 'folder', label: 'Folder', sortable: false, render: (row) => <a href={sharePointFolderUrl(row.recordStartTime)} target="_blank" rel="noopener">View Folder</a> },
                  ]}
                  rows={q5Rows}
                  sortState={q5Sort}
                  onSort={(key) => setQ5Sort((current) => ({ sortCol: key, sortDir: current.sortCol === key ? current.sortDir * -1 : 1 }))}
                  emptyMessage="Nothing pending upload."
                />
              </article>

              <article className="card inset">
                <div className="section-heading tight">
                  <h3>Q6 — Already Uploaded</h3>
                  <span className="muted">{q6Rows.length} rows</span>
                </div>
                <Table
                  columns={[
                    { key: 'id', label: 'ID', sortable: true, render: (row) => esc(row.id) },
                    { key: 'recordStartTime', label: 'Start', sortable: true, render: (row) => esc(fmtDate(row.recordStartTime)) },
                    { key: 'uploadedToSharePoint', label: 'Uploaded', sortable: true, render: (row) => (
                      <label className="toggle">
                        <input type="checkbox" checked={row.uploadedToSharePoint} onChange={(event) => toggleUpload(row.id, event.target.checked)} />
                        <span>Uploaded to SharePoint</span>
                      </label>
                    ) },
                    { key: 'folder', label: 'Folder', sortable: false, render: (row) => <a href={sharePointFolderUrl(row.recordStartTime)} target="_blank" rel="noopener">View Folder</a> },
                  ]}
                  rows={q6Rows}
                  sortState={q6Sort}
                  onSort={(key) => setQ6Sort((current) => ({ sortCol: key, sortDir: current.sortCol === key ? current.sortDir * -1 : 1 }))}
                  emptyMessage="No uploaded videos yet."
                />
              </article>
            </div>
            <div className="status" aria-live="polite">
              {status.review && <span className={status.review.isError ? 'error' : 'success'}>{status.review.text}</span>}
            </div>
          </div>
        </section>
      )}

      {view === views.data && (
        <section className="card">
          <div className="section-heading">
            <div>
              <h2>Data</h2>
              <p className="muted">Search and sort all recorded videos.</p>
            </div>
            <input
              type="search"
              value={dataSearch}
              onChange={(event) => setDataSearch(event.target.value)}
              placeholder="Search id, notes, subject, behavior"
            />
          </div>
          <Table
            columns={[
              { key: 'id', label: 'ID', sortable: true, render: (row) => esc(row.id) },
              { key: 'recordStartTime', label: 'Start Time', sortable: true, render: (row) => esc(fmtDate(row.recordStartTime)) },
              { key: 'durationSeconds', label: 'Duration', sortable: true, render: (row) => esc(fmtDuration(row.durationSeconds)) },
              { key: 'subjects', label: 'Subjects', sortable: false, render: (row) => esc(row.subjectCodes.join(', ')) },
              { key: 'behaviors', label: 'Behaviors', sortable: false, render: (row) => esc(row.occurrences.map((occurrence) => occurrence.behaviorTypeName).join(', ') || '—') },
              {
                key: 'uploadedToSharePoint',
                label: 'Uploaded',
                sortable: true,
                render: (row) => (
                  <span className={row.uploadedToSharePoint ? 'uploaded-yes' : 'uploaded-no'}>
                    {row.uploadedToSharePoint ? 'Yes' : 'No'}
                  </span>
                ),
              },
              { key: 'folder', label: 'Folder', sortable: false, render: (row) => <a href={sharePointFolderUrl(row.recordStartTime)} target="_blank" rel="noopener">View Folder</a> },
            ]}
            rows={dataRows}
            sortState={dataSort}
            onSort={(key) => setDataSort((current) => ({ sortCol: key, sortDir: current.sortCol === key ? current.sortDir * -1 : 1 }))}
            emptyMessage="No matching videos found."
          />
          <div className="status" aria-live="polite">
            {status.data && <span className={status.data.isError ? 'error' : 'success'}>{status.data.text}</span>}
          </div>
        </section>
      )}

      {view === views.admin && (
        <div className="stack">
        <section className="grid two-up">
          <article className="card">
            <h2>Subjects</h2>
            <div className="form-grid single">
              <label>
                Subject code
                <input value={newSubjectCode} onChange={(event) => setNewSubjectCode(event.target.value)} placeholder="CSH04" />
              </label>
            </div>
            <button type="button" className="primary" onClick={addSubject}>Add Subject</button>
            <div className="list">
              {subjects.map((subject) => (
                <label key={subject.subjectCode} className="list-item">
                  <span>
                    <strong>{subject.subjectCode}</strong>
                  </span>
                  <input
                    type="checkbox"
                    checked={subject.isActive}
                    onChange={() => setSubjects((current) => current.map((entry) => entry.id === subject.id ? { ...entry, isActive: !entry.isActive } : entry))}
                  />
                </label>
              ))}
            </div>
          </article>

          <article className="card">
            <h2>Behavior Types</h2>
            <div className="form-grid single">
              <label>
                New behavior
                <input value={newBehaviorName} onChange={(event) => setNewBehaviorName(event.target.value)} placeholder="Peer aggression" />
              </label>
            </div>
            <button type="button" className="primary" onClick={addBehavior}>Add Behavior</button>
            <div className="list">
              {behaviors.map((behavior) => (
                <label key={behavior.name} className="list-item">
                  <span><strong>{behavior.name}</strong></span>
                  <input
                    type="checkbox"
                    checked={behavior.isActive}
                    onChange={() => setBehaviors((current) => current.map((entry) => entry.id === behavior.id ? { ...entry, isActive: !entry.isActive } : entry))}
                  />
                </label>
              ))}
            </div>
          </article>
        </section>

        <article className="card">
          <h2>Data Safety</h2>
          <p className="muted">Primary storage is Supabase (shared across devices). Export/import is an extra backup tool.</p>
          <div className="button-row">
            <button type="button" onClick={exportAllData}>Export Backup (JSON)</button>
            <button type="button" className="secondary" onClick={openImportDialog}>Import Backup</button>
          </div>
          <input
            ref={importFileInputRef}
            type="file"
            accept="application/json,.json"
            onChange={importAllData}
            style={{ display: 'none' }}
          />
        </article>

        <article className="card">
          <h2>Log Entries</h2>
          <p className="muted">Edit or delete any previously saved recording log.</p>

          <EditVideoPanel
            editingVideoId={editingVideoId}
            videos={videos}
            editVideoStart={editVideoStart}
            setEditVideoStart={setEditVideoStart}
            editVideoDuration={editVideoDuration}
            setEditVideoDuration={setEditVideoDuration}
            editVideoNotes={editVideoNotes}
            setEditVideoNotes={setEditVideoNotes}
            editVideoSubjects={editVideoSubjects}
            subjects={subjects}
            toggleEditSubject={toggleEditSubject}
            behaviors={behaviors}
            editVideoOccurrenceMap={editVideoOccurrenceMap}
            toggleEditOccurrence={toggleEditOccurrence}
            setVideos={setVideos}
            cancelEditVideo={cancelEditVideo}
            saveEditVideo={saveEditVideo}
          />

          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  {[{key: 'id', label: 'ID'}, {key: 'recordStartTime', label: 'Start'}, {key: 'durationSeconds', label: 'Duration'}, {key: 'subjects', label: 'Subjects'}, {key: 'behaviors', label: 'Behaviors'}, {key: 'uploadedToSharePoint', label: 'Uploaded'}].map((col) => (
                    <th key={col.key}>
                      {col.key !== 'subjects' && col.key !== 'behaviors' ? (
                        <button className="th-button" type="button" onClick={() => setAdminLogSort((current) => ({ sortCol: col.key, sortDir: current.sortCol === col.key ? current.sortDir * -1 : 1 }))}>
                          {col.label}{adminLogSort.sortCol === col.key ? (adminLogSort.sortDir === 1 ? ' ▲' : ' ▼') : ''}
                        </button>
                      ) : col.label}
                    </th>
                  ))}
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {adminLogRows.length === 0 ? (
                  <tr><td colSpan={7} className="muted center">No log entries yet.</td></tr>
                ) : (
                  adminLogRows.map((row) => (
                    <tr key={row.__key} className={editingVideoId === row.id ? 'editing-row' : ''}>
                      <td>{row.id}</td>
                      <td>{fmtDate(row.recordStartTime)}</td>
                      <td>{fmtDuration(row.durationSeconds)}</td>
                      <td>{row.subjectCodes.join(', ')}</td>
                      <td>{row.occurrences.map((o) => o.behaviorTypeName).join(', ') || '—'}</td>
                      <td>
                        <label style={{display: 'flex', alignItems: 'center', gap: '4px'}}>
                          <input type="checkbox" checked={row.uploadedToSharePoint} onChange={(event) => setVideos((current) => current.map((video) => video.id === row.id ? { ...video, uploadedToSharePoint: event.target.checked } : video))} />
                          <span>{row.uploadedToSharePoint ? 'Yes' : 'No'}</span>
                        </label>
                      </td>
                      <td>
                        <div className="btn-row-inline">
                          <button type="button" className="btn-sm" onClick={() => startEditVideo(row)} disabled={editingVideoId !== null}>Edit</button>
                          <button type="button" className="btn-sm btn-danger" onClick={() => deleteVideo(row.id)} disabled={editingVideoId !== null}>Delete</button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>

          <div className="status" aria-live="polite">
            {status.admin && <span className={status.admin.isError ? 'error' : 'success'}>{status.admin.text}</span>}
          </div>
        </article>
        </div>
      )}
    </div>
  )
}

export default App
