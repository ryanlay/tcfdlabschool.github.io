import { useEffect, useMemo, useState } from 'react'

const STORAGE_KEYS = {
  subjects: 'lbs-react-subjects',
  behaviors: 'lbs-react-behaviors',
  videos: 'lbs-react-videos',
}

const defaultSubjects = [
  { id: 1, subjectCode: 'CSH01', displayName: 'Classroom Student 01', isActive: true },
  { id: 2, subjectCode: 'CSH02', displayName: 'Classroom Student 02', isActive: true },
  { id: 3, subjectCode: 'CSH03', displayName: 'Classroom Student 03', isActive: true },
]

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

function loadJson(key, fallback) {
  try {
    const raw = localStorage.getItem(key)
    return raw ? JSON.parse(raw) : fallback
  } catch {
    return fallback
  }
}

function saveJson(key, value) {
  localStorage.setItem(key, JSON.stringify(value))
}

function localInputValue(date = new Date()) {
  const copy = new Date(date)
  copy.setMinutes(copy.getMinutes() - copy.getTimezoneOffset())
  return copy.toISOString().slice(0, 16)
}

function fmtDate(value) {
  if (!value) return '—'
  return new Date(value).toLocaleString()
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
  const dd = String(date.getDate()).padStart(2, '0')
  return `https://thecenterfordiscovery.sharepoint.com/:f:/r/sites/LabSchool/Shared%20Documents/Videos/${yyyy}/${mm}/${dd}?csf=1&web=1`
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

function App() {
  const [ready, setReady] = useState(false)
  const [view, setView] = useState(views.home)
  const [subjects, setSubjects] = useState(defaultSubjects)
  const [behaviors, setBehaviors] = useState(defaultBehaviors)
  const [videos, setVideos] = useState(defaultVideos)
  const [status, setStatus] = useState({ home: '', intake: '', review: '', data: '', admin: '' })

  const [intakeStep, setIntakeStep] = useState(1)
  const [recordStartTime, setRecordStartTime] = useState(localInputValue())
  const [durationMinutes, setDurationMinutes] = useState('20')
  const [videoNotes, setVideoNotes] = useState('')
  const [selectedSubjects, setSelectedSubjects] = useState(['CSH01', 'CSH02'])
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

  // Log entry editing
  const [editingVideoId, setEditingVideoId] = useState(null)
  const [editVideoStart, setEditVideoStart] = useState('')
  const [editVideoDuration, setEditVideoDuration] = useState('')
  const [editVideoNotes, setEditVideoNotes] = useState('')
  const [editVideoSubjects, setEditVideoSubjects] = useState([])
  const [editVideoOccurrenceMap, setEditVideoOccurrenceMap] = useState({})
  const [adminLogSort, setAdminLogSort] = useState({ sortCol: 'recordStartTime', sortDir: -1 })

  useEffect(() => {
    const storedSubjects = loadJson(STORAGE_KEYS.subjects, null)
    const storedBehaviors = loadJson(STORAGE_KEYS.behaviors, null)
    const storedVideos = loadJson(STORAGE_KEYS.videos, null)

    setSubjects(Array.isArray(storedSubjects) && storedSubjects.length ? storedSubjects : defaultSubjects)
    setBehaviors(Array.isArray(storedBehaviors) && storedBehaviors.length ? storedBehaviors : defaultBehaviors)
    setVideos(Array.isArray(storedVideos) ? storedVideos : defaultVideos)
    setReady(true)
  }, [])

  useEffect(() => {
    if (!ready) return
    saveJson(STORAGE_KEYS.subjects, subjects)
    saveJson(STORAGE_KEYS.behaviors, behaviors)
    saveJson(STORAGE_KEYS.videos, videos)
  }, [ready, subjects, behaviors, videos])

  const activeSubjects = useMemo(() => subjects.filter((subject) => subject.isActive), [subjects])
  const activeBehaviors = useMemo(() => behaviors.filter((behavior) => behavior.isActive), [behaviors])

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
    setDurationMinutes('20')
    setVideoNotes('')
    setSelectedSubjects([])
    setOccurrenceMap({})
    show('intake', '', false)
  }

  function toggleSubject(code) {
    setSelectedSubjects((current) =>
      current.includes(code) ? current.filter((value) => value !== code) : [...current, code],
    )
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
      show('intake', 'Enter a start time, a duration, and at least one subject.', true)
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
    show('intake', `Saved video ${nextVideo.id}.`, false)

    if (addAnother) {
      setRecordStartTime(localInputValue())
      setDurationMinutes('20')
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
      displayName: newSubjectName.trim() || code,
      isActive: true,
    }

    setSubjects((current) => [nextSubject, ...current])
    setNewSubjectCode('')
    setNewSubjectName('')
    show('admin', `Added subject ${code}.`, false)
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
      for (const behavior of activeBehaviors) {
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
        <img
          className="brand-logo"
          src={`${import.meta.env.BASE_URL}tcfdlogo.png`}
          alt="The Center For Discovery"
        />
      </header>

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
          <button type="button" className="secondary" onClick={() => setView(views.admin)}>Admin</button>
        </div>
      </header>

      {view === views.home && (
        <section className="card">
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
              <h2>Video Intake</h2>
              <p className="muted">Step {intakeStep} of 3</p>
            </div>
            <button type="button" onClick={resetIntake}>Reset</button>
          </div>

          {intakeStep === 1 && (
            <div className="form-grid">
              <label>
                Start time
                <input type="datetime-local" value={recordStartTime} onChange={(event) => setRecordStartTime(event.target.value)} />
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
                    <small>{subject.displayName}</small>
                  </label>
                ))}
              </div>
            </div>
          )}

          {intakeStep === 3 && (
            <div className="stack">
              {selectedSubjects.map((subjectCode) => (
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
            {status.intake && <span className={status.intake.isError ? 'error' : 'success'}>{status.intake.text}</span>}
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
                    {activeSubjects.map((subject) => <option key={subject.subjectCode} value={subject.subjectCode}>{subject.subjectCode}</option>)}
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
              { key: 'uploadedToSharePoint', label: 'Uploaded', sortable: true, render: (row) => (row.uploadedToSharePoint ? 'Yes' : 'No') },
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
            <div className="form-grid">
              <label>
                Subject code
                <input value={newSubjectCode} onChange={(event) => setNewSubjectCode(event.target.value)} placeholder="CSH04" />
              </label>
              <label>
                Display name
                <input value={newSubjectName} onChange={(event) => setNewSubjectName(event.target.value)} placeholder="Classroom Student 04" />
              </label>
            </div>
            <button type="button" className="primary" onClick={addSubject}>Add Subject</button>
            <div className="list">
              {subjects.map((subject) => (
                <label key={subject.subjectCode} className="list-item">
                  <span>
                    <strong>{subject.subjectCode}</strong> <span className="muted">{subject.displayName}</span>
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
          <h2>Log Entries</h2>
          <p className="muted">Edit or delete any previously saved recording log.</p>

          {editingVideoId !== null && (() => {
            const editing = videos.find((v) => v.id === editingVideoId)
            return (
              <div className="edit-panel">
                <h3>Editing Video {editingVideoId}</h3>
                <div className="form-grid">
                  <label>
                    Start time
                    <input type="datetime-local" value={editVideoStart} onChange={(event) => setEditVideoStart(event.target.value)} />
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

                <p className="muted" style={{marginTop: '12px', marginBottom: '4px'}}>Subjects present</p>
                <div className="chip-grid">
                  {subjects.map((subject) => (
                    <label key={subject.subjectCode} className={`chip ${editVideoSubjects.includes(subject.subjectCode) ? 'selected' : ''}`}>
                      <input
                        type="checkbox"
                        checked={editVideoSubjects.includes(subject.subjectCode)}
                        onChange={() => toggleEditSubject(subject.subjectCode)}
                      />
                      <span>{subject.subjectCode}</span>
                      <small>{subject.displayName}</small>
                    </label>
                  ))}
                </div>

                {editVideoSubjects.length > 0 && (
                  <>
                    <p className="muted" style={{marginTop: '12px', marginBottom: '4px'}}>Behavior occurrences</p>
                    <div className="stack">
                      {editVideoSubjects.map((subjectCode) => (
                        <div key={subjectCode} className="card inset">
                          <h4 style={{margin: '0 0 8px'}}>{subjectCode}</h4>
                          <div className="chip-grid">
                            {activeBehaviors.map((behavior) => {
                              const checked = Boolean(editVideoOccurrenceMap[`${subjectCode}::${behavior.name}`])
                              return (
                                <label key={behavior.name} className={`chip ${checked ? 'selected' : ''}`}>
                                  <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={() => toggleEditOccurrence(subjectCode, behavior.name)}
                                  />
                                  <span>{behavior.name}</span>
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
          })()}

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
                      <td>{row.uploadedToSharePoint ? 'Yes' : 'No'}</td>
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
