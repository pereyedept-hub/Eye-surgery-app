import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'
import { api, isApiConfigured } from './api'
import './App.css'

/* =====================================================
   DATE HELPERS
   (Excel can hand us a Date object, a serial number, or a
   plain string in DD/MM/YYYY or YYYY-MM-DD form — normalise
   everything to YYYY-MM-DD so dates can be compared safely)
   ===================================================== */

function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569)
  const utcValue = utcDays * 86400
  return new Date(utcValue * 1000)
}

function formatDateYMD(date) {
  const y = date.getFullYear()
  const m = String(date.getMonth() + 1).padStart(2, '0')
  const d = String(date.getDate()).padStart(2, '0')
  return `${y}-${m}-${d}`
}

function normalizeDate(value) {
  if (value === null || value === undefined || value === '') return null

  if (value instanceof Date) {
    return formatDateYMD(value)
  }

  if (typeof value === 'number') {
    const d = excelSerialToDate(value)
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`
  }

  const str = String(value).trim()
  if (!str) return null

  // ISO with a time part (e.g. 2026-08-14T18:30:00.000Z): convert to the local date
  if (/^\d{4}-\d{2}-\d{2}T/.test(str)) {
    const parsedIso = new Date(str)
    if (!isNaN(parsedIso.getTime())) {
      return formatDateYMD(parsedIso)
    }
  }

  // ISO: YYYY-MM-DD
  let match = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/)
  if (match) {
    const [, y, m, d] = match
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`
  }

  // DD/MM/YYYY or DD-MM-YYYY
  match = str.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/)
  if (match) {
    const [, d, m, y] = match
    return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`
  }

  // Fallback to native parsing
  const parsed = new Date(str)
  if (!isNaN(parsed.getTime())) {
    return formatDateYMD(parsed)
  }

  return null
}

function formatDateDisplay(ymd) {
  if (!ymd) return ''
  const [y, m, d] = ymd.split('-')
  return `${d}/${m}/${y}`
}

function genId() {
  return `p_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`
}

/* =====================================================
   RECORD HELPERS
   ===================================================== */

function getValue(record, columnName) {
  const key = Object.keys(record).find(
    (key) => key.trim().toLowerCase() === columnName.trim().toLowerCase()
  )

  return key ? record[key] : ''
}

// Finds the admission date for one eye. Tries the exact header first, then
// any header that mentions "admission" and the eye (RE/LE/right/left).
function getAdmissionDate(record, eye) {
  const exact = getValue(record, `${eye} DATE OF ADMISSION`)
  if (exact !== '' && exact !== undefined && exact !== null) return exact

  const words = eye === 'RE' ? ['re ', 're_', '[re', 'right'] : ['le ', 'le_', '[le', 'left']
  const key = Object.keys(record).find((k) => {
    const lower = k.trim().toLowerCase() + ' '
    return (
      lower.includes('admission') &&
      words.some((w) => lower.startsWith(w) || lower.includes(w))
    )
  })

  return key ? record[key] : ''
}

function getPatientId(record) {
  return (
    getValue(record, 'UMID Number') ||
    getValue(record, 'CR Number') ||
    getValue(record, 'Name')
  )
}

function App() {
  const [records, setRecords] = useState([])
  const [recordsLoading, setRecordsLoading] = useState(true)
  const [selectedRecord, setSelectedRecord] = useState(null)

  const [searchTerm, setSearchTerm] = useState('')
  const [showTheatreList, setShowTheatreList] = useState(false)
  const [showDateModal, setShowDateModal] = useState(false)
  const [dateInput, setDateInput] = useState('')
  const [theatreListDate, setTheatreListDate] = useState('')
  const [theatreList, setTheatreList] = useState([])
  const [completedRecords, setCompletedRecords] = useState({})
  const [apiNotice, setApiNotice] = useState(null)
  const [syncing, setSyncing] = useState(false)

  // Load the shared records + completed-checklist status from the
  // Google Sheet backend as soon as the app opens, so everyone sees
  // the same data rather than whatever's cached in their own browser.
  useEffect(() => {
    if (!isApiConfigured()) {
      setApiNotice(
        'This app isn\u2019t connected to the shared Google Sheet yet. Data entered now will not be saved or shared until API_URL is set in api.js.'
      )
      setRecordsLoading(false)
      return
    }

    let cancelled = false

    Promise.all([api.getRecords(), api.getCompletedRecords()])
      .then(([sharedRecords, completed]) => {
        if (cancelled) return
        setRecords(sharedRecords)
        setCompletedRecords(completed)
      })
      .catch((err) => {
        if (cancelled) return
        setApiNotice(`Couldn\u2019t load shared data: ${err.message}`)
      })
      .finally(() => {
        if (!cancelled) setRecordsLoading(false)
      })

    return () => {
      cancelled = true
    }
  }, [])

  const createTheatreList = async (admissionDate) => {
    const target = normalizeDate(admissionDate)
    setShowDateModal(false)
    setTheatreListDate(admissionDate)

    // Check whether a colleague has already generated (and possibly
    // edited) the list for this date before building a fresh one.
    if (isApiConfigured()) {
      try {
        const existing = await api.getTheatreList(target)
        if (existing && existing.length > 0) {
          setTheatreList(existing)
          setShowTheatreList(true)
          return
        }
      } catch (err) {
        alert(`Couldn\u2019t check the shared theatre list, generating locally: ${err.message}`)
      }
    }

    const matchingRecords = records.filter((record) => {
      const reDate = normalizeDate(getAdmissionDate(record, 'RE'))
      const leDate = normalizeDate(getAdmissionDate(record, 'LE'))

      return (reDate && reDate === target) || (leDate && leDate === target)
    })

    const list = matchingRecords.map((record) => ({
      id: genId(),
      name: getValue(record, 'Name'),
      umid: getValue(record, 'UMID Number'),
      crNo: getValue(record, 'CR Number'),
      age: getValue(record, 'Age'),
      sex: getValue(record, 'Sex'),
      diagnosis: getValue(record, 'Diagnosis'),
      reIOL: getValue(record, 'REIOL Power'),
      leIOL: getValue(record, 'LEIOL Power'),
      eyeToBeOperated: '',
      surgery: ''
    }))

    if (list.length === 0) {
      const sample = records[0] || {}
      const admissionHeaders = Object.keys(sample).filter((k) =>
        k.toLowerCase().includes('admission')
      )
      const found = admissionHeaders.length
        ? admissionHeaders
            .map((k) => `"${k}" (e.g. ${String(sample[k])})`)
            .join(', ')
        : 'none \u2014 no column with "admission" in its name'

      alert(
        `No patients have an admission date of ${formatDateDisplay(target)}.\n\n` +
          `Admission columns found in your file: ${found}\n\n` +
          `Check the date, or that these columns hold the dates you expect. ` +
          `You can still add patients manually.`
      )
    }

    setTheatreList(list)
    setShowTheatreList(true)
  }

  const openDateModal = () => {
    setDateInput('')
    setShowDateModal(true)
  }

  const filteredRecords = records.filter((record) => {
    const search = searchTerm.trim().toLowerCase()

    if (!search) {
      return true
    }

    const name = String(getValue(record, 'Name') || '').toLowerCase()
    const umid = String(getValue(record, 'UMID Number') || '').toLowerCase()
    const crNumber = String(getValue(record, 'CR Number') || '').toLowerCase()

    return (
      name.includes(search) ||
      umid.includes(search) ||
      crNumber.includes(search)
    )
  })

  const handleFileUpload = (event) => {
    const file = event.target.files[0]

    if (!file) return

    const reader = new FileReader()

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result)

      const workbook = XLSX.read(data, {
        type: 'array'
      })

      const worksheet = workbook.Sheets[workbook.SheetNames[0]]
      const jsonData = XLSX.utils.sheet_to_json(worksheet)

      setRecords(jsonData)
      setSelectedRecord(null)

      if (isApiConfigured()) {
        api
          .uploadRecords(jsonData)
          .then(() => {
            setApiNotice(null)
          })
          .catch((err) => {
            alert(
              `The records loaded here, but couldn\u2019t be pushed to the shared sheet yet: ${err.message}`
            )
          })
      }
    }

    reader.readAsArrayBuffer(file)
  }

  const handleSyncFromSource = async () => {
    if (!isApiConfigured()) {
      alert('This app isn\u2019t connected to the shared Google Sheet yet.')
      return
    }

    setSyncing(true)

    try {
      await api.syncFromSource()
      const freshRecords = await api.getRecords()
      setRecords(freshRecords)
      setSelectedRecord(null)
      alert('Records refreshed from your Google Sheet.')
    } catch (err) {
      alert(`Couldn\u2019t sync from the Google Sheet: ${err.message}`)
    } finally {
      setSyncing(false)
    }
  }

  const pendingCount = records.filter(
    (record) => !completedRecords[getPatientId(record)]
  ).length

  const completedCount = records.length - pendingCount

  return (
    <div className="app">

      <header className="header">
        <h1>Checklist Management System</h1>
        <p>Manage your patient checklists and records</p>
      </header>

      {apiNotice && (
        <div className="api-notice">
          {apiNotice}
        </div>
      )}

      <main className="dashboard">

        {/* Summary cards */}

        <div className="card">
          <h2>Records</h2>
          <p className="number">{records.length}</p>
          <p>Total records</p>
        </div>

        <div className="card">
          <h2>Pending</h2>
          <p className="number">{records.length > 0 ? pendingCount : 0}</p>
          <p>Pending checklists</p>
        </div>

        <div className="card">
          <h2>Completed</h2>
          <p className="number">{records.length > 0 ? completedCount : 0}</p>
          <p>Completed checklists</p>
        </div>

        {/* Upload */}

        <div className="actions">

          <label className="upload-button">
            Upload Excel

            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileUpload}
              hidden
            />
          </label>

          <button onClick={handleSyncFromSource} disabled={syncing}>
            {syncing ? 'Syncing…' : 'Sync from Google Sheet'}
          </button>

        </div>

        {recordsLoading && (
          <p className="empty-note">Loading shared records…</p>
        )}

        {/* Records table */}
        {records.length > 0 && !selectedRecord && (

          <div className="records-table">
            <div className="theatre-button-row">
              <button onClick={openDateModal}>
                Create Theatre List
              </button>
            </div>

            <h2>Imported Records</h2>

            <div className="search-box">
              <input
                type="text"
                placeholder="Search by Name, UMID or CR Number..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>

            <div className="table-container">

              <table>

                <thead>

                  <tr>
                    <th>Name</th>
                    <th>Age</th>
                    <th>Sex</th>
                    <th>UMID Number</th>
                    <th>Diagnosis</th>
                    <th>Status</th>
                    <th>Action</th>
                  </tr>

                </thead>

                <tbody>

                  {filteredRecords.map((record, index) => (

                    <tr key={index}>

                      <td>{getValue(record, 'Name')}</td>
                      <td>{getValue(record, 'Age')}</td>
                      <td>{getValue(record, 'Sex')}</td>
                      <td>{getValue(record, 'UMID Number')}</td>
                      <td>{getValue(record, 'Diagnosis')}</td>

                      <td>
                        {completedRecords[getPatientId(record)]
                          ? 'Completed'
                          : 'Pending'}
                      </td>

                      <td>
                        <button onClick={() => setSelectedRecord(record)}>
                          Open Checklist
                        </button>
                      </td>

                    </tr>

                  ))}

                </tbody>

              </table>

            </div>

          </div>

        )}

        {/* Checklist */}

        {selectedRecord && (

          <Checklist
            record={selectedRecord}
            onBack={() => setSelectedRecord(null)}
            onSaved={(patientId) =>
              setCompletedRecords((prev) => ({ ...prev, [patientId]: true }))
            }
          />

        )}

        {/* Theatre List */}

        {showTheatreList && (
          <TheatreList
            theatreList={theatreList}
            setTheatreList={setTheatreList}
            admissionDate={theatreListDate}
            onBack={() => setShowTheatreList(false)}
          />
        )}

        {/* Admission date prompt */}

        {showDateModal && (
          <AdmissionDateModal
            value={dateInput}
            onChange={setDateInput}
            onCancel={() => setShowDateModal(false)}
            onConfirm={() => {
              if (!dateInput) {
                alert('Please select the date of admission')
                return
              }
              createTheatreList(dateInput)
            }}
          />
        )}

      </main>

    </div>
  )
}


/* =====================================================
   ADMISSION DATE MODAL
   ===================================================== */

function AdmissionDateModal({ value, onChange, onCancel, onConfirm }) {
  return (
    <div className="modal-overlay">

      <div className="modal-box">

        <h2>Date of Admission</h2>

        <p>
          Enter the date of admission to generate the theatre
          list for that date. Patients whose RE or LE date of
          admission matches will be included.
        </p>

        <div className="field">
          <label>Date of Admission</label>
          <input
            type="date"
            value={value}
            onChange={(e) => onChange(e.target.value)}
            autoFocus
          />
        </div>

        <div className="modal-actions">
          <button onClick={onCancel} className="secondary-button">
            Cancel
          </button>

          <button onClick={onConfirm}>
            Generate Theatre List
          </button>
        </div>

      </div>

    </div>
  )
}


/* =====================================================
   CHECKLIST COMPONENT
   ===================================================== */

function Checklist({ record, onBack, onSaved }) {

  const patientId = getPatientId(record)
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)

  const [formData, setFormData] = useState({
    visionRE: getValue(record, 'Vision [Right Eye ]'),
    visionLE: getValue(record, 'Vision [Left Eye ]'),

    iopRE: getValue(record, 'RE IOP'),
    iopLE: getValue(record, 'LE IOP'),

    anteriorRE: getValue(record, 'RE ANTERIOR SEGMENT'),
    anteriorLE: getValue(record, 'LE ANTERIOR SEGMENT'),

    fundusRE: getValue(record, 'RE Fundus'),
    fundusLE: getValue(record, 'LE Fundus'),

    k1RE: getValue(record, 'RE K1'),
    k1LE: getValue(record, 'LE K1'),

    k1AxisRE: getValue(record, 'RE K1 AXIS'),
    k1AxisLE: getValue(record, 'LE K1 AXIS'),

    k2RE: getValue(record, 'RE K2'),
    k2LE: getValue(record, 'LE K2'),

    k2AxisRE: getValue(record, 'RE K2 AXIS'),
    k2AxisLE: getValue(record, 'LE K2 AXIS'),

    axialRE: getValue(record, 'RE Axial Length'),
    axialLE: getValue(record, 'LE Axial Length'),

    acdRE: getValue(record, 'RE ACD'),
    acdLE: getValue(record, 'LE ACD'),

    lensRE: getValue(record, 'RE LENS THICKNESS'),
    lensLE: getValue(record, 'LE LENS THICKNESS'),

    wtwRE: getValue(record, 'RE WTW'),
    wtwLE: getValue(record, 'LE WTW'),

    iolRE: getValue(record, 'REIOL Power'),
    iolLE: getValue(record, 'LEIOL Power'),

    tc: getValue(record, 'TC'),
    dc: getValue(record, 'DC'),
    hb: getValue(record, 'Hb%'),
    fastingSugar: getValue(record, 'Fasting Sugar'),
    postprandialSugar: getValue(record, 'Postprandial Sugar'),
    pt: getValue(record, 'PT'),
    inr: getValue(record, 'INR'),
    urea: getValue(record, 'Urea'),
    creatinine: getValue(record, 'Creatinine'),

    hbsag: getValue(record, 'Serology [HBSAG]'),
    tpha: getValue(record, 'Serology [TPHA]'),
    hiv: getValue(record, 'Serology [HIV]'),
    hcv: getValue(record, 'Serology [HCV]'),

    systemicIllness: getValue(record, 'Systemic Illness'),
    fitness: getValue(record, 'Fitness obtained from'),
    remarks: getValue(record, 'Remarks'),
    scan: getValue(record, 'A SCAN DONE WITH'),
    formula: getValue(record, 'FORMULA USED'),

    reAdmission: getValue(record, 'RE DATE OF ADMISSION'),
    leAdmission: getValue(record, 'LE DATE OF ADMISSION')
  })

  useEffect(() => {
    if (!isApiConfigured()) {
      setLoading(false)
      return
    }

    let cancelled = false

    api
      .getChecklist(patientId)
      .then((saved) => {
        if (!cancelled && saved) {
          setFormData(saved)
        }
      })
      .catch(() => {
        // no saved checklist yet, or the request failed — keep defaults
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })

    return () => {
      cancelled = true
    }
  }, [patientId])

  const updateField = (field, value) => {
    setFormData({
      ...formData,
      [field]: value
    })
  }

  const handleSave = async () => {
    if (!isApiConfigured()) {
      alert('This app isn\u2019t connected to the shared Google Sheet yet, so this checklist can\u2019t be saved.')
      return
    }

    setSaving(true)

    try {
      await api.saveChecklist(patientId, formData)
      onSaved(patientId)
      alert('Checklist saved successfully!')
    } catch (err) {
      alert(`Couldn\u2019t save the checklist: ${err.message}`)
    } finally {
      setSaving(false)
    }
  }

  return (

    <div className="checklist">

      <div className="checklist-header">

        <button onClick={onBack}>
          ← Back to Records
        </button>

        <div>
          <button onClick={handleSave} disabled={saving}>
            {saving ? 'Saving\u2026' : 'Save Checklist'}
          </button>

          <button onClick={() => window.print()}>
            Print Checklist
          </button>
        </div>

      </div>

      {loading && (
        <p className="empty-note">Loading saved checklist…</p>
      )}

      <div className="checklist-title">

        <h1>
          SOUTHERN RAILWAY HEADQUARTERS HOSPITAL
        </h1>

        <h2>
          EYE DEPARTMENT, PERAMBUR
        </h2>

        <h2>
          CHECKLIST SHEET
        </h2>

      </div>


      {/* Patient details */}

      <section>

        <h3>Patient Details</h3>

        <div className="form-grid">

          <div>
            <label>Name</label>
            <input
              value={getValue(record, 'Name')}
              readOnly
            />
          </div>

          <div>
            <label>Age</label>
            <input
              value={getValue(record, 'Age')}
              readOnly
            />
          </div>

          <div>
            <label>Sex</label>
            <input
              value={getValue(record, 'Sex')}
              readOnly
            />
          </div>

          <div>
            <label>UMID Number</label>
            <input
              value={getValue(record, 'UMID Number')}
              readOnly
            />
          </div>

          <div>
            <label>CR Number</label>
            <input
              value={getValue(record, 'CR Number')}
              readOnly
            />
          </div>

          <div>
            <label>Phone Number</label>
            <input
              value={getValue(record, 'Phone number')}
              readOnly
            />
          </div>

          <div className="full-width">
            <label>Diagnosis</label>
            <input
              value={getValue(record, 'Diagnosis')}
              readOnly
            />
          </div>

        </div>

      </section>


      {/* Eye examination */}

      <section>

        <h3>Eye Examination</h3>

        <table className="checklist-table">

          <thead>

            <tr>
              <th>Parameter</th>
              <th>RE</th>
              <th>LE</th>
            </tr>

          </thead>

          <tbody>

            <EditableEyeRow
              label="Vision"
              re={formData.visionRE}
              le={formData.visionLE}
              onChangeRE={(v) => updateField('visionRE', v)}
              onChangeLE={(v) => updateField('visionLE', v)}
            />

            <EditableEyeRow
              label="IOP"
              re={formData.iopRE}
              le={formData.iopLE}
              onChangeRE={(v) => updateField('iopRE', v)}
              onChangeLE={(v) => updateField('iopLE', v)}
            />

            <EditableEyeRow
              label="Anterior Segment"
              re={formData.anteriorRE}
              le={formData.anteriorLE}
              onChangeRE={(v) => updateField('anteriorRE', v)}
              onChangeLE={(v) => updateField('anteriorLE', v)}
            />

            <EditableEyeRow
              label="Fundus"
              re={formData.fundusRE}
              le={formData.fundusLE}
              onChangeRE={(v) => updateField('fundusRE', v)}
              onChangeLE={(v) => updateField('fundusLE', v)}
            />

            <EditableEyeRow
              label="K1"
              re={formData.k1RE}
              le={formData.k1LE}
              onChangeRE={(v) => updateField('k1RE', v)}
              onChangeLE={(v) => updateField('k1LE', v)}
            />

            <EditableEyeRow
              label="K1 Axis"
              re={formData.k1AxisRE}
              le={formData.k1AxisLE}
              onChangeRE={(v) => updateField('k1AxisRE', v)}
              onChangeLE={(v) => updateField('k1AxisLE', v)}
            />

            <EditableEyeRow
              label="K2"
              re={formData.k2RE}
              le={formData.k2LE}
              onChangeRE={(v) => updateField('k2RE', v)}
              onChangeLE={(v) => updateField('k2LE', v)}
            />

            <EditableEyeRow
              label="K2 Axis"
              re={formData.k2AxisRE}
              le={formData.k2AxisLE}
              onChangeRE={(v) => updateField('k2AxisRE', v)}
              onChangeLE={(v) => updateField('k2AxisLE', v)}
            />

            <EditableEyeRow
              label="Axial Length"
              re={formData.axialRE}
              le={formData.axialLE}
              onChangeRE={(v) => updateField('axialRE', v)}
              onChangeLE={(v) => updateField('axialLE', v)}
            />

            <EditableEyeRow
              label="ACD"
              re={formData.acdRE}
              le={formData.acdLE}
              onChangeRE={(v) => updateField('acdRE', v)}
              onChangeLE={(v) => updateField('acdLE', v)}
            />

            <EditableEyeRow
              label="Lens Thickness"
              re={formData.lensRE}
              le={formData.lensLE}
              onChangeRE={(v) => updateField('lensRE', v)}
              onChangeLE={(v) => updateField('lensLE', v)}
            />

            <EditableEyeRow
              label="WTW"
              re={formData.wtwRE}
              le={formData.wtwLE}
              onChangeRE={(v) => updateField('wtwRE', v)}
              onChangeLE={(v) => updateField('wtwLE', v)}
            />

            <EditableEyeRow
              label="IOL Power"
              re={formData.iolRE}
              le={formData.iolLE}
              onChangeRE={(v) => updateField('iolRE', v)}
              onChangeLE={(v) => updateField('iolLE', v)}
            />

          </tbody>

        </table>

      </section>

      {/* Investigations */}

      <section>

        <h3>Investigations</h3>

        <div className="investigation-grid">

          <EditableField
            label="TC"
            value={formData.tc}
            onChange={(v) => updateField('tc', v)}
          />

          <EditableField
            label="DC"
            value={formData.dc}
            onChange={(v) => updateField('dc', v)}
          />

          <EditableField
            label="Hb%"
            value={formData.hb}
            onChange={(v) => updateField('hb', v)}
          />

          <EditableField
            label="Fasting Sugar"
            value={formData.fastingSugar}
            onChange={(v) => updateField('fastingSugar', v)}
          />

          <EditableField
            label="Postprandial Sugar"
            value={formData.postprandialSugar}
            onChange={(v) => updateField('postprandialSugar', v)}
          />

          <EditableField
            label="PT"
            value={formData.pt}
            onChange={(v) => updateField('pt', v)}
          />

          <EditableField
            label="INR"
            value={formData.inr}
            onChange={(v) => updateField('inr', v)}
          />

          <EditableField
            label="Urea"
            value={formData.urea}
            onChange={(v) => updateField('urea', v)}
          />

          <EditableField
            label="Creatinine"
            value={formData.creatinine}
            onChange={(v) => updateField('creatinine', v)}
          />

          <EditableField
            label="HBSAG"
            value={formData.hbsag}
            onChange={(v) => updateField('hbsag', v)}
          />

          <EditableField
            label="TPHA"
            value={formData.tpha}
            onChange={(v) => updateField('tpha', v)}
          />

          <EditableField
            label="HIV"
            value={formData.hiv}
            onChange={(v) => updateField('hiv', v)}
          />

          <EditableField
            label="HCV"
            value={formData.hcv}
            onChange={(v) => updateField('hcv', v)}
          />

        </div>

      </section>


      {/* Other information */}

      <section>

        <h3>Other Information</h3>

        <EditableField
          label="Systemic Illness"
          value={formData.systemicIllness}
          onChange={(v) => updateField('systemicIllness', v)}
        />

        <EditableField
          label="Fitness obtained from"
          value={formData.fitness}
          onChange={(v) => updateField('fitness', v)}
        />

        <EditableField
          label="Remarks"
          value={formData.remarks}
          onChange={(v) => updateField('remarks', v)}
        />

        <EditableField
          label="A SCAN DONE WITH"
          value={formData.scan}
          onChange={(v) => updateField('scan', v)}
        />

        <EditableField
          label="FORMULA USED"
          value={formData.formula}
          onChange={(v) => updateField('formula', v)}
        />

        <div className="form-grid">

          <EditableField
            label="RE Date of Admission"
            value={formData.reAdmission}
            onChange={(v) => updateField('reAdmission', v)}
          />

          <EditableField
            label="LE Date of Admission"
            value={formData.leAdmission}
            onChange={(v) => updateField('leAdmission', v)}
          />

        </div>

      </section>

    </div>
  )
}

function EditableEyeRow({
  label,
  re,
  le,
  onChangeRE,
  onChangeLE
}) {

  return (
    <tr>

      <td>{label}</td>

      <td>
        <input
          value={re || ''}
          onChange={(e) => onChangeRE(e.target.value)}
        />
      </td>

      <td>
        <input
          value={le || ''}
          onChange={(e) => onChangeLE(e.target.value)}
        />
      </td>

    </tr>
  )
}


/* General field */

function EditableField({
  label,
  value,
  onChange
}) {

  return (
    <div className="field">

      <label>{label}</label>

      <input
        value={value || ''}
        onChange={(e) => onChange(e.target.value)}
      />

    </div>
  )
}

function TheatreList({
  theatreList,
  setTheatreList,
  admissionDate,
  onBack
}) {

  const [saving, setSaving] = useState(false)

  const saveTheatreList = async () => {
    if (!isApiConfigured()) {
      alert('This app isn\u2019t connected to the shared Google Sheet yet, so this list can\u2019t be saved.')
      return
    }

    setSaving(true)

    try {
      await api.saveTheatreList(normalizeDate(admissionDate), theatreList)
      alert('Theatre list saved \u2014 everyone will now see this version.')
    } catch (err) {
      alert(`Couldn\u2019t save the theatre list: ${err.message}`)
    } finally {
      setSaving(false)
    }
  }

  const printTheatreList = () => {
    window.print()
  }

  const updateTheatreField = (id, field, value) => {
    setTheatreList(
      theatreList.map((patient) =>
        patient.id === id
          ? { ...patient, [field]: value }
          : patient
      )
    )
  }

  const addPatient = () => {
    setTheatreList([
      ...theatreList,
      {
        id: genId(),
        name: '',
        umid: '',
        crNo: '',
        age: '',
        sex: '',
        diagnosis: '',
        reIOL: '',
        leIOL: '',
        eyeToBeOperated: '',
        surgery: ''
      }
    ])
  }

  const removePatient = (id) => {
    setTheatreList(theatreList.filter((patient) => patient.id !== id))
  }

  return (
    <div className="theatre-list-page">

      <h2>
        Theatre List
        {admissionDate
          ? ` \u2014 ${formatDateDisplay(normalizeDate(admissionDate))}`
          : ''}
      </h2>

      <div>
        <button onClick={saveTheatreList} disabled={saving}>
          {saving ? 'Saving\u2026' : 'Save Theatre List'}
        </button>

        <button onClick={addPatient}>
          Add Patient
        </button>

        <button onClick={printTheatreList}>
          Print Theatre List
        </button>

        <button onClick={onBack}>
          Back
        </button>
      </div>

      {theatreList.length === 0 && (
        <p className="empty-note">
          No patients found for this date yet. Use "Add Patient" to
          add one manually, or go back and pick a different date.
        </p>
      )}

      <div className="table-container">

        <table>

          <thead>
            <tr>
              <th>S.No.</th>
              <th>Patient Name</th>
              <th>UMID</th>
              <th>CR No.</th>
              <th>Age</th>
              <th>Sex</th>
              <th>Diagnosis</th>
              <th>RE IOL Power</th>
              <th>LE IOL Power</th>
              <th>Eye to be Operated</th>
              <th>Surgery</th>
              <th>Remove</th>
            </tr>
          </thead>

          <tbody>

            {theatreList.map((patient, index) => (

              <tr key={patient.id}>

                <td>{index + 1}</td>

                <td>
                  <input
                    type="text"
                    value={patient.name}
                    onChange={(e) =>
                      updateTheatreField(patient.id, 'name', e.target.value)
                    }
                  />
                </td>

                <td>
                  <input
                    type="text"
                    value={patient.umid}
                    onChange={(e) =>
                      updateTheatreField(patient.id, 'umid', e.target.value)
                    }
                  />
                </td>

                <td>
                  <input
                    type="text"
                    value={patient.crNo}
                    onChange={(e) =>
                      updateTheatreField(patient.id, 'crNo', e.target.value)
                    }
                  />
                </td>

                <td>
                  <input
                    type="text"
                    value={patient.age}
                    onChange={(e) =>
                      updateTheatreField(patient.id, 'age', e.target.value)
                    }
                  />
                </td>

                <td>
                  <input
                    type="text"
                    value={patient.sex}
                    onChange={(e) =>
                      updateTheatreField(patient.id, 'sex', e.target.value)
                    }
                  />
                </td>

                <td>
                  <input
                    type="text"
                    value={patient.diagnosis}
                    onChange={(e) =>
                      updateTheatreField(patient.id, 'diagnosis', e.target.value)
                    }
                  />
                </td>

                <td>
                  <input
                    type="text"
                    value={patient.reIOL}
                    onChange={(e) =>
                      updateTheatreField(patient.id, 'reIOL', e.target.value)
                    }
                  />
                </td>

                <td>
                  <input
                    type="text"
                    value={patient.leIOL}
                    onChange={(e) =>
                      updateTheatreField(patient.id, 'leIOL', e.target.value)
                    }
                  />
                </td>

                <td>
                  <input
                    type="text"
                    value={patient.eyeToBeOperated}
                    onChange={(e) =>
                      updateTheatreField(
                        patient.id,
                        'eyeToBeOperated',
                        e.target.value
                      )
                    }
                  />
                </td>

                <td>
                  <input
                    type="text"
                    value={patient.surgery}
                    onChange={(e) =>
                      updateTheatreField(
                        patient.id,
                        'surgery',
                        e.target.value
                      )
                    }
                  />
                </td>

                <td>
                  <button
                    className="secondary-button"
                    onClick={() => removePatient(patient.id)}
                  >
                    ✕
                  </button>
                </td>

              </tr>

            ))}

          </tbody>

        </table>

      </div>

    </div>
  )
}

export default App
