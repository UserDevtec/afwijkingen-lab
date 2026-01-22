import { useMemo, useState } from 'react'
import ExcelJS from 'exceljs'
import * as XLSX from 'xlsx'
import './App.css'

const normalize = (value) => String(value ?? '').trim().toLowerCase()

const getColumnIndex = (headers, name) =>
  headers.findIndex((header) => normalize(header) === normalize(name))

const parseExcelDate = (value) => {
  if (!value && value !== 0) return null
  if (value instanceof Date) return value
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (!parsed) return null
    return new Date(parsed.y, parsed.m - 1, parsed.d)
  }
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (!trimmed) return null
    const parsed = new Date(trimmed)
    if (Number.isNaN(parsed.getTime())) return null
    return parsed
  }
  return null
}

const formatDate = (value) => {
  const date = value instanceof Date ? value : parseExcelDate(value)
  if (!date) return ''
  return date.toLocaleDateString('nl-NL')
}

const getWeekNumber = (date) => {
  const temp = new Date(date.getTime())
  temp.setHours(0, 0, 0, 0)
  temp.setDate(temp.getDate() + 3 - ((temp.getDay() + 6) % 7))
  const week1 = new Date(temp.getFullYear(), 0, 4)
  return (
    1 +
    Math.round(
      ((temp.getTime() - week1.getTime()) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7
    )
  )
}

const buildEmailDraft = (stationLabel) => {
  const now = new Date()
  const weekNumber = String(getWeekNumber(now)).padStart(2, '0')
  const year = String(now.getFullYear())
  const station = stationLabel?.trim() || 'Station'
  const subject = `${station} - Afwijkingen maatregelen - deadlines week ${weekNumber}-${year}`
  const body = [
    'Beste collega s,',
    '',
    "Bij deze het overzicht van de maatregelen horende bij afwijkingen waarvoor de implementatiedatum verstreken is (zie de 'Geplande datum klaar' kolom).",
    "Ik ontvang graag een status update omtrent deze maatregelen. De updates die mij bekend zijn staan onder de 'Opmerkingen' kolom.",
    '',
    'Ik hoop jullie hiermee voldoende te hebben geinformeerd. Bij vragen hoor ik het graag.',
  ].join('\n')
  return { subject, body }
}

const buildTimestamp = () => {
  const now = new Date()
  const pad = (value) => String(value).padStart(2, '0')
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(
    now.getHours()
  )}-${pad(now.getMinutes())}-${pad(now.getSeconds())}`
}

const computeColumnWidths = (headers, rows, metaRows) => {
  const widths = headers.map(() => 10)
  const applyRow = (row) => {
    row.forEach((value, index) => {
      if (index >= widths.length) return
      const length = String(value ?? '').length
      const next = Math.min(Math.max(length + 2, 10), 60)
      widths[index] = Math.max(widths[index], next)
    })
  }
  headers.forEach((value, index) => {
    const length = String(value ?? '').length
    widths[index] = Math.max(widths[index], Math.min(Math.max(length + 2, 10), 60))
  })
  rows.forEach(applyRow)
  metaRows.forEach(applyRow)
  return widths
}

const styleMetaBlock = (sheet) => {
  for (let rowIndex = 1; rowIndex <= 3; rowIndex += 1) {
    const row = sheet.getRow(rowIndex)
    row.height = 18
    const labelCell = row.getCell(1)
    labelCell.font = { bold: true }
    labelCell.alignment = { vertical: 'middle' }
    const valueCell = row.getCell(2)
    valueCell.alignment = { vertical: 'middle' }
  }
}

const styleTableHeader = (sheet, columnCount) => {
  const headerRow = sheet.getRow(4)
  headerRow.height = 20
  for (let colIndex = 1; colIndex <= columnCount; colIndex += 1) {
    const cell = headerRow.getCell(colIndex)
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.alignment = { vertical: 'middle', wrapText: true }
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF630D80' },
    }
  }
}

const styleTableRows = (sheet, rowCount, columnCount) => {
  const lightFill = { argb: 'FFC1E62E' }
  const darkFill = { argb: 'FFBAFF33' }
  for (let rowIndex = 5; rowIndex < 5 + rowCount; rowIndex += 1) {
    const row = sheet.getRow(rowIndex)
    const fillColor = rowIndex % 2 === 0 ? darkFill : lightFill
    for (let colIndex = 1; colIndex <= columnCount; colIndex += 1) {
      const cell = row.getCell(colIndex)
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: fillColor,
      }
    }
  }
}

const setCellValue = (sheet, rowIndex, colIndex, value) => {
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
  if (value === null || value === undefined || value === '') {
    sheet[address] = { t: 's', v: '' }
    return
  }
  if (value instanceof Date) {
    sheet[address] = { t: 'd', v: value }
    return
  }
  if (typeof value === 'number') {
    sheet[address] = { t: 'n', v: value }
    return
  }
  sheet[address] = { t: 's', v: String(value) }
}

const readWorkbook = async (file) => {
  const buffer = await file.arrayBuffer()
  return XLSX.read(buffer, { type: 'array', cellDates: true })
}

function App() {
  const [dashboardFile, setDashboardFile] = useState(null)
  const [databaseFile, setDatabaseFile] = useState(null)
  const [overzichtFile, setOverzichtFile] = useState(null)
  const [station, setStation] = useState('')
  const [logEntries, setLogEntries] = useState([])
  const [busyAction, setBusyAction] = useState('')
  const [achterstalligRows, setAchterstalligRows] = useState([])
  const [conceptRows, setConceptRows] = useState([])
  const [actiehouders, setActiehouders] = useState([])
  const [emailDraft, setEmailDraft] = useState(null)
  const [powerBiReady, setPowerBiReady] = useState(false)
  const [dragTarget, setDragTarget] = useState('')
  const [activePanel, setActivePanel] = useState('results')

  const summaryStats = useMemo(
    () => ({
      achterstallig: achterstalligRows.length,
      concept: conceptRows.length,
      actiehouders: actiehouders.length,
    }),
    [achterstalligRows.length, conceptRows.length, actiehouders.length]
  )

  const addLog = (message, type = 'info') => {
    const timestamp = new Date().toISOString().replace('T', ' ').slice(0, 19)
    setLogEntries((prev) => [...prev.slice(-199), { message, timestamp, type }])
  }

  const handleDashboardUpload = async (file) => {
    setDashboardFile(file)
    if (!file) return
    try {
      const wb = await readWorkbook(file)
      const sheet =
        wb.Sheets['Afwijking achterstallig'] || wb.Sheets[wb.SheetNames[0]] || null
      const stationCell = sheet?.B1?.v
      if (stationCell) {
        setStation(String(stationCell))
        addLog('Station gevuld vanuit dashboard.')
      }
    } catch (error) {
      addLog('Dashboard kon niet worden uitgelezen.', 'error')
    }
  }

  const runDataOphalen = async () => {
    if (!overzichtFile) return
    setBusyAction('data')
    addLog('Data ophalen gestart.')
    try {
      const wb = await readWorkbook(overzichtFile)
      const sheet = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
      if (!rows.length) throw new Error('Geen data gevonden in overzicht.')
      const headers = rows[0].map((value) => String(value ?? '').trim())

      const colCode = getColumnIndex(headers, 'Code')
      const colTitel = getColumnIndex(headers, 'Titel')
      const colMaatregelCode = getColumnIndex(headers, 'Code (2)')
      const colMaatregel = getColumnIndex(headers, 'Maatregel')
      const colStatus = getColumnIndex(headers, 'Status (2)')
      const colStatusSingle = getColumnIndex(headers, 'Status')
      const colActiehouder = getColumnIndex(headers, 'Actiehouder')
      const colOpsteller = getColumnIndex(headers, 'Opsteller')
      const colDatum = getColumnIndex(headers, 'Geplande datum klaar')

      if (
        [colCode, colTitel, colMaatregelCode, colMaatregel, colStatus, colActiehouder, colDatum].some(
          (index) => index === -1
        )
      ) {
        throw new Error('Kolomkoppen ontbreken in overzicht.')
      }

      const now = new Date()
      const thirtyOneDays = new Date(now.getTime() + 31 * 24 * 60 * 60 * 1000)
      const achterstallig = []
      const concept = []
      const actiehouderSet = new Set()

      for (let i = 1; i < rows.length; i += 1) {
        const row = rows[i]
        if (!row || row.length === 0) continue

        const status = String(row[colStatus] ?? '').trim()
        const statusSingle = String(row[colStatusSingle] ?? '').trim()
        const geplandeDatum = parseExcelDate(row[colDatum])
        const opmerking = geplandeDatum
          ? geplandeDatum < now
            ? 'Deadline verlopen'
            : geplandeDatum <= thirtyOneDays
              ? 'Deadline verloopt binnen 31 dagen'
              : 'Geen actie vereist'
          : 'Geen datum'

        if (status === 'Vigerend' && opmerking !== 'Geen actie vereist') {
          achterstallig.push({
            code: row[colCode],
            titel: row[colTitel],
            maatregelCode: row[colMaatregelCode],
            maatregel: row[colMaatregel],
            status,
            actiehouder: row[colActiehouder],
            geplandeDatum,
            opmerking,
          })
          if (row[colActiehouder]) {
            actiehouderSet.add(String(row[colActiehouder]))
          }
        }

        if (statusSingle === 'Concept' && colOpsteller !== -1) {
          concept.push({
            code: row[colCode],
            titel: row[colTitel],
            status: 'Concept',
            opsteller: row[colOpsteller],
            geplandeDatum,
          })
        }
      }

      achterstallig.sort((a, b) => String(a.code).localeCompare(String(b.code)))
      concept.sort((a, b) => String(a.code).localeCompare(String(b.code)))

      setAchterstalligRows(achterstallig)
      setConceptRows(concept)
      setActiehouders(Array.from(actiehouderSet).sort((a, b) => a.localeCompare(b)))
      addLog('Data ophalen afgerond.')
    } catch (error) {
      addLog(error instanceof Error ? error.message : 'Data ophalen mislukt.', 'error')
    } finally {
      setBusyAction('')
    }
  }

  const runEmailDraft = () => {
    const draft = buildEmailDraft(station)
    setEmailDraft(draft)
    addLog('Email concept gegenereerd.')
  }

  const runPowerBiExport = async () => {
    if (!databaseFile || !overzichtFile) return
    setBusyAction('powerbi')
    addLog('PowerBI data voorbereiden gestart.')
    try {
      const dbWb = await readWorkbook(databaseFile)
      const ovWb = await readWorkbook(overzichtFile)
      const dbSheet =
        dbWb.Sheets['Database'] || dbWb.Sheets[dbWb.SheetNames[0]] || null
      const ovSheet = ovWb.Sheets[ovWb.SheetNames[0]]
      if (!dbSheet || !ovSheet) {
        throw new Error('Database of overzicht ontbreekt.')
      }

      const dbRows = XLSX.utils.sheet_to_json(dbSheet, { header: 1, defval: '' })
      const ovRows = XLSX.utils.sheet_to_json(ovSheet, { header: 1, defval: '' })
      if (dbRows.length < 2 || ovRows.length < 2) {
        throw new Error('Niet genoeg rijen om data te kopieren.')
      }

      const dbHeaders = dbRows[1].map((value) => String(value ?? '').trim())
      const ovHeaders = ovRows[0].map((value) => String(value ?? '').trim())
      const ovHeaderMap = new Map(
        ovHeaders.map((header, index) => [normalize(header), index]).filter(([key]) => key)
      )
      const dataRowCount = ovRows.length - 1
      const dbDataRowCount = Math.max(dbRows.length - 2, 0)
      const maxRows = Math.max(dataRowCount, dbDataRowCount)
      const dateExportCol = getColumnIndex(dbHeaders, 'Date export')
      const stationCol = getColumnIndex(dbHeaders, 'Station')

      for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
        const ovRow = ovRows[rowIndex + 1] || []
        const dbRowIndex = rowIndex + 2

        dbHeaders.forEach((header, colIndex) => {
          const ovColIndex = ovHeaderMap.get(normalize(header))
          if (ovColIndex === undefined || ovColIndex === null || ovColIndex === -1) return
          const value = rowIndex < dataRowCount ? ovRow[ovColIndex] : ''
          setCellValue(dbSheet, dbRowIndex, colIndex, value)
        })

        if (dateExportCol !== -1) {
          setCellValue(
            dbSheet,
            dbRowIndex,
            dateExportCol,
            rowIndex < dataRowCount ? new Date() : ''
          )
        }
        if (stationCol !== -1) {
          setCellValue(
            dbSheet,
            dbRowIndex,
            stationCol,
            rowIndex < dataRowCount ? station || '' : ''
          )
        }
      }

      const range = XLSX.utils.decode_range(dbSheet['!ref'] || 'A1')
      range.e.r = Math.max(range.e.r, maxRows + 1)
      range.e.c = Math.max(range.e.c, dbHeaders.length - 1)
      dbSheet['!ref'] = XLSX.utils.encode_range(range)

      XLSX.writeFile(dbWb, 'Afwijkingen database bijgewerkt.xlsx')
      setPowerBiReady(true)
      addLog('PowerBI export gedownload.')
    } catch (error) {
      addLog(error instanceof Error ? error.message : 'PowerBI export mislukt.', 'error')
    } finally {
      setBusyAction('')
    }
  }

  const copyToClipboard = async (value) => {
    if (!value) return
    try {
      await navigator.clipboard.writeText(value)
      addLog('Tekst gekopieerd naar klembord.')
    } catch (error) {
      addLog('Kopieren mislukt.', 'error')
    }
  }

  const downloadDashboardExport = async () => {
    if (!achterstalligRows.length && !conceptRows.length && !actiehouders.length) {
      addLog('Geen data om te exporteren.', 'error')
      return
    }
    const headersAchterstallig = [
      'Afw. Code',
      'Afwijking Titel',
      'Maatregel Code',
      'Maatregel',
      'Status',
      'Actiehouder',
      'Geplande datum klaar',
      'Opmerking',
    ]
    const exportMetaRows = [
      ['Project', station || ''],
      ['Type', 'Afwijkingen overzicht'],
      ['Datum DB', new Date().toLocaleString('nl-NL')],
    ]
    const achterstalligData = achterstalligRows.map((row) => [
      row.code,
      row.titel,
      row.maatregelCode,
      row.maatregel,
      row.status,
      row.actiehouder,
      row.geplandeDatum,
      row.opmerking,
    ])
    const headersConcept = [
      'Afw. Code',
      'Afwijking Titel',
      'Status',
      'Opsteller',
      'Geplande datum klaar',
    ]
    const conceptData = conceptRows.map((row) => [
      row.code,
      row.titel,
      row.status,
      row.opsteller,
      row.geplandeDatum,
    ])
    const actiehouderRows = actiehouders.map((name) => [name])

    const workbook = new ExcelJS.Workbook()
    workbook.creator = 'afwijkingen-lab'

    const buildSheet = (name, headers, rows, tableName) => {
      const sheet = workbook.addWorksheet(name)
      exportMetaRows.forEach((row) => sheet.addRow(row))
      const tableRows =
        rows && rows.length ? rows : [headers.map(() => '')]
      sheet.addTable({
        name: tableName,
        ref: 'A4',
        headerRow: true,
        totalsRow: false,
        style: { theme: 'TableStyleLight1', showRowStripes: false },
        columns: headers.map((header) => ({ name: header })),
        rows: tableRows,
      })
      styleMetaBlock(sheet)
      styleTableHeader(sheet, headers.length)
      styleTableRows(sheet, tableRows.length, headers.length)
      const widths = computeColumnWidths(headers, rows, exportMetaRows)
      sheet.columns = headers.map((_, index) => ({
        width: widths[index],
      }))
      sheet.views = [{ state: 'frozen', ySplit: 4, topLeftCell: 'A5' }]
    }

    buildSheet(
      'Afwijking achterstallig',
      headersAchterstallig,
      achterstalligData,
      'AchterstalligTable'
    )
    buildSheet('Afwijking concept', headersConcept, conceptData, 'ConceptTable')
    buildSheet('Actiehouders', ['Actiehouder'], actiehouderRows, 'ActiehoudersTable')

    const filename = `Afwijkingen_dashboard_export_${buildTimestamp()}.xlsx`
    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = filename
    link.rel = 'noopener'
    document.body.appendChild(link)
    link.click()
    link.remove()
    setTimeout(() => URL.revokeObjectURL(url), 1000)
    addLog(`Dashboard export gedownload: ${filename}`)
  }

  const handleDrop = (target) => async (event) => {
    event.preventDefault()
    const file = event.dataTransfer.files?.[0] || null
    setDragTarget('')
    if (!file) return
    if (target === 'dashboard') {
      await handleDashboardUpload(file)
      return
    }
    if (target === 'database') {
      setDatabaseFile(file)
      setPowerBiReady(false)
      return
    }
    if (target === 'overzicht') {
      setOverzichtFile(file)
      setPowerBiReady(false)
    }
  }

  const handleDragOver = (target) => (event) => {
    event.preventDefault()
    if (dragTarget !== target) {
      setDragTarget(target)
    }
  }

  const handleDragLeave = (target) => (event) => {
    event.preventDefault()
    if (dragTarget === target) {
      setDragTarget('')
    }
  }

  return (
    <div className="app">
      <header className="hero">
        <div>
          <p className="eyebrow">Afwijkingen Workflow Lab</p>
          <h1>Afwijkingen dashboard lab</h1>
          <p className="subtitle">
            Upload het dashboard, de database en het overzicht. Daarna kun je data ophalen,
            een email opstellen en een PowerBI export maken.
          </p>
        </div>
        <div className="hero-card">
          <p className="stat-label">Status</p>
          <h3>Omgeving klaar voor acties</h3>
          <p className="meta">Deze tool verwerkt bestanden lokaal in de browser.</p>
        </div>
      </header>

      <section className="panel">
        <div className="panel-header">
          <h2>Uploads</h2>
          <div className="inline-field">
            <span>Station</span>
            <input
              type="text"
              value={station}
              onChange={(event) => setStation(event.target.value)}
              placeholder="Bijv. Utrecht CS"
            />
          </div>
        </div>
        <div className="upload-grid">
          <label
            className={`upload-card ${dashboardFile ? 'ready' : ''} ${
              dragTarget === 'dashboard' ? 'dragging' : ''
            }`}
            onDrop={handleDrop('dashboard')}
            onDragOver={handleDragOver('dashboard')}
            onDragLeave={handleDragLeave('dashboard')}
          >
            <span className="upload-title">Dashboard</span>
            <span className="upload-sub">Afwijkingen dashboard (.xlsm)</span>
            <span className="upload-file">{dashboardFile ? dashboardFile.name : 'Nog geen bestand'}</span>
            <input
              type="file"
              accept=".xlsm,.xlsx"
              onChange={(event) => handleDashboardUpload(event.target.files?.[0] || null)}
            />
            <span className="upload-cta">Kies bestand</span>
          </label>

          <label
            className={`upload-card ${databaseFile ? 'ready' : ''} ${
              dragTarget === 'database' ? 'dragging' : ''
            }`}
            onDrop={handleDrop('database')}
            onDragOver={handleDragOver('database')}
            onDragLeave={handleDragLeave('database')}
          >
            <span className="upload-title">Database</span>
            <span className="upload-sub">Afwijkingen database (.xlsx)</span>
            <span className="upload-file">{databaseFile ? databaseFile.name : 'Nog geen bestand'}</span>
            <input
              type="file"
              accept=".xlsx,.xlsm"
              onChange={(event) => {
                setDatabaseFile(event.target.files?.[0] || null)
                setPowerBiReady(false)
              }}
            />
            <span className="upload-cta">Kies bestand</span>
          </label>

          <label
            className={`upload-card ${overzichtFile ? 'ready' : ''} ${
              dragTarget === 'overzicht' ? 'dragging' : ''
            }`}
            onDrop={handleDrop('overzicht')}
            onDragOver={handleDragOver('overzicht')}
            onDragLeave={handleDragLeave('overzicht')}
          >
            <span className="upload-title">Overzicht</span>
            <span className="upload-sub">Afwijkingen overzicht (.xlsx)</span>
            <span className="upload-file">{overzichtFile ? overzichtFile.name : 'Nog geen bestand'}</span>
            <input
              type="file"
              accept=".xlsx,.xlsm"
              onChange={(event) => {
                setOverzichtFile(event.target.files?.[0] || null)
                setPowerBiReady(false)
              }}
            />
            <span className="upload-cta">Kies bestand</span>
          </label>
        </div>
      </section>

      <section className="panel">
        <div className="panel-header">
          <h2>Acties</h2>
          <div className="panel-actions">
            <button
              className="ghost"
              type="button"
              onClick={() => {
                setAchterstalligRows([])
                setConceptRows([])
                setActiehouders([])
                setEmailDraft(null)
                setPowerBiReady(false)
                setLogEntries([])
              }}
            >
              Reset scherm
            </button>
          </div>
        </div>
        <div className="action-grid">
          <button
            className="primary"
            type="button"
            onClick={runDataOphalen}
            disabled={!overzichtFile || busyAction === 'data'}
          >
            {busyAction === 'data' ? 'Data ophalen...' : 'Data ophalen'}
          </button>
          <button className="ghost" type="button" onClick={runEmailDraft}>
            Email opstellen
          </button>
          <button
            className="primary"
            type="button"
            onClick={runPowerBiExport}
            disabled={!databaseFile || !overzichtFile || busyAction === 'powerbi'}
          >
            {busyAction === 'powerbi' ? 'PowerBI data...' : 'PowerBI data'}
          </button>
          <button className="ghost" type="button" onClick={() => void downloadDashboardExport()}>
            Dashboard export
          </button>
        </div>
        <div className="output-cards">
          <div className="stat-card">
            <p className="stat-label">Achterstallig</p>
            <p className="stat-value">{summaryStats.achterstallig}</p>
            <p className="stat-note">Maatregelen met actie</p>
          </div>
          <div className="stat-card">
            <p className="stat-label">Concept</p>
            <p className="stat-value">{summaryStats.concept}</p>
            <p className="stat-note">Nog in concept</p>
          </div>
          <div className="stat-card">
            <p className="stat-label">Actiehouders</p>
            <p className="stat-value">{summaryStats.actiehouders}</p>
            <p className="stat-note">Unieke namen</p>
          </div>
          <div className="stat-card">
            <p className="stat-label">PowerBI export</p>
            <p className="stat-value">{powerBiReady ? 'Klaar' : 'Niet gedaan'}</p>
            <p className="stat-note">Download gereed</p>
          </div>
        </div>
      </section>

      <section className="panel output-panel">
        <div className="panel-header">
          <h2>Resultaten</h2>
          <p className="meta">Afwijkingen achterstallig en concept overzicht.</p>
        </div>
        <div className="toggle-row">
          <button
            className="ghost toggle"
            type="button"
            onClick={() =>
              setActivePanel((prev) => (prev === 'results' ? 'none' : 'results'))
            }
          >
            {activePanel === 'results' ? 'Hide resultaten' : 'Show resultaten'}
          </button>
          <button
            className="ghost toggle"
            type="button"
            onClick={() =>
              setActivePanel((prev) => (prev === 'email' ? 'none' : 'email'))
            }
          >
            {activePanel === 'email' ? 'Hide email' : 'Show email'}
          </button>
          <button
            className="ghost toggle"
            type="button"
            onClick={() => setActivePanel((prev) => (prev === 'log' ? 'none' : 'log'))}
          >
            {activePanel === 'log' ? 'Hide log' : 'Show log'}
          </button>
        </div>
        <div className={`toggle-panel ${activePanel === 'results' ? 'open' : ''}`}>
          <div className="panel-body">
            <div className="output-stack">
            <div className="table-card">
              <div className="table-header">
                <h3>Achterstallig</h3>
                <span className="meta">{achterstalligRows.length} rijen</span>
              </div>
            <div className="table-scroll">
              {achterstalligRows.length ? (
                <table>
                  <thead>
                    <tr>
                      <th>Code</th>
                      <th>Afwijking titel</th>
                      <th>Maatregel code</th>
                      <th>Maatregel</th>
                      <th>Status</th>
                      <th>Actiehouder</th>
                      <th>Geplande datum</th>
                      <th>Opmerking</th>
                    </tr>
                  </thead>
                  <tbody>
                    {achterstalligRows.map((row, index) => (
                      <tr key={`achterstallig-${index}`}>
                        <td>{row.code}</td>
                        <td>{row.titel}</td>
                        <td>{row.maatregelCode}</td>
                        <td>{row.maatregel}</td>
                        <td>{row.status}</td>
                        <td>{row.actiehouder}</td>
                        <td>{formatDate(row.geplandeDatum)}</td>
                        <td>{row.opmerking}</td>
                      </tr>
                    ))}
                    </tbody>
                  </table>
                ) : (
                  <p className="empty">Nog geen data opgehaald.</p>
                )}
              </div>
            </div>

            <div className="table-card">
              <div className="table-header">
                <h3>Concept</h3>
                <span className="meta">{conceptRows.length} rijen</span>
              </div>
              <div className="table-scroll">
                {conceptRows.length ? (
                  <table>
                    <thead>
                      <tr>
                        <th>Code</th>
                        <th>Afwijking titel</th>
                        <th>Opsteller</th>
                        <th>Geplande datum</th>
                      </tr>
                    </thead>
                    <tbody>
                      {conceptRows.map((row, index) => (
                        <tr key={`concept-${index}`}>
                          <td>{row.code}</td>
                          <td>{row.titel}</td>
                          <td>{row.opsteller}</td>
                          <td>{formatDate(row.geplandeDatum)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <p className="empty">Nog geen concept data.</p>
                )}
              </div>
            </div>

            <div className="table-card">
              <div className="table-header">
                <h3>Actiehouders</h3>
                <span className="meta">{actiehouders.length} namen</span>
              </div>
              <div className="list-grid">
                {actiehouders.length ? (
                  actiehouders.map((name) => (
                    <span className="pill" key={name}>
                      {name}
                    </span>
                  ))
                ) : (
                  <p className="empty">Nog geen actiehouders.</p>
                )}
              </div>
            </div>
            </div>
          </div>
        </div>

        <div className={`toggle-panel ${activePanel === 'email' ? 'open' : ''}`}>
          <div className="panel-body">
            <div className="table-card">
              <div className="panel-header">
                <h3>Email concept</h3>
                <div className="panel-actions">
                  <button
                    className="ghost"
                    type="button"
                    onClick={() => copyToClipboard(emailDraft?.subject)}
                    disabled={!emailDraft}
                  >
                    Kopieer onderwerp
                  </button>
                  <button
                    className="ghost"
                    type="button"
                    onClick={() => copyToClipboard(emailDraft?.body)}
                    disabled={!emailDraft}
                  >
                    Kopieer bericht
                  </button>
                </div>
              </div>
              {emailDraft ? (
                <div className="email-draft">
                  <div>
                    <p className="stat-label">Onderwerp</p>
                    <p className="email-subject">{emailDraft.subject}</p>
                  </div>
                  <div>
                    <p className="stat-label">Bericht</p>
                    <textarea readOnly value={emailDraft.body} rows={8} />
                  </div>
                </div>
              ) : (
                <p className="empty">Klik op "Email opstellen" om een concept te maken.</p>
              )}
            </div>
          </div>
        </div>

        <div className={`toggle-panel ${activePanel === 'log' ? 'open' : ''}`}>
          <div className="panel-body">
            <div className="table-card log-panel">
              <div className="panel-header">
                <h3>Logboek</h3>
              </div>
              <div className="log-entries">
                {logEntries.length ? (
                  <ul>
                    {logEntries.map((entry, index) => (
                      <li key={`${entry.timestamp}-${index}`} className={entry.type || 'info'}>
                        <span className="log-time">{entry.timestamp}</span>
                        <span className="log-msg">{entry.message}</span>
                      </li>
                    ))}
                  </ul>
                ) : (
                  <p className="empty">Nog geen log regels.</p>
                )}
              </div>
            </div>
          </div>
        </div>
      </section>
    </div>
  )
}

export default App
