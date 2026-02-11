import { useEffect, useMemo, useRef, useState } from 'react'
import ExcelJS from 'exceljs'
import html2canvas from 'html2canvas'
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
    "Beste collega's,",
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

const polarToCartesian = (centerX, centerY, radius, angleInDegrees) => {
  const angleInRadians = ((angleInDegrees - 90) * Math.PI) / 180.0
  return {
    x: centerX + radius * Math.cos(angleInRadians),
    y: centerY + radius * Math.sin(angleInRadians),
  }
}

const describeArc = (x, y, radius, startAngle, endAngle) => {
  const start = polarToCartesian(x, y, radius, endAngle)
  const end = polarToCartesian(x, y, radius, startAngle)
  const largeArcFlag = endAngle - startAngle <= 180 ? '0' : '1'
  return `M ${start.x} ${start.y} A ${radius} ${radius} 0 ${largeArcFlag} 0 ${end.x} ${end.y}`
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
  const [lastEmailStation, setLastEmailStation] = useState('')
  const [showHelp, setShowHelp] = useState(false)
  const [helpItems, setHelpItems] = useState([])
  const [isHelpCompact, setIsHelpCompact] = useState(false)
  const [powerBiStats, setPowerBiStats] = useState(null)
  const [statsDownloading, setStatsDownloading] = useState(false)
  const helpRafRef = useRef(null)
  const uploadsRef = useRef(null)
  const actionsRef = useRef(null)
  const outputRef = useRef(null)
  const toggleRowRef = useRef(null)
  const statsRef = useRef(null)

  const summaryStats = useMemo(
    () => ({
      achterstallig: achterstalligRows.length,
      concept: conceptRows.length,
      actiehouders: actiehouders.length,
    }),
    [achterstalligRows.length, conceptRows.length, actiehouders.length]
  )

  const powerBiPercent = powerBiStats ? Math.round(powerBiStats.onTimePercent) : 0
  const powerBiOnTime = powerBiStats ? powerBiStats.validDates - powerBiStats.overdueCount : 0
  const powerBiOnTimePercent = powerBiStats
    ? powerBiStats.validDates
      ? Math.round((powerBiOnTime / powerBiStats.validDates) * 100)
      : 0
    : 0
  const powerBiLatePercent = powerBiStats
    ? powerBiStats.validDates
      ? Math.round((powerBiStats.overdueCount / powerBiStats.validDates) * 100)
      : 0
    : 0
  const powerBiLateAngle = powerBiStats
    ? powerBiStats.validDates
      ? (powerBiStats.overdueCount / powerBiStats.validDates) * 360
      : 0
    : 0
  const powerBiDisciplineData = powerBiStats
    ? Array.from(powerBiStats.disciplineTotals?.entries?.() || [])
        .map(([discipline, totalDays]) => ({
          discipline,
          totalDays,
          count: powerBiStats.disciplineCounts?.get?.(discipline) || 0,
        }))
        .map((row) => ({
          ...row,
          avgDays: row.count ? row.totalDays / row.count : 0,
        }))
        .sort((a, b) => b.avgDays - a.avgDays)
    : []
  const powerBiDisciplineMaxCount = powerBiDisciplineData.length
    ? Math.max(...powerBiDisciplineData.map((row) => row.count), 1)
    : 1
  const powerBiDisciplineMaxAvg = powerBiDisciplineData.length
    ? Math.max(...powerBiDisciplineData.map((row) => row.avgDays), 1)
    : 1
  const powerBiDisciplineScaleMax = Math.max(powerBiDisciplineMaxAvg, powerBiDisciplineMaxCount, 1)
  const powerBiTrafficLabel = powerBiStats
    ? powerBiStats.traffic === 'red'
      ? 'Rood'
      : powerBiStats.traffic === 'orange'
        ? 'Oranje'
        : 'Groen'
    : 'Onbekend'

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

        const actiehouderValue = String(row[colActiehouder] ?? '').trim()
        if (status === 'Vigerend' && opmerking !== 'Geen actie vereist' && actiehouderValue) {
          achterstallig.push({
            code: row[colCode],
            titel: row[colTitel],
            maatregelCode: row[colMaatregelCode],
            maatregel: row[colMaatregel],
            status,
            actiehouder: actiehouderValue,
            geplandeDatum,
            opmerking,
          })
          if (actiehouderValue) {
            actiehouderSet.add(actiehouderValue)
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
      await runPowerBiExport()
    } catch (error) {
      addLog(error instanceof Error ? error.message : 'Data ophalen mislukt.', 'error')
    } finally {
      setBusyAction('')
    }
  }

  const runEmailDraft = () => {
    const draft = buildEmailDraft(station)
    setEmailDraft(draft)
    setLastEmailStation(station)
    addLog('Email concept gegenereerd.')
  }

  const runPowerBiExport = async () => {
    if (!overzichtFile) return
    setBusyAction((prev) => (prev ? prev : 'powerbi'))
    addLog('PowerBI concept gestart.')
    try {
      const wb = await readWorkbook(overzichtFile)
      const sheet = wb.Sheets[wb.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })
      if (!rows.length) throw new Error('Geen data gevonden in overzicht.')
      const headers = rows[0].map((value) => String(value ?? '').trim())

      const colBeoordeling = getColumnIndex(headers, 'Maatregelen beoordeling')
      const colStatus = getColumnIndex(headers, 'Status (2)')
      const colPlanned = getColumnIndex(headers, 'Geplande datum klaar')
      const colDone = getColumnIndex(headers, 'Datum klaar')
      const colActiehouder = getColumnIndex(headers, 'Actiehouder')
      const colMaatregel = getColumnIndex(headers, 'Maatregel')
      const colCode = getColumnIndex(headers, 'Code')
      const colTitel = getColumnIndex(headers, 'Titel')
      const colDiscipline = getColumnIndex(headers, 'Veroorzakende Discipline')

      if (
        [colBeoordeling, colStatus, colPlanned, colDone].some((index) => index === -1)
      ) {
        throw new Error('Kolomkoppen ontbreken in overzicht voor PowerBI.')
      }

      let totalFiltered = 0
      let validDates = 0
      let overdueCount = 0
      let missingDates = 0
      const onTimeRows = []
      const lateRows = []
      const missingRows = []
      const disciplineTotals = new Map()
      const disciplineCounts = new Map()

      for (let i = 1; i < rows.length; i += 1) {
        const row = rows[i]
        if (!row || row.length === 0) continue

        const beoordeling = String(row[colBeoordeling] ?? '').trim()
        const status = String(row[colStatus] ?? '').trim()
        if (
          normalize(beoordeling) !== normalize('Maatregelen nodig') ||
          normalize(status) !== normalize('Afgehandeld')
        ) {
          continue
        }

        totalFiltered += 1
        const plannedDate = parseExcelDate(row[colPlanned])
        const doneDate = parseExcelDate(row[colDone])
        const actiehouder = colActiehouder !== -1 ? String(row[colActiehouder] ?? '').trim() : ''
        const maatregel = colMaatregel !== -1 ? String(row[colMaatregel] ?? '').trim() : ''

        if (!plannedDate || !doneDate) {
          missingDates += 1
          missingRows.push({
            code: colCode !== -1 ? row[colCode] : '',
            titel: colTitel !== -1 ? row[colTitel] : '',
            maatregel,
            actiehouder,
            geplandeDatum: plannedDate,
            datumKlaar: doneDate,
          })
          continue
        }

        validDates += 1
        const rowPayload = {
          code: colCode !== -1 ? row[colCode] : '',
          titel: colTitel !== -1 ? row[colTitel] : '',
          maatregel,
          actiehouder,
          geplandeDatum: plannedDate,
          datumKlaar: doneDate,
        }
        if (doneDate > plannedDate) {
          overdueCount += 1
          lateRows.push(rowPayload)
          const diffDays = Math.ceil((doneDate - plannedDate) / 86400000)
          const discipline = colDiscipline !== -1 ? String(row[colDiscipline] ?? '').trim() : ''
          if (discipline) {
            disciplineTotals.set(
              discipline,
              (disciplineTotals.get(discipline) || 0) + Math.max(diffDays, 0)
            )
            disciplineCounts.set(discipline, (disciplineCounts.get(discipline) || 0) + 1)
          }
        } else {
          onTimeRows.push(rowPayload)
        }

      }

      const overduePercent = validDates ? (overdueCount / validDates) * 100 : 0
      const onTimePercent = validDates ? ((validDates - overdueCount) / validDates) * 100 : 0
      const traffic =
        onTimePercent <= 35 ? 'red' : onTimePercent < 75 ? 'orange' : 'green'

      setPowerBiStats({
        totalFiltered,
        validDates,
        overdueCount,
        missingDates,
        overduePercent,
        onTimePercent,
        traffic,
        onTimeRows,
        lateRows,
        missingRows,
        disciplineTotals,
        disciplineCounts,
      })
      setPowerBiReady(true)
      addLog('PowerBI data klaar.')
      if (missingDates > 0) {
        addLog(
          `${missingDates} rijen missen Geplande datum klaar of Datum klaar.`,
          'error'
        )
      }
    } catch (error) {
      addLog(error instanceof Error ? error.message : 'PowerBI data ophalen mislukt.', 'error')
      setPowerBiReady(false)
      setPowerBiStats(null)
    } finally {
      setBusyAction((prev) => (prev === 'powerbi' ? '' : prev))
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

  const downloadStatsImage = async () => {
    if (!statsRef.current || !powerBiStats || statsDownloading) return
    setStatsDownloading(true)
    try {
      const canvas = await html2canvas(statsRef.current, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
      })
      const url = canvas.toDataURL('image/png')
      const link = document.createElement('a')
      link.href = url
      link.download = `Statistieken_${buildTimestamp()}.png`
      link.rel = 'noopener'
      document.body.appendChild(link)
      link.click()
      link.remove()
      addLog('Statistieken afbeelding gedownload.')
    } catch (error) {
      addLog('Download statistieken mislukt.', 'error')
    } finally {
      setStatsDownloading(false)
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
      setPowerBiStats(null)
      return
    }
    if (target === 'overzicht') {
      setOverzichtFile(file)
      setPowerBiReady(false)
      setPowerBiStats(null)
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

  useEffect(() => {
    if (activePanel !== 'email') return
    if (!emailDraft || lastEmailStation !== station) {
      runEmailDraft()
    }
  }, [activePanel, emailDraft, lastEmailStation, station])

  useEffect(() => {
    if (!showHelp) return
    const buildHelpItems = () => {
      const compact = window.innerWidth <= 1024 || window.innerHeight <= 720
      setIsHelpCompact(compact)
      const padding = 16
      const cardWidth = 280
      const cardGap = 14
      const clamp = (value, min, max) => Math.min(Math.max(value, min), max)
      const withRect = (ref) => {
        if (!ref.current) return null
        return ref.current.getBoundingClientRect()
      }
      const placeCard = (rect, preferLeft = false) => {
        const roomRight = window.innerWidth - rect.right
        const roomLeft = rect.left
        const canPlaceRight = roomRight >= cardWidth + cardGap
        const canPlaceLeft = roomLeft >= cardWidth + cardGap
        let left = clamp(rect.left, padding, window.innerWidth - cardWidth - padding)
        if (preferLeft) {
          if (canPlaceLeft) {
            left = rect.left - cardWidth - cardGap
          } else if (canPlaceRight) {
            left = rect.right + cardGap
          }
        } else if (canPlaceRight) {
          left = rect.right + cardGap
        }
        const top = clamp(rect.top, padding, window.innerHeight - 140)
        return { left, top }
      }
      const nextItems = []
      const uploadsRect = withRect(uploadsRef)
      if (uploadsRect) {
        nextItems.push({
          id: 'uploads',
          title: 'Stap 1: Uploads',
          body: 'Upload het dashboard, de database en het overzicht (drag & drop kan).',
          ...(compact
            ? {}
            : {
                spot: {
                  top: uploadsRect.top,
                  left: uploadsRect.left,
                  width: uploadsRect.width,
                  height: uploadsRect.height,
                },
                card: placeCard(uploadsRect),
              }),
        })
      }
      const actionsRect = withRect(actionsRef)
      if (actionsRect) {
        nextItems.push({
          id: 'actions',
          title: 'Stap 2: Acties',
          body: 'Klik "Data ophalen" en daarna "Dashboard export" om te downloaden.',
          ...(compact
            ? {}
            : {
                spot: {
                  top: actionsRect.top,
                  left: actionsRect.left,
                  width: actionsRect.width,
                  height: actionsRect.height,
                },
                card: placeCard(actionsRect),
              }),
        })
      }
      const toggleRect = withRect(toggleRowRef)
      if (toggleRect) {
        nextItems.push({
          id: 'toggles',
          title: 'Stap 3: Panels',
          body: 'Gebruik de toggles om Resultaten, Email, Log of PowerBI te openen.',
          ...(compact
            ? {}
            : {
                spot: {
                  top: toggleRect.top,
                  left: toggleRect.left,
                  width: toggleRect.width,
                  height: toggleRect.height,
                },
                card: placeCard(toggleRect, true),
              }),
        })
      }
      const outputRect = withRect(outputRef)
      if (outputRect) {
        nextItems.push({
          id: 'output',
          title: 'Stap 4: Resultaten',
          body: 'Bekijk tabellen, kopieer email en controleer het logboek.',
          ...(compact
            ? {}
            : {
                spot: {
                  top: outputRect.top,
                  left: outputRect.left,
                  width: outputRect.width,
                  height: outputRect.height,
                },
                card: placeCard(outputRect),
              }),
        })
      }
      if (!compact) {
        const sorted = [...nextItems].filter((item) => item.card).sort((a, b) => a.card.top - b.card.top)
        const minGap = 16
        for (let i = 1; i < sorted.length; i += 1) {
          const prev = sorted[i - 1].card
          const current = sorted[i].card
          if (current.top < prev.top + 140 + minGap) {
            current.top = prev.top + 140 + minGap
          }
        }
      }
      setHelpItems(nextItems)
    }
    const scheduleUpdate = () => {
      if (helpRafRef.current != null) return
      helpRafRef.current = requestAnimationFrame(() => {
        helpRafRef.current = null
        buildHelpItems()
      })
    }
    scheduleUpdate()
    window.addEventListener('resize', scheduleUpdate)
    window.addEventListener('scroll', scheduleUpdate, true)
    return () => {
      window.removeEventListener('resize', scheduleUpdate)
      window.removeEventListener('scroll', scheduleUpdate, true)
      if (helpRafRef.current != null) {
        cancelAnimationFrame(helpRafRef.current)
        helpRafRef.current = null
      }
    }
  }, [showHelp])

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

      <section className="panel" ref={uploadsRef}>
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
                setPowerBiStats(null)
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
                setPowerBiStats(null)
              }}
            />
            <span className="upload-cta">Kies bestand</span>
          </label>
        </div>
      </section>

      <section className="panel" ref={actionsRef}>
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
                setPowerBiStats(null)
                setLogEntries([])
              }}
            >
              Verversen
            </button>
          </div>
        </div>
        <div className="action-grid">
          <button
            className="ghost"
            type="button"
            onClick={runDataOphalen}
            disabled={!overzichtFile || busyAction === 'data'}
          >
            {busyAction === 'data' ? 'Data ophalen...' : 'Data ophalen'}
          </button>
          <button
            className="primary"
            type="button"
            onClick={() => void downloadDashboardExport()}
            disabled={!achterstalligRows.length && !conceptRows.length && !actiehouders.length}
          >
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
            <p className="stat-label">Statistieken</p>
            <p className="stat-value">{powerBiReady ? 'Gereed' : '-'}</p>
            <p className="stat-note">Download gereed</p>
          </div>
        </div>
      </section>

      <section className="panel output-panel" ref={outputRef}>
        <div className="panel-header">
          <h2>Resultaten</h2>
          <p className="meta">Afwijkingen achterstallig en concept overzicht.</p>
        </div>
        <div className="toggle-row" ref={toggleRowRef}>
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
          <button
            className="ghost toggle"
            type="button"
            onClick={() => setActivePanel((prev) => (prev === 'powerbi' ? 'none' : 'powerbi'))}
          >
            {activePanel === 'powerbi' ? 'Hide statistieken' : 'Show statistieken'}
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

        <div className={`toggle-panel ${activePanel === 'powerbi' ? 'open' : ''}`}>
          <div className="panel-body">
            <div className="table-card" ref={statsRef}>
              <div className="table-header">
                <div className="header-stack">
                  <h3>Statistieken</h3>
                  <span className="meta">
                    {new Date().toLocaleString('nl-NL', {
                      dateStyle: 'short',
                      timeStyle: 'short',
                    })}
                  </span>
                </div>
                <button
                  className="icon-button"
                  type="button"
                  onClick={downloadStatsImage}
                  disabled={!powerBiStats || statsDownloading}
                  aria-label="Download statistieken"
                >
                  <svg viewBox="0 0 24 24" aria-hidden="true">
                    <path d="M12 3a1 1 0 0 1 1 1v8.59l2.3-2.3a1 1 0 1 1 1.4 1.42l-4 4a1 1 0 0 1-1.4 0l-4-4a1 1 0 0 1 1.4-1.42l2.3 2.3V4a1 1 0 0 1 1-1z" />
                    <path d="M5 19a1 1 0 0 1 1-1h12a1 1 0 1 1 0 2H6a1 1 0 0 1-1-1z" />
                  </svg>
                </button>
              </div>
              {powerBiStats ? (
                <div className="stats-capture">
                <div className="chart-grid">
                  <div className="chart-card">
                    <div className="chart-header">
                      <span className="chart-title">Stoplicht chart</span>
                      <span className="meta">{powerBiPercent}%</span>
                    </div>
                    <div className="gauge">
                      <svg className="gauge-svg" viewBox="0 0 200 120" aria-hidden="true">
                        <path
                          d="M 20 100 A 80 80 0 0 1 180 100"
                          className="gauge-segment red"
                          pathLength="100"
                        />
                        <path
                          d="M 20 100 A 80 80 0 0 1 180 100"
                          className="gauge-segment orange"
                          pathLength="100"
                        />
                        <path
                          d="M 20 100 A 80 80 0 0 1 180 100"
                          className="gauge-segment green"
                          pathLength="100"
                        />
                        <line
                          x1="100"
                          y1="100"
                          x2={polarToCartesian(100, 100, 62, -90 + powerBiPercent * 1.8).x}
                          y2={polarToCartesian(100, 100, 62, -90 + powerBiPercent * 1.8).y}
                          className="gauge-needle"
                        />
                        <circle cx="100" cy="100" r="10" className="gauge-cap" />
                      </svg>
                      <div className="gauge-value">{powerBiPercent}</div>
                    </div>
                    <div className="gauge-legend">
                      <span className="legend-chip red">{'<= 35%'}</span>
                      <span className="legend-chip orange">{'< 75%'}</span>
                      <span className="legend-chip green">{'>= 75%'}</span>
                    </div>
                  </div>
                  <div className="chart-card">
                    <div className="chart-header">
                      <span className="chart-title">Planning gereed chart</span>
                      <span className="meta">{powerBiStats.validDates} rijen</span>
                    </div>
                    <div className="pie-wrap">
                      <svg
                        className="pie-svg"
                        viewBox="0 0 140 140"
                        preserveAspectRatio="xMidYMid meet"
                        aria-hidden="true"
                      >
                        {powerBiLateAngle > 0 && powerBiLateAngle < 360 ? (
                          <path
                            className="pie-bg"
                            d={describeArc(70, 70, 46, powerBiLateAngle, 360)}
                          />
                        ) : powerBiLateAngle <= 0 ? (
                          <circle className="pie-bg" cx="70" cy="70" r="46" />
                        ) : null}
                        {powerBiLateAngle >= 360 ? (
                          <circle className="pie-slice" cx="70" cy="70" r="46" />
                        ) : powerBiLateAngle > 0 ? (
                          <path
                            className="pie-slice"
                            d={describeArc(70, 70, 46, 0, powerBiLateAngle)}
                          />
                        ) : null}
                      </svg>
                      <div className="pie-legend">
                        <span className="legend-row">
                          <span className="legend-dot red" />
                          Te laat: {powerBiStats.overdueCount}
                        </span>
                        <span className="legend-row">
                          <span className="legend-dot green" />
                          Op tijd: {powerBiOnTime}
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="chart-card">
                    <div className="chart-header">
                      <span className="chart-title">Details</span>
                    </div>
                    <div className="ring-legend">
                      <span>Gefilterd: {powerBiStats.totalFiltered}</span>
                      <span>Geldige datums: {powerBiStats.validDates}</span>
                      <span>Overschreden: {powerBiStats.overdueCount}</span>
                      <span>Op tijd: {powerBiOnTime}</span>
                      <span>Ontbrekend: {powerBiStats.missingDates}</span>
                    </div>
                  </div>
                </div>
                <div className="chart-card wide-chart">
                  <div className="chart-header">
                    <span className="chart-title">Veroorzakende discipline</span>
                    <span className="meta">Gem. dagen te laat + aantal</span>
                  </div>
                  {powerBiDisciplineData.length ? (
                    <div className="discipline-chart">
                      <div className="discipline-legend">
                        <span className="legend-row">
                          <span className="legend-dot blue" />
                          Gemiddeld dagen te laat
                        </span>
                        <span className="legend-row">
                          <span className="legend-dot navy" />
                          Aantal afwijkingen
                        </span>
                      </div>
                      <div className="discipline-bars-wrap">
                        <div className="discipline-axis">
                          {[1, 0.75, 0.5, 0.25, 0].map((tick) => (
                            <div className="axis-tick" key={tick}>
                              <span className="axis-label">
                                {Math.round(powerBiDisciplineScaleMax * tick)}
                              </span>
                            </div>
                          ))}
                        </div>
                        <div className="discipline-bars">
                          {powerBiDisciplineData.map((item) => (
                            <div className="discipline-bar" key={item.discipline}>
                              <div
                                className="bar-pair"
                                data-tooltip={`Gemiddeld dagen te laat: ${Math.round(
                                  item.avgDays
                                )}\nAantal afwijkingen: ${item.count}`}
                              >
                                <span
                                  className="bar-avg"
                                  style={{
                                    height: `${(item.avgDays / powerBiDisciplineScaleMax) * 100}%`,
                                  }}
                                />
                                <span
                                  className="bar-count"
                                  style={{
                                    height: `${(item.count / powerBiDisciplineScaleMax) * 100}%`,
                                  }}
                                />
                              </div>
                              <span className="bar-label" title={item.discipline}>
                                {item.discipline}
                              </span>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                  ) : (
                    <p className="empty">Geen discipline-data beschikbaar.</p>
                  )}
                </div>
                <div className="table-card mini-table">
                  <div className="table-header">
                    <h3>Planning gereed overzicht</h3>
                    <span className="meta">Geldige datums: {powerBiStats.validDates}</span>
                  </div>
                  <div className="table-scroll">
                    <table>
                      <thead>
                        <tr>
                          <th>Status</th>
                          <th>Aantal</th>
                          <th>Percentage</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td>Op tijd</td>
                          <td>{powerBiOnTime}</td>
                          <td>{powerBiOnTimePercent}%</td>
                        </tr>
                        <tr>
                          <td>Te laat</td>
                          <td>{powerBiStats.overdueCount}</td>
                          <td>{powerBiLatePercent}%</td>
                        </tr>
                        <tr>
                          <td>Ontbrekend</td>
                          <td>{powerBiStats.missingDates}</td>
                          <td>-</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
                <div className="table-card mini-table">
                  <div className="table-header">
                    <h3>Op tijd</h3>
                    <span className="meta">{powerBiStats.onTimeRows.length} rijen</span>
                  </div>
                  <div className="table-scroll">
                    {powerBiStats.onTimeRows.length ? (
                      <table>
                        <thead>
                          <tr>
                            <th>Code</th>
                            <th>Titel</th>
                            <th>Maatregel</th>
                            <th>Actiehouder</th>
                            <th>Geplande datum</th>
                            <th>Datum klaar</th>
                          </tr>
                        </thead>
                        <tbody>
                          {powerBiStats.onTimeRows.map((row, index) => (
                            <tr key={`ontime-${index}`}>
                              <td>{row.code}</td>
                              <td>{row.titel}</td>
                              <td>{row.maatregel}</td>
                              <td>{row.actiehouder}</td>
                              <td>{formatDate(row.geplandeDatum)}</td>
                              <td>{formatDate(row.datumKlaar)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    ) : (
                      <p className="empty">Geen op tijd regels.</p>
                    )}
                  </div>
                </div>
                <div className="table-card mini-table">
                  <div className="table-header">
                    <h3>Te laat</h3>
                    <span className="meta">{powerBiStats.lateRows.length} rijen</span>
                  </div>
                  <div className="table-scroll">
                    {powerBiStats.lateRows.length ? (
                      <table>
                        <thead>
                          <tr>
                            <th>Code</th>
                            <th>Titel</th>
                            <th>Maatregel</th>
                            <th>Actiehouder</th>
                            <th>Geplande datum</th>
                            <th>Datum klaar</th>
                          </tr>
                        </thead>
                        <tbody>
                          {powerBiStats.lateRows.map((row, index) => (
                            <tr key={`late-${index}`}>
                              <td>{row.code}</td>
                              <td>{row.titel}</td>
                              <td>{row.maatregel}</td>
                              <td>{row.actiehouder}</td>
                              <td>{formatDate(row.geplandeDatum)}</td>
                              <td>{formatDate(row.datumKlaar)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    ) : (
                      <p className="empty">Geen te laat regels.</p>
                    )}
                  </div>
                </div>
                <div className="table-card mini-table">
                  <div className="table-header">
                    <h3>Ontbrekende datums</h3>
                    <span className="meta">{powerBiStats.missingRows.length} rijen</span>
                  </div>
                  <div className="table-scroll">
                    {powerBiStats.missingRows.length ? (
                      <table>
                        <thead>
                          <tr>
                            <th>Code</th>
                            <th>Titel</th>
                            <th>Maatregel</th>
                            <th>Actiehouder</th>
                            <th>Geplande datum</th>
                            <th>Datum klaar</th>
                          </tr>
                        </thead>
                        <tbody>
                          {powerBiStats.missingRows.map((row, index) => (
                            <tr key={`missing-${index}`}>
                              <td>{row.code}</td>
                              <td>{row.titel}</td>
                              <td>{row.maatregel}</td>
                              <td>{row.actiehouder}</td>
                              <td>{formatDate(row.geplandeDatum)}</td>
                              <td>{formatDate(row.datumKlaar)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    ) : (
                      <p className="empty">Geen ontbrekende datums.</p>
                    )}
                  </div>
                </div>
                </div>
              ) : (
                <p className="empty">
                  Klik op "Data ophalen" om de statistieken te visualiseren.
                </p>
              )}
            </div>
          </div>
        </div>
      </section>
      <div className={`help-fab-wrap ${showHelp ? 'open' : ''}`}>
        <button
          className="help-fab"
          type="button"
          onClick={() => setShowHelp((prev) => !prev)}
        >
          ?
        </button>
        <span className="help-fab-label">Hulp nodig? Jonathan van der Gouwe</span>
      </div>
      {showHelp ? (
        <div
          className={`help-overlay ${isHelpCompact ? 'compact' : ''}`}
          onClick={() => {
            setShowHelp(false)
          }}
        >
          <button
            className="help-close ghost"
            type="button"
            onClick={(event) => {
              event.stopPropagation()
              setShowHelp(false)
            }}
          >
            Sluiten
          </button>
          {isHelpCompact ? (
            <div
              className="help-list"
              onClick={(event) => {
                event.stopPropagation()
              }}
            >
              {helpItems.map((item) => (
                <div key={item.id} className="help-card help-card-compact">
                  <h4>{item.title}</h4>
                  <p>{item.body}</p>
                </div>
              ))}
            </div>
          ) : (
            helpItems.map((item) => (
              <div key={item.id} className="help-item">
                <div
                  className="help-spot"
                  style={{
                    top: `${item.spot.top}px`,
                    left: `${item.spot.left}px`,
                    width: `${item.spot.width}px`,
                    height: `${item.spot.height}px`,
                  }}
                />
                <div
                  className="help-card"
                  style={{ top: `${item.card.top}px`, left: `${item.card.left}px` }}
                  onClick={(event) => event.stopPropagation()}
                >
                  <h4>{item.title}</h4>
                  <p>{item.body}</p>
                </div>
              </div>
            ))
          )}
        </div>
      ) : null}
    </div>
  )
}

export default App
