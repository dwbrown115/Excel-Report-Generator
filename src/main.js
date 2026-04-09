import './style.css'
import * as XLSX from 'xlsx'

document.querySelector('#app').innerHTML = `
  <main class="app-shell">
    <header>
      <p class="eyebrow">Excel Report Generator</p>
      <h1>Mapped Unique Values from Columns A, B, and C</h1>
      <p class="subtitle">
        Upload a spreadsheet and this page will map unique A values to unique B/C pairs
        found under each A value.
      </p>
    </header>

    <label class="file-input-wrap" for="excel-file">
      <span>Select Excel File</span>
      <input id="excel-file" type="file" accept=".xlsx,.xlsm,.xlsb,.xltx,.xltm,.xlam,.xls,.xlt,.xml,.csv,.txt,.prn,.dif,.slk,.ods,.fods" />
    </label>

    <p id="status" class="status">Waiting for a file...</p>

    <div class="actions-row">
      <button id="export-report" class="action-button" type="button" disabled>
        Export Excel Report
      </button>
    </div>

    <div class="tab-bar" role="tablist" aria-label="Report Views">
      <button id="tab-mapped" class="tab-button is-active" role="tab" aria-selected="true" aria-controls="panel-mapped">Mapped View</button>
      <button id="tab-calibration" class="tab-button" role="tab" aria-selected="false" aria-controls="panel-calibration">Calibration Review (O)</button>
    </div>

    <section id="panel-mapped" class="results" aria-live="polite" role="tabpanel" aria-labelledby="tab-mapped">
      <h2 id="result-heading">Mapped rows (0)</h2>
      <div class="table-wrap">
        <table id="results-table">
          <thead>
            <tr>
              <th id="header-a">Column A</th>
              <th id="header-b">Column B</th>
              <th id="header-c">Column C</th>
              <th id="header-total">Entries Per Location</th>
              <th id="header-pass-fail">Number Passed/Failed</th>
              <th id="header-analytics">E Averages / Anomalies</th>
            </tr>
          </thead>
          <tbody id="results-body"></tbody>
        </table>
      </div>
    </section>

    <section id="panel-calibration" class="results is-hidden" aria-live="polite" role="tabpanel" aria-labelledby="tab-calibration">
      <h2 id="calibration-heading">Calibration groups (0)</h2>
      <div class="table-wrap">
        <table id="calibration-table">
          <thead>
            <tr>
              <th>O Value</th>
              <th>Rows</th>
              <th>Pass Rate %</th>
              <th>Avg Attempts To Pass</th>
              <th>Avg RLU</th>
              <th>Std Dev</th>
              <th>CV%</th>
              <th>Outlier %</th>
              <th>Sequence Spread %</th>
              <th>Risk %</th>
              <th>Assessment</th>
            </tr>
          </thead>
          <tbody id="calibration-body"></tbody>
        </table>
      </div>
    </section>
  </main>
`

const fileInput = document.querySelector('#excel-file')
const status = document.querySelector('#status')
const resultHeading = document.querySelector('#result-heading')
const headerAElement = document.querySelector('#header-a')
const headerBElement = document.querySelector('#header-b')
const headerCElement = document.querySelector('#header-c')
const headerTotalElement = document.querySelector('#header-total')
const headerPassFailElement = document.querySelector('#header-pass-fail')
const headerAnalyticsElement = document.querySelector('#header-analytics')
const resultBody = document.querySelector('#results-body')
const calibrationHeadingElement = document.querySelector('#calibration-heading')
const calibrationBodyElement = document.querySelector('#calibration-body')
const mappedTabButton = document.querySelector('#tab-mapped')
const calibrationTabButton = document.querySelector('#tab-calibration')
const mappedPanel = document.querySelector('#panel-mapped')
const calibrationPanel = document.querySelector('#panel-calibration')
const exportButton = document.querySelector('#export-report')

let latestMappedValues = null
let latestCalibrationRows = []

const compareAlpha = (a, b) =>
  a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true })

const parseTestNumber = (value) => {
  const raw = String(value ?? '').trim()
  if (!raw) {
    return null
  }
  const parsed = Number(raw)
  return Number.isFinite(parsed) ? parsed : null
}

const isPassingValue = (value) => {
  const normalized = String(value ?? '').trim().toLowerCase()
  return normalized.includes('pass') && !normalized.includes('fail')
}

const compareKLabels = (left, right) => {
  const leftIsPass = isPassingValue(left)
  const rightIsPass = isPassingValue(right)

  if (leftIsPass && !rightIsPass) {
    return -1
  }
  if (!leftIsPass && rightIsPass) {
    return 1
  }

  return compareAlpha(left, right)
}

const getAverageAttemptsToPass = (events) => {
  if (events.length === 0) {
    return { average: null, hasAnyPass: false }
  }

  const sequences = []
  let currentSequence = []

  for (const event of events) {
    if (currentSequence.length === 0) {
      currentSequence.push(event)
      continue
    }

    const previous = currentSequence[currentSequence.length - 1]
    const isConsecutive =
      previous.testNumber !== null &&
      event.testNumber !== null &&
      event.testNumber === previous.testNumber + 1

    if (isConsecutive) {
      currentSequence.push(event)
    } else {
      sequences.push(currentSequence)
      currentSequence = [event]
    }
  }

  if (currentSequence.length > 0) {
    sequences.push(currentSequence)
  }

  const attempts = []
  let hasAnyPass = false

  for (const sequence of sequences) {
    const firstPassIndex = sequence.findIndex((event) => event.isPass)
    if (firstPassIndex !== -1) {
      hasAnyPass = true
      attempts.push(firstPassIndex + 1)
    }
  }

  if (attempts.length === 0) {
    return { average: null, hasAnyPass: false }
  }

  const average = attempts.reduce((sum, value) => sum + value, 0) / attempts.length
  return { average, hasAnyPass }
}

const getRowColorClass = ({ average, hasAnyPass }) => {
  if (!hasAnyPass || average === null) {
    return 'result-row-no-pass'
  }
  if (average > 4) {
    return 'result-row-red'
  }
  if (average > 2) {
    return 'result-row-orange'
  }
  if (average > 1) {
    return 'result-row-yellow'
  }
  return 'result-row-green'
}

const formatPercent = (value) => {
  const rounded = Math.round(value * 100) / 100
  return Number.isInteger(rounded) ? `${rounded}` : `${rounded}`
}

const formatMetric = (value) => {
  const rounded = Math.round(value * 100) / 100
  return Number.isInteger(rounded) ? `${rounded}` : `${rounded}`
}

const formatAverage = (value) => {
  if (value === null) {
    return 'N/A'
  }
  const rounded = Math.round(value * 100) / 100
  return Number.isInteger(rounded) ? `${rounded}` : `${rounded}`
}

const parseNumericValue = (value) => {
  const raw = String(value ?? '').trim()
  if (!raw) {
    return null
  }
  const parsed = Number(raw)
  return Number.isFinite(parsed) ? parsed : null
}

const parsePassingRange = (rawValue) => {
  const raw = String(rawValue ?? '').trim().toLowerCase()
  if (!raw) {
    return null
  }

  const betweenMatch = raw.match(/(-?\d+(?:\.\d+)?)\s*(?:-|to)\s*(-?\d+(?:\.\d+)?)/)
  if (betweenMatch) {
    const first = Number(betweenMatch[1])
    const second = Number(betweenMatch[2])
    if (Number.isFinite(first) && Number.isFinite(second)) {
      return {
        min: Math.min(first, second),
        max: Math.max(first, second)
      }
    }
  }

  const lessOrEqualMatch = raw.match(/^<=\s*(-?\d+(?:\.\d+)?)$/)
  if (lessOrEqualMatch) {
    const max = Number(lessOrEqualMatch[1])
    if (Number.isFinite(max)) {
      return { min: -Infinity, max }
    }
  }

  const lessMatch = raw.match(/^<\s*(-?\d+(?:\.\d+)?)$/)
  if (lessMatch) {
    const max = Number(lessMatch[1])
    if (Number.isFinite(max)) {
      return { min: -Infinity, max: max - Number.EPSILON }
    }
  }

  const greaterOrEqualMatch = raw.match(/^>=\s*(-?\d+(?:\.\d+)?)$/)
  if (greaterOrEqualMatch) {
    const min = Number(greaterOrEqualMatch[1])
    if (Number.isFinite(min)) {
      return { min, max: Infinity }
    }
  }

  const greaterMatch = raw.match(/^>\s*(-?\d+(?:\.\d+)?)$/)
  if (greaterMatch) {
    const min = Number(greaterMatch[1])
    if (Number.isFinite(min)) {
      return { min: min + Number.EPSILON, max: Infinity }
    }
  }

  const singleNumber = Number(raw)
  if (Number.isFinite(singleNumber)) {
    return { min: singleNumber, max: singleNumber }
  }

  return null
}

const isValueInRange = (value, range) => {
  if (value === null || !range) {
    return false
  }
  return value >= range.min && value <= range.max
}

const getMean = (values) => {
  if (values.length === 0) {
    return null
  }
  return values.reduce((sum, value) => sum + value, 0) / values.length
}

const getStandardDeviation = (values) => {
  if (values.length < 2) {
    return 0
  }

  const mean = getMean(values)
  const variance =
    values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / values.length

  return Math.sqrt(variance)
}

const splitEventsIntoSequences = (events) => {
  if (events.length === 0) {
    return []
  }

  const sortedEvents = [...events].sort((a, b) => {
    if (a.testNumber === null && b.testNumber === null) {
      return 0
    }
    if (a.testNumber === null) {
      return 1
    }
    if (b.testNumber === null) {
      return -1
    }
    return a.testNumber - b.testNumber
  })

  const sequences = []
  let currentSequence = []

  for (const event of sortedEvents) {
    if (currentSequence.length === 0) {
      currentSequence.push(event)
      continue
    }

    const previous = currentSequence[currentSequence.length - 1]
    const isConsecutive =
      previous.testNumber !== null &&
      event.testNumber !== null &&
      event.testNumber === previous.testNumber + 1

    if (isConsecutive) {
      currentSequence.push(event)
    } else {
      sequences.push(currentSequence)
      currentSequence = [event]
    }
  }

  if (currentSequence.length > 0) {
    sequences.push(currentSequence)
  }

  return sequences
}

const getPercentile = (sortedValues, percentile) => {
  if (sortedValues.length === 0) {
    return null
  }
  const index = (sortedValues.length - 1) * percentile
  const lowerIndex = Math.floor(index)
  const upperIndex = Math.ceil(index)

  if (lowerIndex === upperIndex) {
    return sortedValues[lowerIndex]
  }

  const weight = index - lowerIndex
  return (
    sortedValues[lowerIndex] * (1 - weight) + sortedValues[upperIndex] * weight
  )
}

const getOutlierValues = (numericValues) => {
  if (numericValues.length < 4) {
    return []
  }

  const sorted = [...numericValues].sort((a, b) => a - b)
  const q1 = getPercentile(sorted, 0.25)
  const q3 = getPercentile(sorted, 0.75)
  const iqr = q3 - q1
  const lowerBound = q1 - 1.5 * iqr
  const upperBound = q3 + 1.5 * iqr

  return numericValues.filter((value) => value < lowerBound || value > upperBound)
}

const getCalibrationRiskSummary = (stats) => {
  let score = 0

  if (stats.passRate < 0.5) {
    score += 20
  } else if (stats.passRate < 0.75) {
    score += 10
  }

  if (stats.avgAttemptsToPass === null) {
    score += 15
  } else if (stats.avgAttemptsToPass > 4) {
    score += 18
  } else if (stats.avgAttemptsToPass > 2) {
    score += 10
  }

  if (stats.cv > 0.5) {
    score += 35
  } else if (stats.cv > 0.35) {
    score += 20
  }

  if (stats.sequenceSpread > 0.6) {
    score += 30
  } else if (stats.sequenceSpread > 0.35) {
    score += 18
  }

  if (stats.outlierRate > 0.25) {
    score += 25
  } else if (stats.outlierRate > 0.12) {
    score += 12
  }

  if (stats.rows >= 6 && stats.stdDev > 0 && stats.avgRlu > 0 && stats.stdDev >= stats.avgRlu * 0.75) {
    score += 10
  }

  const capped = Math.min(100, Math.round(score))

  if (capped >= 75) {
    return { score: capped, label: 'Out of Calibration' }
  }
  if (capped >= 50) {
    return { score: capped, label: 'Unstable' }
  }
  if (capped >= 25) {
    return { score: capped, label: 'Watch' }
  }
  return { score: capped, label: 'Stable' }
}

const buildCalibrationByO = (rows, startRowIndex) => {
  const groups = new Map()

  for (const row of rows.slice(startRowIndex)) {
    const oValue = String(row[14] ?? '').trim()
    const eValue = parseNumericValue(row[4])
    const gValue = parseTestNumber(row[6])
    const kValue = String(row[10] ?? '').trim()

    if (!oValue) {
      continue
    }

    if (!groups.has(oValue)) {
      groups.set(oValue, [])
    }

    groups.get(oValue).push({
      eValue,
      testNumber: gValue,
      isPass: isPassingValue(kValue)
    })
  }

  const result = [...groups.entries()].map(([oValue, events]) => {
    const numericValues = events.map((event) => event.eValue).filter((value) => value !== null)
    const rowsCount = events.length
    const passCount = events.filter((event) => event.isPass).length
    const passRate = rowsCount > 0 ? passCount / rowsCount : 0
    const attemptsStats = getAverageAttemptsToPass(events)
    const avgAttemptsToPass = attemptsStats.average
    const avgRlu = getMean(numericValues)
    const stdDev = getStandardDeviation(numericValues)
    const cv = avgRlu && avgRlu !== 0 ? stdDev / Math.abs(avgRlu) : 0
    const outlierRate =
      numericValues.length > 0
        ? getOutlierValues(numericValues).length / numericValues.length
        : 0

    const sequences = splitEventsIntoSequences(events)
    const sequenceMeans = sequences
      .map((sequence) => getMean(sequence.map((event) => event.eValue).filter((value) => value !== null)))
      .filter((value) => value !== null)

    const maxSeqMean = sequenceMeans.length > 0 ? Math.max(...sequenceMeans) : 0
    const minSeqMean = sequenceMeans.length > 0 ? Math.min(...sequenceMeans) : 0
    const sequenceSpread =
      avgRlu && avgRlu !== 0 ? (maxSeqMean - minSeqMean) / Math.abs(avgRlu) : 0

    const risk = getCalibrationRiskSummary({
      passRate,
      avgAttemptsToPass,
      cv,
      sequenceSpread,
      outlierRate,
      rows: rowsCount,
      stdDev,
      avgRlu: avgRlu ?? 0
    })

    return {
      oValue,
      rows: rowsCount,
      passRate,
      avgAttemptsToPass,
      avgRlu,
      stdDev,
      cv,
      outlierRate,
      sequenceSpread,
      risk
    }
  })

  return result.sort((left, right) => compareAlpha(left.oValue, right.oValue))
}

const getAnomalySummary = (eValues, nonNumericCount, testEvents) => {
  const flags = []

  const hasOnlyFails =
    testEvents.length > 0 && testEvents.every((event) => !event.isPass)

  if (hasOnlyFails) {
    flags.push('All tests failed (no passes)')
  }

  if (nonNumericCount > 0) {
    flags.push(`Non-numeric RLU: ${nonNumericCount}`)
  }

  if (eValues.some((value) => value < 0)) {
    const negativeValues = eValues.filter((value) => value < 0)
    const uniqueNegativeValues = [...new Set(negativeValues)].sort((a, b) => a - b)
    flags.push(`Negative RLU values: ${uniqueNegativeValues.join(', ')}`)
  }

  if (eValues.length >= 4) {
    const sorted = [...eValues].sort((a, b) => a - b)
    const q1 = getPercentile(sorted, 0.25)
    const q3 = getPercentile(sorted, 0.75)
    const iqr = q3 - q1
    const lowerBound = q1 - 1.5 * iqr
    const upperBound = q3 + 1.5 * iqr
    const outlierValues = testEvents
      .filter(
        (event) =>
          event.eValue !== null &&
          (event.eValue < lowerBound || event.eValue > upperBound) &&
          !(event.isPass && isValueInRange(event.eValue, event.passingRange))
      )
      .map((event) => event.eValue)

    if (outlierValues.length > 0) {
      const uniqueOutlierValues = [...new Set(outlierValues)].sort((a, b) => a - b)
      flags.push(
        `Outliers: ${uniqueOutlierValues.join(', ')} (outside IQR bounds ${formatAverage(
          lowerBound
        )} to ${formatAverage(upperBound)})`
      )
    }
  }

  return flags.join('; ')
}

const buildAnalyticsText = (pair) => {
  const overallAverage = getMean(pair.eValues)

  const byKAverage = [...pair.eByK.entries()]
    .sort(([leftKey], [rightKey]) => compareKLabels(leftKey, rightKey))
    .map(([kValue, stats]) => `${kValue}: ${formatAverage(stats.sum / stats.count)}`)
    .join(', ')

  const limitSummary = [...pair.limitValues]
    .sort(compareAlpha)
    .map((value) => `${value} Limit`)
    .join(', ')

  const anomalySummary = getAnomalySummary(
    pair.eValues,
    pair.nonNumericECount,
    pair.testEvents
  )
  const anomalyText = anomalySummary ? ` | Anomalies: ${anomalySummary}` : ''

  return `Overall RLU Avg: ${formatAverage(overallAverage)} | RLU Avg by K: ${
    byKAverage || 'N/A'
  } | Limits: ${limitSummary || 'N/A'}${anomalyText}`
}

const getIntegrityRiskDetails = (testEvents) => {
  const transitions = []

  for (let index = 1; index < testEvents.length; index += 1) {
    const previous = testEvents[index - 1]
    const current = testEvents[index]

    const isConsecutive =
      previous.testNumber !== null &&
      current.testNumber !== null &&
      current.testNumber === previous.testNumber + 1

    if (!isConsecutive) {
      continue
    }

    if (
      !previous.isPass &&
      current.isPass &&
      previous.eValue !== null &&
      current.eValue !== null
    ) {
      transitions.push({
        fromTest: previous.testNumber,
        toTest: current.testNumber,
        failedE: previous.eValue,
        passedE: current.eValue,
        passingRange: current.passingRange
      })
    }
  }

  if (transitions.length === 0) {
    return { percentage: 0, alarmingSequences: [] }
  }

  const alarmingSequences = transitions.filter((transition) => {
    const steepDrop = transition.passedE <= transition.failedE * 0.5
    const passInDeclaredRange = isValueInRange(
      transition.passedE,
      transition.passingRange
    )
    return steepDrop && !passInDeclaredRange
  })

  return {
    percentage: (alarmingSequences.length / transitions.length) * 100,
    alarmingSequences
  }
}

const escapeHtml = (value) =>
  value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')

const getMappedRowsForDisplay = (orderedAValues, groupedValues) => {
  const rows = []

  for (const aValue of orderedAValues) {
    const bAndCValues = [...groupedValues.get(aValue).values()].sort((left, right) => {
      const byB = compareAlpha(left.bValue, right.bValue)
      if (byB !== 0) {
        return byB
      }
      return compareAlpha(left.cValue, right.cValue)
    })

    rows.push({
      type: 'group',
      aValue,
      bValue: '',
      cValue: '',
      entries: '',
      passFail: '',
      analytics: '',
      rowColorClass: ''
    })

    for (const { bValue, cValue, rowCount, kCounts, testEvents, ...pair } of bAndCValues) {
      const passStats = getAverageAttemptsToPass(testEvents)
      const rowColorClass = getRowColorClass(passStats)
      const analyticsText = buildAnalyticsText({ ...pair, testEvents })
      const integrityRisk = getIntegrityRiskDetails(testEvents)
      const alarmingSequenceText = integrityRisk.alarmingSequences
        .map(
          (sequence) =>
            `G${sequence.fromTest}->G${sequence.toTest} (E ${formatMetric(
              sequence.failedE
            )} -> ${formatMetric(sequence.passedE)})`
        )
        .join('; ')
      const riskText =
        integrityRisk.percentage > 0
          ? ` | Data Integrity Risk: ${formatPercent(
              integrityRisk.percentage
            )}% | Alarming Sequences: ${alarmingSequenceText}`
          : ''

      const nonZeroKCounts = [...kCounts.entries()]
        .filter(([, count]) => count > 0)
        .sort(([leftKey], [rightKey]) => compareKLabels(leftKey, rightKey))

      const kBreakdown =
        nonZeroKCounts.length === 1
          ? `${nonZeroKCounts[0][0]}: ${formatPercent(
              (nonZeroKCounts[0][1] / rowCount) * 100
            )}%`
          : nonZeroKCounts
              .map(
                ([key, count]) =>
                  `${key}: ${formatPercent((count / rowCount) * 100)}%`
              )
              .join(', ')

      rows.push({
        type: 'result',
        aValue: '',
        bValue,
        cValue,
        entries: String(rowCount),
        passFail: kBreakdown || 'No values',
        analytics: `${analyticsText}${riskText}`,
        rowColorClass
      })
    }
  }

  return rows
}

const getCalibrationRowsForDisplay = (calibrationRows) =>
  calibrationRows.map((row) => ({
    oValue: row.oValue,
    rows: String(row.rows),
    passRate: `${formatPercent(row.passRate * 100)}%`,
    avgAttemptsToPass:
      row.avgAttemptsToPass === null ? 'No pass' : formatAverage(row.avgAttemptsToPass),
    avgRlu: formatAverage(row.avgRlu),
    stdDev: formatAverage(row.stdDev),
    cv: `${formatPercent(row.cv * 100)}%`,
    outlierRate: `${formatPercent(row.outlierRate * 100)}%`,
    sequenceSpread: `${formatPercent(row.sequenceSpread * 100)}%`,
    risk: `${row.risk.score}%`,
    assessment: row.risk.label
  }))

const getAutoColumnWidth = ({ headerText, values, min = 12, max = 80, padding = 2 }) => {
  const widestValueLength = values.reduce((widest, value) => {
    const currentLength = String(value ?? '').length
    return Math.max(widest, currentLength)
  }, String(headerText ?? '').length)

  return Math.max(min, Math.min(max, widestValueLength + padding))
}

const getExportFileName = (defaultName) => {
  const userValue = window.prompt('Enter Excel export file name', defaultName)
  if (!userValue) {
    return null
  }
  const trimmed = userValue.trim()
  if (!trimmed) {
    return null
  }
  return trimmed.toLowerCase().endsWith('.xlsx') ? trimmed : `${trimmed}.xlsx`
}

const saveBufferWithDialog = async (buffer, suggestedName) => {
  if ('showSaveFilePicker' in window) {
    const handle = await window.showSaveFilePicker({
      suggestedName,
      types: [
        {
          description: 'Excel Workbook',
          accept: {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
          }
        }
      ]
    })
    const writable = await handle.createWritable()
    await writable.write(buffer)
    await writable.close()
    return
  }

  const chosenName = getExportFileName(suggestedName)
  if (!chosenName) {
    throw new Error('Export canceled')
  }

  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  })
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = chosenName
  document.body.appendChild(anchor)
  anchor.click()
  document.body.removeChild(anchor)
  URL.revokeObjectURL(url)
}

const applyWorksheetBaseStyle = (sheet) => {
  sheet.views = [{ state: 'frozen', ySplit: 1 }]

  const headerRow = sheet.getRow(1)
  headerRow.font = { bold: true, color: { argb: 'FF111827' } }
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF8EF' }
  }
  headerRow.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }
  headerRow.height = 24

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return
    }
    row.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }
  })
}

const getEstimatedWrappedLineCount = (text, columnWidth) => {
  const value = String(text ?? '')
  if (!value) {
    return 1
  }

  const effectiveWidth = Math.max(1, Math.floor((columnWidth ?? 12) - 1))
  return value
    .split('\n')
    .reduce(
      (total, segment) => total + Math.max(1, Math.ceil(segment.length / effectiveWidth)),
      0
    )
}

const applyMappedRowHeights = (sheet, mappedRows) => {
  const passFailWidth = sheet.getColumn(5).width ?? 40
  const analyticsWidth = sheet.getColumn(6).width ?? 72

  mappedRows.forEach((row, index) => {
    const excelRow = sheet.getRow(index + 2)

    if (row.type === 'group') {
      excelRow.height = 22
      return
    }

    const passFailLines = getEstimatedWrappedLineCount(row.passFail, passFailWidth)
    const analyticsLines = getEstimatedWrappedLineCount(row.analytics, analyticsWidth)
    const lineCount = Math.max(1, passFailLines, analyticsLines)
    const estimatedHeight = 18 + (lineCount - 1) * 14

    excelRow.height = Math.min(220, Math.max(20, estimatedHeight))
  })
}

const applyMappedRowColor = (sheet, excelRowIndex, rowType, rowColorClass) => {
  if (rowType !== 'result') {
    return
  }

  const colorByClass = {
    'result-row-no-pass': 'FFDA9694',
    'result-row-red': 'FFFF0000',
    'result-row-orange': 'FFFFA500',
    'result-row-yellow': 'FFFFFF00',
    'result-row-green': 'FF00A651'
  }

  const fillColor = colorByClass[rowColorClass]
  if (!fillColor) {
    return
  }

  for (let column = 2; column <= 5; column += 1) {
    sheet.getCell(excelRowIndex, column).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: fillColor }
    }
  }
}

const addMappedColorKeyInColumnH = (sheet) => {
  sheet.getColumn(8).width = 42

  const legendItems = [
    { row: 2, color: 'FFDA9694', label: 'Key: No pass in sequence' },
    { row: 3, color: 'FFFF0000', label: 'Key: Average attempts > 4' },
    { row: 4, color: 'FFFFA500', label: 'Key: Average attempts > 2' },
    { row: 5, color: 'FFFFFF00', label: 'Key: Average attempts > 1' },
    { row: 6, color: 'FF00A651', label: 'Key: Average attempts = 1' }
  ]

  legendItems.forEach((item) => {
    const cell = sheet.getCell(item.row, 8)
    cell.value = item.label
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: item.color }
    }
    cell.font = { bold: true, color: { argb: 'FF111827' } }
    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true }
    cell.border = {
      bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } }
    }
  })
}

const exportReportWorkbook = async () => {
  if (!latestMappedValues || latestMappedValues.orderedAValues.length === 0) {
    status.textContent = 'Load a file before exporting.'
    return
  }

  exportButton.disabled = true
  status.textContent = 'Preparing styled Excel export...'

  try {
    const { default: ExcelJS } = await import('exceljs')
    const workbook = new ExcelJS.Workbook()
    workbook.creator = 'Excel Report Generator'
    workbook.created = new Date()

    const mappedSheet = workbook.addWorksheet('Mapped View')
    const mappedRows = getMappedRowsForDisplay(
      latestMappedValues.orderedAValues,
      latestMappedValues.groupedValues
    )
    const mappedColumnBWidth = getAutoColumnWidth({
      headerText: latestMappedValues.headerB,
      values: mappedRows.map((row) => row.bValue),
      min: 22,
      max: 120,
      padding: 3
    })

    mappedSheet.columns = [
      { header: latestMappedValues.headerA, key: 'aValue', width: 26 },
      { header: latestMappedValues.headerB, key: 'bValue', width: mappedColumnBWidth },
      { header: latestMappedValues.headerC, key: 'cValue', width: 22 },
      { header: 'Entries Per Location', key: 'entries', width: 20 },
      { header: 'Number Passed/Failed', key: 'passFail', width: 40 },
      { header: 'RLU Averages / Limits / Anomalies', key: 'analytics', width: 72 }
    ]

    for (const row of mappedRows) {
      const added = mappedSheet.addRow({
        aValue: row.aValue,
        bValue: row.bValue,
        cValue: row.cValue,
        entries: row.entries,
        passFail: row.passFail,
        analytics: row.analytics
      })

      if (row.type === 'group') {
        added.getCell(1).font = { bold: true }
      }

      applyMappedRowColor(mappedSheet, added.number, row.type, row.rowColorClass)
    }

    applyWorksheetBaseStyle(mappedSheet)
    applyMappedRowHeights(mappedSheet, mappedRows)
    mappedSheet.getColumn(6).alignment = {
      wrapText: true,
      vertical: 'middle',
      horizontal: 'center'
    }
    mappedSheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } }
        }
      })
    })
    addMappedColorKeyInColumnH(mappedSheet)

    const calibrationSheet = workbook.addWorksheet('Calibration Review')
    const calibrationDisplayRows = getCalibrationRowsForDisplay(latestCalibrationRows)

    calibrationSheet.columns = [
      { header: 'O Value', key: 'oValue', width: 24 },
      { header: 'Rows', key: 'rows', width: 12 },
      { header: 'Pass Rate %', key: 'passRate', width: 14 },
      { header: 'Avg Attempts To Pass', key: 'avgAttemptsToPass', width: 22 },
      { header: 'Avg RLU', key: 'avgRlu', width: 14 },
      { header: 'Std Dev', key: 'stdDev', width: 14 },
      { header: 'CV%', key: 'cv', width: 12 },
      { header: 'Outlier %', key: 'outlierRate', width: 14 },
      { header: 'Sequence Spread %', key: 'sequenceSpread', width: 18 },
      { header: 'Risk %', key: 'risk', width: 12 },
      { header: 'Assessment', key: 'assessment', width: 20 }
    ]

    if (calibrationDisplayRows.length === 0) {
      calibrationSheet.addRow({ oValue: 'No non-empty O values found.' })
      calibrationSheet.mergeCells('A2:K2')
    } else {
      calibrationDisplayRows.forEach((row) => calibrationSheet.addRow(row))
    }

    applyWorksheetBaseStyle(calibrationSheet)
    calibrationSheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.border = {
          bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } }
        }
      })
    })

    const buffer = await workbook.xlsx.writeBuffer()
    const suggestedName = `excel-report-${new Date().toISOString().slice(0, 10)}.xlsx`
    await saveBufferWithDialog(buffer, suggestedName)

    status.textContent = 'Excel report exported successfully.'
  } catch (error) {
    if (error instanceof Error && error.message === 'Export canceled') {
      status.textContent = 'Export canceled.'
    } else {
      console.error(error)
      status.textContent = 'Export failed. Please try again.'
    }
  } finally {
    exportButton.disabled = false
  }
}

const mapColumnAtoBAndCValues = (worksheet) => {
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    raw: false,
    defval: ''
  })

  const defaultHeaderA = 'Column A'
  const defaultHeaderB = 'Column B'
  const defaultHeaderC = 'Column C'
  const firstDataLikeRowIndex = rows.findIndex((row) => {
    const aValue = String(row[0] ?? '').trim()
    const bValue = String(row[1] ?? '').trim()
    const cValue = String(row[2] ?? '').trim()
    return aValue || bValue || cValue
  })

  const headerRow =
    firstDataLikeRowIndex >= 0 ? rows[firstDataLikeRowIndex] : []
  const headerA = String(headerRow[0] ?? '').trim() || defaultHeaderA
  const headerB = String(headerRow[1] ?? '').trim() || defaultHeaderB
  const headerC = String(headerRow[2] ?? '').trim() || defaultHeaderC

  const groupedValues = new Map()

  const startRowIndex =
    firstDataLikeRowIndex >= 0 ? firstDataLikeRowIndex + 1 : rows.length

  for (const row of rows.slice(startRowIndex)) {
    const aValue = String(row[0] ?? '').trim()
    const bValue = String(row[1] ?? '').trim()
    const cValue = String(row[2] ?? '').trim()
    const eValue = parseNumericValue(row[4])
    const limitValue = String(row[5] ?? '').trim()
    const passingRange = parsePassingRange(limitValue)
    const gValue = parseTestNumber(row[6])
    const kValue = String(row[10] ?? '').trim()

    if (!aValue) {
      continue
    }

    if (!groupedValues.has(aValue)) {
      groupedValues.set(aValue, new Map())
    }

    if (bValue || cValue) {
      const pairKey = `${bValue}\u0000${cValue}`
      const currentPair = groupedValues.get(aValue).get(pairKey)

      if (currentPair) {
        currentPair.rowCount += 1
        currentPair.testEvents.push({
          testNumber: gValue,
          isPass: isPassingValue(kValue),
          eValue,
          passingRange
        })
        if (eValue === null) {
          currentPair.nonNumericECount += 1
        } else {
          currentPair.eValues.push(eValue)
          const groupKey = kValue || '(blank)'
          const currentKStats = currentPair.eByK.get(groupKey) ?? { sum: 0, count: 0 }
          currentKStats.sum += eValue
          currentKStats.count += 1
          currentPair.eByK.set(groupKey, currentKStats)
        }
        if (limitValue) {
          currentPair.limitValues.add(limitValue)
        }
        if (kValue) {
          currentPair.kCounts.set(kValue, (currentPair.kCounts.get(kValue) ?? 0) + 1)
        }
      } else {
        const kCounts = new Map()
        const eValues = []
        const eByK = new Map()

        if (eValue !== null) {
          eValues.push(eValue)
          const groupKey = kValue || '(blank)'
          eByK.set(groupKey, { sum: eValue, count: 1 })
        }

        if (kValue) {
          kCounts.set(kValue, 1)
        }

        groupedValues.get(aValue).set(pairKey, {
          bValue,
          cValue,
          rowCount: 1,
          testEvents: [
            {
              testNumber: gValue,
              isPass: isPassingValue(kValue),
              eValue,
              passingRange
            }
          ],
          kCounts,
          eValues,
          eByK,
          limitValues: new Set(limitValue ? [limitValue] : []),
          nonNumericECount: eValue === null ? 1 : 0
        })
      }
    }
  }

  const orderedAValues = [...groupedValues.keys()].sort(compareAlpha)
  const totalMappedRows = [...groupedValues.values()].reduce(
    (count, valueMap) => count + valueMap.size,
    0
  )

  return {
    headerA,
    headerB,
    headerC,
    orderedAValues,
    groupedValues,
    totalMappedRows,
    rows,
    startRowIndex
  }
}

const renderCalibrationResults = (calibrationRows) => {
  calibrationHeadingElement.textContent = `Calibration groups (${calibrationRows.length})`

  if (calibrationRows.length === 0) {
    calibrationBodyElement.innerHTML = `
      <tr>
        <td colspan="11">No non-empty O values found.</td>
      </tr>
    `
    return
  }

  calibrationBodyElement.innerHTML = calibrationRows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.oValue)}</td>
          <td>${row.rows}</td>
          <td>${formatPercent(row.passRate * 100)}%</td>
          <td>${row.avgAttemptsToPass === null ? 'No pass' : formatAverage(row.avgAttemptsToPass)}</td>
          <td>${formatAverage(row.avgRlu)}</td>
          <td>${formatAverage(row.stdDev)}</td>
          <td>${formatPercent(row.cv * 100)}%</td>
          <td>${formatPercent(row.outlierRate * 100)}%</td>
          <td>${formatPercent(row.sequenceSpread * 100)}%</td>
          <td>${row.risk.score}%</td>
          <td>${escapeHtml(row.risk.label)}</td>
        </tr>
      `
    )
    .join('')
}

const setActiveTab = (tabName) => {
  const showMapped = tabName === 'mapped'
  mappedTabButton.classList.toggle('is-active', showMapped)
  mappedTabButton.setAttribute('aria-selected', String(showMapped))
  calibrationTabButton.classList.toggle('is-active', !showMapped)
  calibrationTabButton.setAttribute('aria-selected', String(!showMapped))

  mappedPanel.classList.toggle('is-hidden', !showMapped)
  calibrationPanel.classList.toggle('is-hidden', showMapped)
}

mappedTabButton.addEventListener('click', () => setActiveTab('mapped'))
calibrationTabButton.addEventListener('click', () => setActiveTab('calibration'))

const renderMappedResults = ({ headerA, headerB, headerC, orderedAValues, groupedValues }) => {
  headerAElement.textContent = headerA
  headerBElement.textContent = headerB
  headerCElement.textContent = headerC
  headerTotalElement.textContent = 'Entries Per Location'
  headerPassFailElement.textContent = 'Number Passed/Failed'
  headerAnalyticsElement.textContent = 'RLU Averages / Limits / Anomalies'
  resultHeading.textContent = `Mapped rows (${orderedAValues.length})`

  if (orderedAValues.length === 0) {
    resultBody.innerHTML = `
      <tr>
        <td colspan="6">No non-empty values found in column A.</td>
      </tr>
    `
    return
  }

  const tableRows = getMappedRowsForDisplay(orderedAValues, groupedValues).map((row) => {
    if (row.type === 'group') {
      return `
        <tr class="group-row">
          <td>${escapeHtml(row.aValue)}</td>
          <td></td>
          <td></td>
          <td></td>
          <td></td>
          <td></td>
        </tr>
      `
    }

    return `
      <tr class="${row.rowColorClass}">
        <td></td>
        <td>${escapeHtml(row.bValue)}</td>
        <td>${escapeHtml(row.cValue)}</td>
        <td>${row.entries}</td>
        <td>${escapeHtml(row.passFail)}</td>
        <td class="analytics-cell">${escapeHtml(row.analytics)}</td>
      </tr>
    `
  })

  resultBody.innerHTML = tableRows.join('')
}

exportButton.addEventListener('click', exportReportWorkbook)

fileInput.addEventListener('change', async (event) => {
  const file = event.target.files?.[0]

  if (!file) {
    status.textContent = 'No file selected.'
    latestMappedValues = null
    latestCalibrationRows = []
    exportButton.disabled = true
    renderMappedResults({
      headerA: 'Column A',
      headerB: 'Column B',
      headerC: 'Column C',
      orderedAValues: [],
      groupedValues: new Map()
    })
    renderCalibrationResults([])
    return
  }

  status.textContent = `Reading ${file.name}...`

  try {
    const fileData = await file.arrayBuffer()
    const workbook = XLSX.read(fileData, { type: 'array' })
    const firstSheetName = workbook.SheetNames[0]

    if (!firstSheetName) {
      status.textContent = 'No worksheet found in this file.'
      latestMappedValues = null
      latestCalibrationRows = []
      exportButton.disabled = true
      renderMappedResults({
        headerA: 'Column A',
        headerB: 'Column B',
        headerC: 'Column C',
        orderedAValues: [],
        groupedValues: new Map()
      })
      renderCalibrationResults([])
      return
    }

    const worksheet = workbook.Sheets[firstSheetName]
    const mappedValues = mapColumnAtoBAndCValues(worksheet)
    const calibrationRows = buildCalibrationByO(
      mappedValues.rows,
      mappedValues.startRowIndex
    )

    latestMappedValues = mappedValues
    latestCalibrationRows = calibrationRows
    exportButton.disabled = mappedValues.orderedAValues.length === 0

    renderMappedResults(mappedValues)
    renderCalibrationResults(calibrationRows)
    status.textContent = `Loaded ${mappedValues.orderedAValues.length} unique A values and ${mappedValues.totalMappedRows} unique mapped B/C rows from ${firstSheetName}.`
  } catch (error) {
    console.error(error)
    status.textContent = 'Could not parse file. Please upload a valid Excel-importable file.'
    latestMappedValues = null
    latestCalibrationRows = []
    exportButton.disabled = true
    renderMappedResults({
      headerA: 'Column A',
      headerB: 'Column B',
      headerC: 'Column C',
      orderedAValues: [],
      groupedValues: new Map()
    })
    renderCalibrationResults([])
  }
})
