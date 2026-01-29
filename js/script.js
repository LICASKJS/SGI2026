const DEFAULT_WORKBOOK_ID = "indicadores"
const WORKBOOK_PATHS = {
  [DEFAULT_WORKBOOK_ID]: "dados/TESTE PARA INDICADORES - CABEÇALHOS.xlsx",
  remGeral: "dados/REM GERAL.xlsx",
}

const STORAGE_KEYS = {
  data: (workbookId) => `sgiWorkbookData:${workbookId}`,
  source: (workbookId) => `sgiWorkbookSource:${workbookId}`,
  selectedSheet: (pageId) => `sgiSelectedSheet:${pageId}`,
}

const pages = [
  {
    id: "acidentes-c-afastamento",
    sheet: "ACIDENTES C AFASTAMENTO",
    sheetLabel: "ACIDENTES COM AFASTAMENTO",
    sheetAliases: ["ACIDENTES COM AFASTAMENTO"],
    file: "index.html",
    navLabel: "Acidentes com Afastamento",
    icon: "shield",
    heroTitle: "Acidentes com Afastamento",
    view: "both",
    description: "Registros oficiais de acidentes com afastamento informados na planilha corporativa.",
  },
  {
    id: "acidentes-s-afastamento",
    sheet: "ACIDENTES S AFASTAMENTO",
    sheetLabel: "ACIDENTES SEM AFASTAMENTO",
    sheetAliases: ["ACIDENTES SEM AFASTAMENTO"],
    file: "racs.html",
    navLabel: "Acidentes sem Afastamento",
    icon: "alert",
    heroTitle: "Acidentes sem Afastamento",
    view: "table",
    description: "Eventos de acidentes sem afastamento categorizados exatamente como na planilha.",
  },
  {
    id: "aud-comp",
    sheet: "AUD COMP",
    file: "racs-dentro-prazo.html",
    navLabel: "Auditoria Comportamental",
    icon: "clipboard",
    heroTitle: "Auditoria Comportamental",
    view: "chart",
    description: "Resumo das auditorias de comportamento (AUD COMP) com todos os desvios identificados.",
  },
  {
    id: "garantia-vida",
    sheet: "GARANTIA DE QUA DE VIDA",
    file: "racs-fora-prazo.html",
    navLabel: "Garantia de Qualidade de Vida",
    icon: "heart",
    heroTitle: "Garantia de Qualidade de Vida",
    view: "both",
    description: "Indicadores de garantia de qualidade de vida, com os mesmos termos utilizados na planilha.",
  },
  {
    id: "acoes-corretivas",
    sheet: "AÇÕES CORRETIVAS",
    file: "racs-vencidas.html",
    navLabel: "Ações Corretivas",
    icon: "check",
    heroTitle: "Ações Corretivas",
    view: "table",
    description: "Painel das ações corretivas (RACs) e seus prazos conforme cadastro oficial.",
  },
  {
    id: "programas",
    sheet: "PROGRAMAS",
    file: "analise.html",
    navLabel: "Programas Corporativos",
    icon: "layers",
    heroTitle: "Programas Corporativos",
    view: "both",
    description: "Acompanhamento dos programas ativos, seguindo a estrutura da planilha.",
  },
  {
    id: "campanhas",
    sheet: "CAMPANHAS",
    file: "campanhas.html",
    navLabel: "Campanhas Institucionais",
    icon: "megaphone",
    heroTitle: "Campanhas Institucionais",
    view: "chart",
    description: "Visão das campanhas ativas exatamente como descrito no documento base.",
  },
  {
    id: "geri-geri-co",
    sheet: "GERI  GERI CO",
    file: "geri-geri-co.html",
    navLabel: "GERI e GERI-CO",
    icon: "target",
    heroTitle: "GERI e GERI-CO",
    view: "chart",
    description: "Indicadores de GERI e GERI-CO respeitando a nomenclatura original da planilha.",
  },
  {
    id: "rem-geral",
    workbookId: "remGeral",
    sheet: "RPBC",
    file: "rem-geral.html",
    navLabel: "REM GERAL",
    icon: "file",
    heroTitle: "REM GERAL",
    tableTitle: "REM GERAL",
    sheetSelector: "auto",
    view: "both",
    description: "",
  },
]

const navGroups = [
  {
    id: "acidentes",
    label: "Indicadores de Segurança",
    pages: ["acidentes-c-afastamento", "acidentes-s-afastamento"],
  },
  {
    id: "avaliacoes",
    label: "Auditoria",
    pages: ["aud-comp"],
  },
  {
    id: "gestao",
    label: "RACs e Programas",
    pages: ["acoes-corretivas", "programas"],
  },
  {
    id: "iniciativas",
    label: "Iniciativas",
    pages: ["campanhas", "geri-geri-co"],
  },
  {
    id: "anexos",
    label: "REM GERAL",
    pages: ["rem-geral"],
  },

  {
    id: "avaliacoes",
    label: "Saúde",
    pages: ["garantia-vida"],
  }
]

const iconTemplates = {
  shield: '<path d="M12 3l7 4v5c0 4.5-3 8.5-7 9s-7-4.5-7-9V7z" />',
  alert: '<path d="M12 3l9 16H3z" /><path d="M12 10v3" /><circle cx="12" cy="17" r="0.8" />',
  clipboard: '<rect x="6" y="5" width="12" height="15" rx="2" /><path d="M9 3h6v4H9z" />',
  heart: '<path d="M12 20s-6-4-6-9a4 4 0 0 1 7-2.5A4 4 0 0 1 18 11c0 5-6 9-6 9z" />',
  check: '<path d="M5 13l4 4L19 7" />',
  layers: '<path d="M12 5l9 4-9 4-9-4 9-4z" /><path d="M3 13l9 4 9-4" /><path d="M3 17l9 4 9-4" />',
  megaphone: '<path d="M4 11v2l5 2V9z" /><path d="M9 9l11-3v12l-11-3" /><path d="M13 6v12" />',
  target: '<circle cx="12" cy="12" r="8" /><circle cx="12" cy="12" r="4" /><circle cx="12" cy="12" r="1" />',
  file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><path d="M14 2v6h6" />',
  default: '<circle cx="12" cy="12" r="6" />',
}

function getIconSvg(name) {
  const content = iconTemplates[name] || iconTemplates.default
  return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${content}</svg>`
}


const pageMap = pages.reduce((acc, page) => {
  acc[page.id] = page
  return acc
}, {})

const workbookCache = new Map()
const workbookSource = new Map()
let chartInstance = null

function formatNavLabel(text) {
  if (!text) return ""
  return text
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " ")
    .replace(/\b(\w)/g, (match) => match.toUpperCase())
}

function escapeHtml(str) {
  return str.replace(/[&<>"]/g, (char) => {
    switch (char) {
      case "&":
        return "&amp;"
      case "<":
        return "&lt;"
      case ">":
        return "&gt;"
      case '"':
        return "&quot;"
      default:
        return char
    }
  })
}

function highlightSecondWord(text) {
  if (!text) return ""
  const words = text.split(/\s+/).filter(Boolean)
  if (!words.length) return ""
  return words
    .map((word, index) => {
      const safe = escapeHtml(word)
      if (index === 1) {
        return `<span class="gradient-text">${safe}</span>`
      }
      return safe
    })
    .join(" ")
}

document.addEventListener("DOMContentLoaded", () => {
  const pageId = document.body?.dataset?.page || "acidentes-c-afastamento"
  const pageInfo = pageMap[pageId]

  renderNavigation(pageId)

  if (!pageInfo) {
    return
  }

  updatePageCopy(pageInfo)
  updatePageLayout(pageInfo)
  setupWorkbookUpload(pageInfo)
  const workbookId = getWorkbookId(pageInfo)
  loadWorkbook(workbookId)
    .then((workbook) => {
      const sheetName = setupSheetSelector(pageInfo, workbook, workbookId) || pageInfo.sheet
      const runtimePage = buildRuntimePage(pageInfo, sheetName)
      const tableData = renderSheetTable(runtimePage, workbook, { source: getWorkbookSource(workbookId) })
      renderSheetChart(runtimePage, tableData)
    })
    .catch((error) => showError(error))
})

function renderNavigation(currentId) {
  const currentGroup = navGroups.find((group) => group.pages.includes(currentId))
  renderPrimaryNavigation(currentGroup?.id)
  renderSideNavigation(currentGroup, currentId)
}

function renderPrimaryNavigation(activeGroupId) {
  const container = document.querySelector("[data-nav-primary]")
  if (!container) return

  const markup = navGroups
    .map((group) => {
      const firstPage = group.pages[0]
      const target = pageMap[firstPage]
      if (!target) return ""
      const isActive = group.id === activeGroupId
      return `<a href="${target.file}" class="nav-link${isActive ? " active" : ""}">${escapeHtml(group.label)}</a>`
    })
    .join("")

  container.innerHTML = markup
}

function renderSideNavigation(group, currentPageId) {
  const container = document.querySelector("[data-side-nav]")
  if (!container) return

  if (!group) {
    container.innerHTML = ""
    container.classList.add("is-hidden")
    container.style.display = "none"
    return
  }

  container.classList.remove("is-hidden")
  container.style.display = ""
  container.innerHTML = group.pages
    .map((pageId) => {
      const page = pageMap[pageId]
      if (!page) return ""
      const isActive = pageId === currentPageId
      const icon = getIconSvg(page.icon)
      return `<a href="${page.file}" class="side-nav-link${isActive ? " active" : ""}">
        <span class="side-nav-icon">${icon}</span>
        <span class="side-nav-label">${escapeHtml(page.navLabel || formatNavLabel(page.sheet))}</span>
      </a>`
    })
    .join("")
}

function updatePageCopy(page) {
  document
    .querySelectorAll("[data-sheet-title]")
    .forEach((node) => (node.innerHTML = highlightSecondWord(page.heroTitle || page.sheet)))
  document.querySelectorAll("[data-sheet-description]").forEach((node) => {
    const description = page.description || ""
    node.textContent = description
    node.style.display = description ? "" : "none"
  })
  document
    .querySelectorAll("[data-table-title]")
    .forEach((node) => (node.textContent = page.tableTitle || page.sheetLabel || page.sheet))
}

function getWorkbookId(pageInfo) {
  return pageInfo?.workbookId || DEFAULT_WORKBOOK_ID
}

function getWorkbookPath(workbookId) {
  return WORKBOOK_PATHS[workbookId]
}

function getWorkbookSource(workbookId) {
  const source = workbookSource.get(workbookId)
  if (source) return source

  const workbookPath = getWorkbookPath(workbookId)
  return workbookPath?.split("/").pop() || "arquivo padrão"
}

async function loadWorkbook(workbookId) {
  if (workbookCache.has(workbookId)) {
    return workbookCache.get(workbookId)
  }

  const workbookPath = getWorkbookPath(workbookId)
  if (!workbookPath) {
    throw new Error(`Planilha não configurada: ${workbookId}`)
  }

  const stored = loadWorkbookFromStorage(workbookId)
  if (stored) {
    workbookCache.set(workbookId, stored.workbook)
    workbookSource.set(workbookId, stored.source)
    return stored.workbook
  }

  const response = await fetch(workbookPath)
  if (!response.ok) {
    throw new Error(`Não foi possível carregar a planilha "${workbookPath}".`)
  }

  const buffer = await response.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: "array" })
  workbookCache.set(workbookId, workbook)
  workbookSource.set(workbookId, workbookPath.split("/").pop() || "arquivo padrão")
  return workbook
}

function setupSheetSelector(pageInfo, workbook, workbookId) {
  const select = document.getElementById("sheetSelect")
  if (!select || !pageInfo?.sheetSelector || !workbook) return null

  const availableSheets =
    pageInfo.sheetSelector === "auto" ? workbook.SheetNames : Array.isArray(pageInfo.sheetSelector) ? pageInfo.sheetSelector : []

  select.innerHTML = ""
  availableSheets.forEach((sheetName) => {
    const option = document.createElement("option")
    option.value = sheetName
    option.textContent = sheetName
    select.appendChild(option)
  })

  const stored = localStorage.getItem(STORAGE_KEYS.selectedSheet(pageInfo.id))
  const initialSheet = stored && availableSheets.includes(stored) ? stored : pageInfo.sheet || availableSheets[0]
  if (initialSheet) {
    select.value = initialSheet
  }

  select.onchange = () => {
    const chosen = select.value
    localStorage.setItem(STORAGE_KEYS.selectedSheet(pageInfo.id), chosen)
    const runtimePage = buildRuntimePage(pageInfo, chosen)
    const tableData = renderSheetTable(runtimePage, workbook, { source: getWorkbookSource(workbookId) })
    renderSheetChart(runtimePage, tableData)
  }

  return initialSheet
}

function buildRuntimePage(pageInfo, sheetName) {
  if (!pageInfo || !sheetName) return pageInfo
  if (!pageInfo.sheetSelector) return pageInfo
  return { ...pageInfo, sheet: sheetName, sheetLabel: sheetName, chartLabel: sheetName }
}

function normalizeSheetKey(name) {
  if (!name) return ""
  return String(name)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim()
}

function resolveSheetName(page, workbook) {
  if (!page?.sheet || !workbook?.Sheets) return page?.sheet
  if (workbook.Sheets[page.sheet]) return page.sheet

  const sheetNames = workbook.SheetNames || Object.keys(workbook.Sheets)
  const normalizedMap = new Map()
  sheetNames.forEach((sheetName) => normalizedMap.set(normalizeSheetKey(sheetName), sheetName))

  const candidates = [page.sheet, page.sheetLabel, ...(page.sheetAliases || [])].filter(Boolean)
  for (const candidate of candidates) {
    const resolved = normalizedMap.get(normalizeSheetKey(candidate))
    if (resolved) return resolved
  }

  return page.sheet
}

function detectNumericColumns(headers, rows) {
  if (!headers?.length || !rows?.length) return headers.map(() => false)

  const sample = rows.slice(0, 40)
  return headers.map((_, columnIndex) => {
    let sampleCount = 0
    let numericCount = 0

    sample.forEach((row) => {
      const value = row[columnIndex]
      if (!hasValue(value)) return
      sampleCount += 1
      if (parseNumeric(value) !== null) {
        numericCount += 1
      }
    })

    if (!sampleCount) return false
    return numericCount / sampleCount >= 0.75
  })
}

function renderSheetTable(page, workbook, options = {}) {
  const table = document.getElementById("dataTable")
  const footnote = document.getElementById("tableFootnote")
  if (!table) return

  const sheetName = resolveSheetName(page, workbook)
  const worksheet = workbook.Sheets[sheetName]
  if (!worksheet) {
    table.innerHTML = ""
    if (footnote) footnote.textContent = `A aba "${page.sheet}" não foi encontrada na planilha.`
    return
  }

  const parsed = parseWorksheetTable(worksheet)
  const headers = parsed.headers
  const dataRows = parsed.rows
  const numericColumns = detectNumericColumns(headers, dataRows)

  if (!headers.length) {
    table.innerHTML = ""
    if (footnote) footnote.textContent = `A aba "${page.sheet}" não possui cabeçalho definido.`
    return null
  }

  const thead = document.createElement("thead")
  const headRow = document.createElement("tr")
  headers.forEach((header, index) => {
    const th = document.createElement("th")
    th.textContent = header
    if (numericColumns[index]) {
      th.classList.add("is-numeric")
    }
    headRow.appendChild(th)
  })
  thead.appendChild(headRow)

  const tbody = document.createElement("tbody")
  if (dataRows.length) {
    dataRows.forEach((row) => {
      const tr = document.createElement("tr")
      headers.forEach((_, index) => {
        const td = document.createElement("td")
        const cellValue = row[index]
        td.textContent = cellValue === undefined || cellValue === null ? "" : cellValue
        if (numericColumns[index]) {
          td.classList.add("is-numeric")
        }
        tr.appendChild(td)
      })
      tbody.appendChild(tr)
    })
  } else {
    const tr = document.createElement("tr")
    const td = document.createElement("td")
    td.colSpan = headers.length
    td.textContent = "Nenhum registro disponível nesta aba."
    td.style.textAlign = "center"
    tr.appendChild(td)
    tbody.appendChild(tr)
  }

  table.innerHTML = ""
  table.appendChild(thead)
  table.appendChild(tbody)

  if (footnote) {
    const sourceLabel = options.source || getWorkbookSource(getWorkbookId(page)) || "arquivo padrão"
    const sheetLabel = page.sheetLabel || sheetName || page.sheet
    footnote.textContent = `Fonte: ${sourceLabel} • Aba: ${sheetLabel} • ${dataRows.length} registros • ${headers.length} colunas`
  }

  return { headers, rows: dataRows }
}

function setupWorkbookUpload(pageInfo) {
  const input = document.getElementById("workbookUpload")
  if (!input || !pageInfo) return

  input.addEventListener("change", async (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    setUploadState(true)
    try {
      const buffer = await file.arrayBuffer()
      const workbookId = getWorkbookId(pageInfo)
      const workbook = XLSX.read(buffer, { type: "array" })

      workbookCache.set(workbookId, workbook)
      workbookSource.set(workbookId, file.name)
      persistWorkbook(workbookId, workbook, file.name)

      const sheetName = setupSheetSelector(pageInfo, workbook, workbookId) || pageInfo.sheet
      const runtimePage = buildRuntimePage(pageInfo, sheetName)
      const tableData = renderSheetTable(runtimePage, workbook, { source: file.name })
      renderSheetChart(runtimePage, tableData)
    } catch (error) {
      showError(error)
    } finally {
      setUploadState(false)
      input.value = ""
    }
  })
}

function setUploadState(isLoading) {
  document.querySelectorAll("[data-upload-trigger]").forEach((trigger) => {
    trigger.classList.toggle("is-loading", isLoading)
  })
}

function persistWorkbook(workbookId, workbook, source) {
  try {
    const serialized = XLSX.write(workbook, { bookType: "xlsx", type: "base64" })
    localStorage.setItem(STORAGE_KEYS.data(workbookId), serialized)
    localStorage.setItem(STORAGE_KEYS.source(workbookId), source || "arquivo importado")
  } catch (error) {
    console.warn("Não foi possível salvar a planilha localmente.", error)
  }
}

function loadWorkbookFromStorage(workbookId) {
  try {
    const data = localStorage.getItem(STORAGE_KEYS.data(workbookId))
    if (!data) return null
    const source = localStorage.getItem(STORAGE_KEYS.source(workbookId)) || "arquivo importado"
    const workbook = XLSX.read(data, { type: "base64" })
    return { workbook, source }
  } catch (error) {
    console.warn("Não foi possível carregar a planilha salva.", error)
    return null
  }
}

function renderSheetChart(page, tableData) {
  const canvas = document.getElementById("dataChart")
  const empty = document.getElementById("chartEmpty")
  if (!canvas || !tableData || page.view === "table") {
    toggleChartState(null, empty)
    return
  }

  const numericSeries = extractNumericSeries(tableData)
  if (!numericSeries.length) {
    toggleChartState(null, empty)
    return
  }

  const ctx = canvas.getContext("2d")
  if (chartInstance) {
    chartInstance.destroy()
  }

  toggleChartState("chart", empty)

  chartInstance = new Chart(ctx, {
    type: "bar",
    data: {
      labels: numericSeries.map((item) => item.label),
      datasets: [
        {
          label: page.chartLabel || page.sheetLabel || page.sheet,
          data: numericSeries.map((item) => item.value),
          backgroundColor: "rgba(239, 68, 68, 0.15)",
          borderColor: "#ef4444",
          borderWidth: 2,
          borderRadius: 6,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (context) => `${context.dataset.label}: ${context.parsed.y}`,
          },
        },
      },
      scales: {
        y: { beginAtZero: true, ticks: { precision: 0 } },
      },
    },
  })
}

function extractNumericSeries(tableData) {
  if (!tableData) return []

  const { headers, rows } = tableData
  const series = []

  headers.forEach((header, columnIndex) => {
    let total = 0
    let hasNumeric = false

    rows.forEach((row) => {
      const value = row[columnIndex]
      const numeric = parseNumeric(value)
      if (numeric !== null) {
        hasNumeric = true
        total += numeric
      }
    })

    if (hasNumeric) {
      series.push({ label: header, value: Number(total.toFixed(2)) })
    }
  })

  return series
}

function parseNumeric(value) {
  if (typeof value === "number" && !Number.isNaN(value)) return value
  if (typeof value === "string") {
    const sanitized = value.replace(/\./g, "").replace(",", ".")
    const parsed = parseFloat(sanitized)
    return Number.isNaN(parsed) ? null : parsed
  }
  return null
}

function toggleChartState(state, emptyNode) {
  if (!emptyNode) return
  if (state === "chart") {
    emptyNode.classList.remove("is-visible")
  } else {
    emptyNode.classList.add("is-visible")
    if (chartInstance) {
      chartInstance.destroy()
      chartInstance = null
    }
  }
}

function showError(error) {
  const table = document.getElementById("dataTable")
  const footnote = document.getElementById("tableFootnote")
  if (table) table.innerHTML = ""
  if (footnote) footnote.textContent = error.message
  console.error(error)
}

function updatePageLayout(page) {
  const tableSection = document.querySelector("[data-table-section]")
  const chartSection = document.querySelector("[data-chart-section]")
  const view = page.view || "table"

  if (tableSection) {
    tableSection.style.display = view === "chart" ? "none" : ""
  }

  if (chartSection) {
    if (view === "table") {
      chartSection.style.display = "none"
    } else {
      chartSection.style.display = ""
    }
  }
}

function parseWorksheetTable(worksheet) {
  const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false })
  const rows = rawRows
    .map((row) => (Array.isArray(row) ? row : []))
    .filter((row) => row.some((cell) => hasValue(cell)))

  if (!rows.length) return { headers: [], rows: [] }

  const headerStartIndex = findHeaderStartIndex(rows)
  if (headerStartIndex === -1) return { headers: [], rows: [] }

  const candidateRows = rows.slice(headerStartIndex)
  const maxCols = getMaxColumns(candidateRows.slice(0, 80))
  const headerDepth = inferHeaderDepth(candidateRows, maxCols)

  const headerRows = candidateRows.slice(0, headerDepth).map((row) => normalizeRow(row, maxCols))
  const dataRows = candidateRows.slice(headerDepth).map((row) => normalizeRow(row, maxCols))

  const filledHeaderRows = headerRows.map((row) => forwardFillRow(row).map((cell) => normalizeHeaderCell(cell)))

  let lastCol = -1
  for (let col = 0; col < maxCols; col++) {
    const hasHeader = filledHeaderRows.some((row) => row[col])
    const hasData = dataRows.some((row) => hasValue(row[col]))
    if (hasHeader || hasData) lastCol = col
  }

  const finalCols = lastCol + 1
  if (finalCols <= 0) return { headers: [], rows: [] }

  const headers = Array.from({ length: finalCols }, (_, col) => {
    const parts = filledHeaderRows.map((row) => row[col]).filter(Boolean)
    const uniqueParts = []
    parts.forEach((part) => {
      if (!uniqueParts.includes(part)) uniqueParts.push(part)
    })
    return uniqueParts.join(" / ") || `Coluna ${col + 1}`
  })

  const trimmedRows = dataRows
    .map((row) => row.slice(0, finalCols))
    .filter((row) => row.some((cell) => hasValue(cell)))

  return { headers, rows: trimmedRows }
}

function getEarlyNonEmptyCount(row, limit) {
  let count = 0
  for (let i = 0; i < Math.min(limit, row.length); i++) {
    if (hasValue(row[i])) count += 1
  }
  return count
}

function scoreHeaderRow(row) {
  const stats = getRowStats(row)
  if (stats.nonEmpty < 2 || stats.strings < 1) {
    return Number.NEGATIVE_INFINITY
  }

  const earlyNonEmpty = getEarlyNonEmptyCount(row, 6)
  return stats.strings * 3 + stats.nonEmpty + earlyNonEmpty * 2 - stats.numbers * 4
}

function findHeaderStartIndex(rows) {
  const limit = Math.min(rows.length, 80)
  let bestIndex = -1
  let bestScore = Number.NEGATIVE_INFINITY

  for (let index = 0; index < limit; index++) {
    const score = scoreHeaderRow(rows[index])
    if (score > bestScore) {
      bestScore = score
      bestIndex = index
    }
  }

  if (bestIndex === -1 || bestScore === Number.NEGATIVE_INFINITY) return -1

  let startIndex = bestIndex
  for (let index = bestIndex - 1; index >= 0; index--) {
    const stats = getRowStats(rows[index])
    const earlyNonEmpty = getEarlyNonEmptyCount(rows[index], 6)
    const looksLikeHeader =
      stats.nonEmpty >= 2 &&
      stats.strings >= 1 &&
      stats.numbers <= 1 &&
      (earlyNonEmpty >= 2 || (stats.nonEmpty >= 4 && stats.strings >= 2))

    if (!looksLikeHeader) break
    startIndex = index
  }

  return startIndex
}

function inferHeaderDepth(rows, maxCols) {
  if (!rows.length) return 1

  const firstStats = getRowStats(rows[0])
  const fillRate = maxCols ? firstStats.nonEmpty / maxCols : 1
  if (fillRate >= 0.6) return 1

  let depth = 1
  for (let i = 1; i < Math.min(rows.length, 4); i++) {
    const stats = getRowStats(rows[i])
    if (stats.nonEmpty >= 2 && stats.strings >= 1 && stats.numbers <= 1) {
      depth += 1
      continue
    }
    break
  }

  return depth
}

function getMaxColumns(rows) {
  return rows.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0)
}

function normalizeRow(row, length) {
  const normalized = Array.isArray(row) ? row.slice(0) : []
  while (normalized.length < length) normalized.push("")
  return normalized
}

function forwardFillRow(row) {
  const filled = row.slice(0)
  let last = ""

  for (let i = 0; i < filled.length; i++) {
    const current = normalizeHeaderCell(filled[i])
    if (current) {
      last = current
      filled[i] = current
    } else if (last) {
      filled[i] = last
    } else {
      filled[i] = ""
    }
  }

  return filled
}

function normalizeHeaderCell(value) {
  if (!hasValue(value)) return ""
  return String(value).replace(/\s+/g, " ").trim()
}

function hasValue(value) {
  return value !== undefined && value !== null && String(value).trim() !== ""
}

function getRowStats(row) {
  let nonEmpty = 0
  let numbers = 0
  let strings = 0

  row.forEach((value) => {
    if (!hasValue(value)) return
    nonEmpty += 1

    if (typeof value === "number" && !Number.isNaN(value)) {
      numbers += 1
      return
    }

    if (typeof value === "string") {
      const parsed = parseNumeric(value)
      if (parsed !== null) {
        numbers += 1
      } else {
        strings += 1
      }
      return
    }

    strings += 1
  })

  return { nonEmpty, numbers, strings }
}
