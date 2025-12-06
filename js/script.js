const workbookPath = "dados/TESTE PARA INDICADORES - CABEÇALHOS.xlsx"
const STORAGE_KEYS = {
  data: "sgiWorkbookData",
  source: "sgiWorkbookSource",
}

const pages = [
  {
    id: "acidentes-c-afastamento",
    sheet: "ACIDENTES C AFASTAMENTO",
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
]

const navGroups = [
  {
    id: "acidentes",
    label: "Indicadores de Segurança",
    pages: ["acidentes-c-afastamento", "acidentes-s-afastamento"],
  },
  {
    id: "avaliacoes",
    label: "Auditorias e Garantia",
    pages: ["aud-comp", "garantia-vida"],
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

let workbookCache = null
let workbookSource = "arquivo padrão"
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
  loadWorkbook()
    .then((workbook) => {
      const tableData = renderSheetTable(pageInfo, workbook, { source: workbookSource })
      renderSheetChart(pageInfo, tableData)
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
  document.querySelectorAll("[data-sheet-description]").forEach((node) => (node.textContent = page.description))
  document.querySelectorAll("[data-table-title]").forEach((node) => (node.textContent = page.sheet))
  document
    .querySelectorAll("[data-sheet-note]")
    .forEach((node) => (node.textContent = `Dados da aba "${page.sheet}" - nomes idênticos à planilha.`))
}

async function loadWorkbook() {
  if (workbookCache) return workbookCache

  const stored = loadWorkbookFromStorage()
  if (stored) {
    workbookCache = stored.workbook
    workbookSource = stored.source
    return workbookCache
  }

  const response = await fetch(workbookPath)
  if (!response.ok) {
    throw new Error("Não foi possível carregar o arquivo de indicadores.")
  }
  const buffer = await response.arrayBuffer()
  workbookCache = XLSX.read(buffer, { type: "array" })
  workbookSource = workbookPath.split("/").pop() || "arquivo padrão"
  return workbookCache
}

function renderSheetTable(page, workbook, options = {}) {
  const table = document.getElementById("dataTable")
  const footnote = document.getElementById("tableFootnote")
  if (!table) return

  const worksheet = workbook.Sheets[page.sheet]
  if (!worksheet) {
    table.innerHTML = ""
    if (footnote) footnote.textContent = `A aba "${page.sheet}" não foi encontrada na planilha.`
    return
  }

  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false })
  const headers = (rows.shift() || []).filter((header) => header !== undefined && header !== null && header !== "")
  const dataRows = rows.filter((row) => row.some((cell) => cell !== undefined && cell !== null && cell !== ""))

  if (!headers.length) {
    table.innerHTML = ""
    if (footnote) footnote.textContent = `A aba "${page.sheet}" não possui cabeçalho definido.`
    return null
  }

  const thead = document.createElement("thead")
  const headRow = document.createElement("tr")
  headers.forEach((header) => {
    const th = document.createElement("th")
    th.textContent = header
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
    const sourceLabel = options.source || workbookSource || "arquivo padrão"
    footnote.textContent = `Colunas exibidas exatamente como na planilha: ${headers.join(" | ")} (fonte: ${sourceLabel})`
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
      workbookCache = XLSX.read(buffer, { type: "array" })
      workbookSource = file.name
      persistWorkbook(workbookCache, workbookSource)
      const tableData = renderSheetTable(pageInfo, workbookCache, { source: workbookSource })
      renderSheetChart(pageInfo, tableData)
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

function persistWorkbook(workbook, source) {
  try {
    const serialized = XLSX.write(workbook, { bookType: "xlsx", type: "base64" })
    localStorage.setItem(STORAGE_KEYS.data, serialized)
    localStorage.setItem(STORAGE_KEYS.source, source || "arquivo importado")
  } catch (error) {
    console.warn("Não foi possível salvar a planilha localmente.", error)
  }
}

function loadWorkbookFromStorage() {
  try {
    const data = localStorage.getItem(STORAGE_KEYS.data)
    if (!data) return null
    const source = localStorage.getItem(STORAGE_KEYS.source) || "arquivo importado"
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
          label: page.sheet,
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
