const state = {
  rawRows: [],
  headers: [],
  types: {},
  charts: {
    category: null,
    timeline: null,
  },
};

const elements = {
  status: document.getElementById("status"),
  categorySelect: document.getElementById("categorySelect"),
  valueSelect: document.getElementById("valueSelect"),
  categoryFilter: document.getElementById("categoryFilter"),
  startDate: document.getElementById("startDate"),
  endDate: document.getElementById("endDate"),
  summaryTotal: document.getElementById("summaryTotal"),
  summaryCount: document.getElementById("summaryCount"),
  summaryAverage: document.getElementById("summaryAverage"),
  table: document.getElementById("dataTable"),
  tableCount: document.getElementById("tableCount"),
  downloadCsv: document.getElementById("downloadCsv"),
};

const DEFAULT_FILE = "TESTE222222 - NOTA FISCAL.xlsx";
const GENERIC_HEADER_PATTERN = /^(coluna|column|campo)\s*\d+$/i;
const TIMELINE_DEFAULT_TITLE = "Séries por período";
const FALLBACK_CATEGORY_LIMIT = 12;
const DEFAULT_METRIC_KEY = normalizeKey("Valor Total NF");

function normalizeKey(text) {
  if (!text) return "";
  return String(text)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .trim()
    .toLowerCase();
}

const EMISSION_CANDIDATE_KEYS = [
  "Data Emissão",
  "Data emissão",
  "Data de emissão",
  "Emissão",
].map((header) => normalizeKey(header));

function getEmissionField() {
  return state.headers.find((header) => EMISSION_CANDIDATE_KEYS.includes(normalizeKey(header))) || null;
}

function getTimelineDateField() {
  const emissionField = getEmissionField();
  if (emissionField) return emissionField;
  return state.headers.find((header) => state.types[header] === "date") || null;
}

function selectDefaultMetricOption() {
  if (!elements.valueSelect) return;
  const options = Array.from(elements.valueSelect.options);
  const index = options.findIndex((option) => normalizeKey(option.value) === DEFAULT_METRIC_KEY);
  if (index >= 0) {
    elements.valueSelect.selectedIndex = index;
  } else if (options.length > 0) {
    elements.valueSelect.selectedIndex = 0;
  }
}

function getDefaultMetricField() {
  const header = state.headers.find((h) => normalizeKey(h) === DEFAULT_METRIC_KEY);
  return header || null;
}

document.addEventListener("DOMContentLoaded", () => {
  setupListeners();
  preloadDefaultFile();
});

function showStatus(message, isError = false) {
  if (!elements.status) return;
  elements.status.textContent = message;
  elements.status.style.color = isError ? "#fca5a5" : "var(--muted)";
}

function setupListeners() {
  const fileInput = document.getElementById("fileInput");
  fileInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      showStatus(`Carregando ${file.name}...`);
      const arrayBuffer = await file.arrayBuffer();
      await processWorkbook(arrayBuffer, file.name);
      showStatus(`Arquivo ${file.name} carregado com sucesso.`);
    } catch (error) {
      console.error(error);
      showStatus(
        "Não foi possível ler o arquivo selecionado. Verifique o formato e tente novamente.",
        true,
      );
    }
  });

  [
    elements.categorySelect,
    elements.valueSelect,
    elements.categoryFilter,
    elements.startDate,
    elements.endDate,
  ].forEach((input) => {
    input?.addEventListener("input", handleControlsChange);
  });

  elements.downloadCsv.addEventListener("click", downloadFilteredCsv);
}

async function preloadDefaultFile() {
  try {
    showStatus(`Carregando ${DEFAULT_FILE} padrão...`);
    const response = await fetch(DEFAULT_FILE);
    if (!response.ok) throw new Error("Falha no download do arquivo padrão");
    const arrayBuffer = await response.arrayBuffer();
    await processWorkbook(arrayBuffer, DEFAULT_FILE);
    showStatus(`Arquivo padrão ${DEFAULT_FILE} carregado.`);
  } catch (error) {
    console.warn("Não foi possível carregar o arquivo padrão.", error);
    showStatus(
      "Não foi possível carregar o arquivo padrão. Utilize o botão \"Importar outro arquivo\".",
      true,
    );
  }
}

async function processWorkbook(arrayBuffer, fileName) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) throw new Error("Planilha principal não encontrada");

  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null });
  if (!data.length) throw new Error("Planilha vazia");

  const headerRowIndex = findHeaderRow(data);
  const headerRow = data[headerRowIndex] ?? [];
  const headers = headerRow.map((header, index) => sanitizeHeaderName(header, index));
  const rowsMatrix = data
    .slice(headerRowIndex + 1)
    .filter((row) => row.some((cell) => cell !== null && cell !== ""));

  if (!rowsMatrix.length) {
    throw new Error("Planilha sem dados após a linha de cabeçalho detectada");
  }

  const types = detectColumnTypes(headers, rowsMatrix);
  applyTypeOverrides(headers, types);
  const normalizedRows = normalizeRows(headers, rowsMatrix, types);

  state.rawRows = normalizedRows;
  state.headers = headers;
  state.types = types;

  populateControls(headers, types);
  renderTableHeaders(headers);
  resetControlValues();
  updateAllVisuals();

  console.info(`Planilha ${fileName} carregada com ${rowsMatrix.length} registros.`);
}

function findHeaderRow(data) {
  let bestIndex = 0;
  let bestScore = Number.NEGATIVE_INFINITY;
  const inspectLimit = Math.min(data.length, 30);

  for (let i = 0; i < inspectLimit; i += 1) {
    const row = data[i];
    if (!row) continue;

    const normalized = row.map((cell) =>
      cell === null || cell === undefined ? "" : String(cell).trim(),
    );
    const filled = normalized.filter(Boolean);
    if (filled.length < 2) continue;

    let stringCount = 0;
    let numericLikeCount = 0;
    let genericCount = 0;

    filled.forEach((value) => {
      if (GENERIC_HEADER_PATTERN.test(value)) genericCount += 1;
      if (isLikelyNumeric(value)) {
        numericLikeCount += 1;
      } else {
        stringCount += 1;
      }
    });

    const uniquenessPenalty = filled.length - new Set(filled.map((v) => v.toLowerCase())).size;
    const score =
      stringCount * 2 -
      numericLikeCount * 1.5 -
      genericCount * 3 -
      uniquenessPenalty +
      filled.length;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestIndex;
}

function sanitizeHeaderName(value, index) {
  if (value === null || value === undefined) return `Coluna ${index + 1}`;
  const text = String(value).trim();
  if (!text || GENERIC_HEADER_PATTERN.test(text)) {
    return `Coluna ${index + 1}`;
  }
  return text;
}

function isLikelyNumeric(value) {
  if (!value) return false;
  const normalized = value.replace(/\./g, "").replace(/,/g, ".");
  if (normalized === "") return false;
  return Number.isFinite(Number(normalized));
}

const TYPE_OVERRIDES = new Map(
  [
    { header: "N° NF", type: "string" },
    { header: "Nº NF", type: "string" },
    { header: "Numero NF", type: "string" },
    { header: "Número NF", type: "string" },
    { header: "Série", type: "string" },
    { header: "Serie", type: "string" },
    { header: "CNPJ Emitente", type: "string" },
    { header: "CNPJ EMITENTE", type: "string" },
    { header: "CNPJ Destinatário", type: "string" },
    { header: "CNPJ DESTINATÁRIO", type: "string" },
    { header: "CNPJ Destinatario", type: "string" },
    { header: "CNPJ DESTINATARIO", type: "string" },
    { header: "Chave de Acesso", type: "string" },
    { header: "Chave de acesso", type: "string" },
    { header: "Valor do Produto", type: "number" },
    { header: "Valor Total NF", type: "number" },
    { header: "Valor Produtos", type: "number" },
  ].map(({ header, type }) => [normalizeKey(header), type]),
);

const CURRENCY_COLUMNS = new Set(
  ["Valor do Produto", "Valor Total NF", "Valor Produtos"].map((header) => normalizeKey(header)),
);

function detectColumnTypes(headers, rows) {
  const types = {};
  headers.forEach((header, columnIndex) => {
    const columnValues = rows
      .map((row) => row[columnIndex])
      .filter((value) => value !== null && value !== "");

    let numericCount = 0;
    let dateCount = 0;

    columnValues.forEach((value) => {
      if (isFinite(value)) numericCount += 1;

      if (typeof value === "number") {
        const parsed = XLSX.SSF.parse_date_code(value);
        if (parsed) dateCount += 1;
      } else if (typeof value === "string") {
        const date = new Date(value);
        if (!Number.isNaN(date.getTime())) dateCount += 1;
      }
    });

    const type =
      dateCount >= numericCount && dateCount > 0
        ? "date"
        : numericCount > 0
        ? "number"
        : "string";

    types[header] = type;
  });
  return types;
}

function normalizeRows(headers, rows, types) {
  return rows.map((row) => {
    const entry = {};
    headers.forEach((header, index) => {
      let value = row[index];
      if (value === undefined || value === null || value === "") {
        entry[header] = null;
        return;
      }

      const type = types[header];
      if (type === "number") {
        const numeric = parseNumber(value);
        entry[header] = Number.isFinite(numeric) ? numeric : null;
      } else if (type === "date") {
        let date;
        if (typeof value === "number") {
          const parsed = XLSX.SSF.parse_date_code(value);
          if (parsed) {
            date = new Date(
              parsed.y,
              parsed.m - 1,
              parsed.d,
              parsed.H,
              parsed.M,
              parsed.S,
            );
          }
        }
        if (!date) {
          const tryDate = new Date(value);
          date = Number.isNaN(tryDate.getTime()) ? null : tryDate;
        }
        entry[header] = date;
      } else {
        entry[header] = String(value).trim();
      }
    });
    return entry;
  });
}

function applyTypeOverrides(headers, types) {
  headers.forEach((header) => {
    const override = TYPE_OVERRIDES.get(normalizeKey(header));
    if (override) {
      types[header] = override;
    }
  });
}

function parseNumber(value) {
  if (typeof value === "number") return value;
  if (typeof value !== "string") return NaN;

  let normalized = value.trim();
  if (!normalized) return NaN;

  normalized = normalized.replace(/\u0000/g, "");
  normalized = normalized.replace(/[R$\s]/gi, "");
  normalized = normalized.replace(/[^0-9,.-]+/g, "");

  const hasComma = normalized.includes(",");
  const hasDot = normalized.includes(".");
  if (hasComma && hasDot) {
    if (normalized.lastIndexOf(",") > normalized.lastIndexOf(".")) {
      normalized = normalized.replace(/\./g, "").replace(/,/g, ".");
    } else {
      normalized = normalized.replace(/,/g, "");
    }
  } else if (hasComma && !hasDot) {
    normalized = normalized.replace(/\./g, "").replace(/,/g, ".");
  } else {
    normalized = normalized.replace(/,/g, "");
  }

  const numeric = Number(normalized);
  return Number.isFinite(numeric) ? numeric : NaN;
}

function generateTimelineColors(count) {
  if (count <= 0) return [];

  const startAlpha = 0.85;
  const endAlpha = 0.35;
  const colors = [];

  for (let index = 0; index < count; index += 1) {
    const ratio = count === 1 ? 0 : index / (count - 1);
    const alpha = (startAlpha - (startAlpha - endAlpha) * ratio).toFixed(2);
    colors.push(`rgba(56, 189, 248, ${alpha})`);
  }

  return colors;
}

function formatTimelineLabel(latestDate, key) {
  if (latestDate instanceof Date && !Number.isNaN(latestDate.getTime())) {
    return latestDate.toLocaleDateString("pt-BR", { month: "short", year: "numeric" });
  }

  if (typeof key === "string") {
    const [year, month] = key.split("-");
    if (year && month) {
      const parsed = new Date(Number(year), Number(month) - 1, 1);
      if (!Number.isNaN(parsed.getTime())) {
        return parsed.toLocaleDateString("pt-BR", { month: "short", year: "numeric" });
      }
    }
  }

  return key;
}

const METRIC_EXCLUSIONS = new Set(["Chave de acesso"].map((header) => normalizeKey(header)));

function populateControls(headers, types) {
  const preferredCategories = [
    "Nome Emitente",
    "CNPJ Emitente",
    "Nome Destinatario",
  ];

  const preferredCategoryOptions = preferredCategories
    .map((name) => headers.find((header) => normalizeKey(header) === normalizeKey(name)))
    .filter(Boolean);

  const categoryOptions = preferredCategoryOptions.length
    ? preferredCategoryOptions
    : headers.filter((h) => types[h] === "string");

  fillSelect(elements.categorySelect, categoryOptions);
  fillSelect(
    elements.valueSelect,
    headers.filter((h) => types[h] === "number" && !METRIC_EXCLUSIONS.has(normalizeKey(h))),
  );

  selectDefaultMetricOption();
}

function fillSelect(select, options, allowEmpty = false) {
  if (!select) return;
  select.innerHTML = "";

  if (allowEmpty || !options.length) {
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = allowEmpty
      ? "(Sem data)"
      : "Selecione uma coluna";
    select.appendChild(placeholder);
  }

  options.forEach((option) => {
    const opt = document.createElement("option");
    opt.value = option;
    opt.textContent = option;
    select.appendChild(opt);
  });
}

function resetControlValues() {
  if (elements.categorySelect.options.length > 0) {
    elements.categorySelect.selectedIndex = 0;
  }
  selectDefaultMetricOption();
  elements.categoryFilter.value = "";
  elements.startDate.value = "";
  elements.endDate.value = "";
}

function handleControlsChange() {
  updateAllVisuals();
}

function updateAllVisuals() {
  const filters = collectFilters();
  const filteredRows = applyFilters(state.rawRows, filters);

  updateSummary(filteredRows, filters.valueField);
  updateCategoryChart(filteredRows, filters);
  updateTimelineChart(filteredRows, filters);
  renderTableBody(filteredRows);
}

function collectFilters() {
  const categoryField = elements.categorySelect.value || null;
  const valueField = elements.valueSelect.value || getDefaultMetricField();
  const categoryQuery = elements.categoryFilter.value.trim().toLowerCase();

  const startDate = elements.startDate.value ? new Date(elements.startDate.value) : null;
  const endDate = elements.endDate.value ? new Date(elements.endDate.value) : null;
  const dateField = getTimelineDateField();

  return { categoryField, valueField, dateField, categoryQuery, startDate, endDate };
}

function applyFilters(rows, filters) {
  return rows.filter((row) => {
    if (filters.categoryField && filters.categoryQuery) {
      const value = row[filters.categoryField];
      if (!value || !value.toLowerCase().includes(filters.categoryQuery)) {
        return false;
      }
    }

    const emissionField = getEmissionField();
    if ((filters.startDate || filters.endDate) && emissionField) {
      const dateValue = row[emissionField];
      if (!(dateValue instanceof Date)) return false;

      if (filters.startDate && dateValue < filters.startDate) return false;
      if (filters.endDate) {
        const end = new Date(filters.endDate);
        end.setHours(23, 59, 59, 999);
        if (dateValue > end) return false;
      }
    }

    return true;
  });
}

function updateSummary(rows, valueField) {
  elements.summaryCount.textContent = rows.length.toLocaleString("pt-BR");

  if (valueField) {
    const sum = rows.reduce((acc, row) => acc + (row[valueField] ?? 0), 0);
    const average = rows.length ? sum / rows.length : 0;
    const useCurrency = CURRENCY_COLUMNS.has(normalizeKey(valueField));

    elements.summaryTotal.textContent = formatNumber(sum, useCurrency);
    elements.summaryAverage.textContent = formatNumber(average, useCurrency);
  } else {
    elements.summaryTotal.textContent = "–";
    elements.summaryAverage.textContent = "–";
  }
}

function updateCategoryChart(rows, filters) {
  const canvas = document.getElementById("categoryChart");
  if (!canvas) return;

  const { categoryField, valueField } = filters;
  if (!categoryField) {
    canvas.style.opacity = 0.3;
    renderChart(state.charts.category, "categoryChart", {
      type: "bar",
      data: { labels: [], datasets: [] },
      options: { responsive: true },
    });
    return;
  }
  canvas.style.opacity = 1;

  const aggregation = new Map();
  rows.forEach((row) => {
    const key = row[categoryField] || "(Sem valor)";
    const current = aggregation.get(key) ?? { total: 0, count: 0 };
    current.total += valueField ? row[valueField] ?? 0 : 1;
    current.count += 1;
    aggregation.set(key, current);
  });

  const labels = Array.from(aggregation.keys());
  const data = labels.map((label) => {
    const { total, count } = aggregation.get(label);
    return valueField ? total : count;
  });

  const datasetLabel = valueField ? `Soma de ${valueField}` : `Quantidade por ${categoryField}`;
  const useCurrency = valueField ? CURRENCY_COLUMNS.has(normalizeKey(valueField)) : false;

  state.charts.category = renderChart(state.charts.category, "categoryChart", {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: datasetLabel,
          data,
          backgroundColor: "rgba(14, 165, 233, 0.6)",
          borderRadius: 8,
        },
      ],
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label(context) {
              const value = context.parsed.y;
              return valueField
                ? formatNumber(value, useCurrency)
                : `${value} registros`;
            },
          },
        },
      },
      scales: {
        x: {
          ticks: { color: "#cbd5f5" },
        },
        y: {
          ticks: {
            color: "#cbd5f5",
            callback: (value) => (valueField ? formatNumber(value, useCurrency) : value),
          },
        },
      },
    },
  });
}

function updateTimelineChart(rows, filters) {
  const canvas = document.getElementById("timelineChart");
  if (!canvas) return;

  const headerEl = canvas.closest("article")?.querySelector("h2");
  const { dateField, valueField, categoryField } = filters;
  const useCurrency = valueField ? CURRENCY_COLUMNS.has(normalizeKey(valueField)) : false;

  if (!dateField) {
    const fallbackField =
      categoryField || state.headers.find((header) => state.types[header] === "string");

    if (!fallbackField) {
      canvas.style.opacity = 0.3;
      if (headerEl) headerEl.textContent = TIMELINE_DEFAULT_TITLE;
      renderChart(state.charts.timeline, "timelineChart", {
        type: "line",
        data: { labels: [], datasets: [] },
        options: { responsive: true },
      });
      return;
    }

    const aggregation = new Map();
    rows.forEach((row) => {
      const key = row[fallbackField] || "(Sem valor)";
      const current = aggregation.get(key) ?? { total: 0, count: 0 };
      current.total += valueField ? row[valueField] ?? 0 : 1;
      current.count += 1;
      aggregation.set(key, current);
    });

    const sortedEntries = Array.from(aggregation.entries()).sort((a, b) => {
      const aValue = valueField ? a[1].total : a[1].count;
      const bValue = valueField ? b[1].total : b[1].count;
      return bValue - aValue;
    });

    const limitedEntries = sortedEntries.slice(0, FALLBACK_CATEGORY_LIMIT);
    const labels = limitedEntries.map(([label]) => label);
    const data = limitedEntries.map(([, { total, count }]) => (valueField ? total : count));
    const datasetLabel = valueField
      ? `Soma de ${valueField}`
      : `Quantidade por ${fallbackField}`;

    canvas.style.opacity = 1;
    if (headerEl) headerEl.textContent = `Distribuição por ${fallbackField}`;
    state.charts.timeline = renderChart(state.charts.timeline, "timelineChart", {
      type: "bar",
      data: {
        labels,
        datasets: [
          {
            label: datasetLabel,
            data,
            backgroundColor: "rgba(99, 102, 241, 0.6)",
            borderRadius: 8,
          },
        ],
      },
      options: {
        responsive: true,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label(context) {
                const value = context.parsed.y;
                return valueField
                  ? formatNumber(value, useCurrency)
                  : `${value} registros`;
              },
            },
          },
        },
        scales: {
          x: {
            ticks: { color: "#cbd5f5" },
          },
          y: {
            ticks: {
              color: "#cbd5f5",
              callback: (value) => (valueField ? formatNumber(value, useCurrency) : value),
            },
          },
        },
      },
    });
    return;
  }

  if (headerEl) headerEl.textContent = TIMELINE_DEFAULT_TITLE;
  canvas.style.opacity = 1;

  const aggregation = new Map();
  rows.forEach((row) => {
    const date = row[dateField];
    if (!(date instanceof Date)) return;
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    const current = aggregation.get(key) ?? { total: 0, count: 0, date };
    current.total += valueField ? row[valueField] ?? 0 : 1;
    current.count += 1;
    if (!current.date || current.date < date) current.date = date;
    aggregation.set(key, current);
  });

  if (!aggregation.size) {
    canvas.style.opacity = 0.3;
    state.charts.timeline = renderChart(state.charts.timeline, "timelineChart", {
      type: "bar",
      data: { labels: [], datasets: [] },
      options: { responsive: true },
    });
    return;
  }

  const sortedEntries = Array.from(aggregation.entries())
    .map(([key, stats]) => ({
      key,
      label: formatTimelineLabel(stats.date, key),
      value: valueField ? stats.total : stats.count,
    }))
    .sort((a, b) => b.value - a.value);

  const limitedEntries = sortedEntries.slice(0, FALLBACK_CATEGORY_LIMIT);
  const labels = limitedEntries.map((entry) => entry.label);
  const data = limitedEntries.map((entry) => entry.value);
  const palette = generateTimelineColors(data.length);

  const datasetLabel = valueField
    ? `Maiores valores por período (${valueField})`
    : "Maiores quantidades por período";

  state.charts.timeline = renderChart(state.charts.timeline, "timelineChart", {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: datasetLabel,
          data,
          backgroundColor: palette,
          borderColor: palette.map(() => "rgba(14, 165, 233, 1)"),
          borderWidth: 1.5,
          borderRadius: 10,
        },
      ],
    },
    options: {
      responsive: true,
      indexAxis: "y",
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label(context) {
              const parsedValue =
                typeof context.parsed.x === "number" ? context.parsed.x : context.parsed.y;
              return valueField
                ? formatNumber(parsedValue, useCurrency)
                : `${parsedValue} registros`;
            },
          },
        },
      },
      scales: {
        x: {
          ticks: {
            color: "#cbd5f5",
            callback: (value) =>
              valueField
                ? formatNumber(Number(value), useCurrency)
                : Number(value).toLocaleString("pt-BR"),
          },
          grid: { color: "rgba(148, 163, 184, 0.2)" },
        },
        y: {
          ticks: { color: "#cbd5f5" },
          grid: { color: "rgba(148, 163, 184, 0.1)" },
        },
      },
    },
  });
}

function renderChart(existingChart, canvasId, config) {
  if (existingChart) {
    existingChart.config.type = config.type;
    existingChart.data = config.data;
    existingChart.options = config.options;
    existingChart.update();
    return existingChart;
  }
  const canvas = document.getElementById(canvasId);
  if (!canvas) return existingChart;
  const ctx = canvas.getContext("2d");
  if (!ctx) return existingChart;
  return new Chart(ctx, config);
}

function renderTableHeaders(headers) {
  if (!elements.table) return;
  const thead = elements.table.querySelector("thead");
  thead.innerHTML = "";

  const row = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    row.appendChild(th);
  });
  thead.appendChild(row);
}

function renderTableBody(rows) {
  const tbody = elements.table.querySelector("tbody");
  tbody.innerHTML = "";

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    state.headers.forEach((header) => {
      const td = document.createElement("td");
      const value = row[header];
      if (value instanceof Date) {
        td.textContent = value.toLocaleDateString("pt-BR");
      } else if (typeof value === "number") {
        const useCurrency = CURRENCY_COLUMNS.has(normalizeKey(header));
        td.textContent = formatNumber(value, useCurrency);
      } else {
        td.textContent = value ?? "";
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  elements.tableCount.textContent = `${rows.length.toLocaleString("pt-BR")} registros`;
}

function formatNumber(value, currency) {
  if (!Number.isFinite(value)) return "–";
  return value.toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
    ...(currency ? { style: "currency", currency: "BRL" } : {}),
  });
}

function downloadFilteredCsv() {
  const filters = collectFilters();
  const filteredRows = applyFilters(state.rawRows, filters);

  if (!filteredRows.length) {
    showStatus("Nenhum registro para exportar.");
    return;
  }

  const csvRows = [state.headers.join(";")];
  filteredRows.forEach((row) => {
    const values = state.headers.map((header) => {
      const value = row[header];
      if (value instanceof Date) {
        return value.toISOString();
      }
      if (value === null || value === undefined) return "";
      return String(value).replace(/"/g, '""');
    });
    csvRows.push(values.map((val) => `"${val}"`).join(";"));
  });

  const blob = new Blob([csvRows.join("\n")], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "dados-filtrados.csv";
  link.click();
  URL.revokeObjectURL(url);

  showStatus("Exportação CSV concluída.");
}
