const DATASETS = [
  "data/parking_analytics_report_2025-10-19_to_2026-01-15.xlsx",
  "data/parking_analytics_report_2026-01-11_to_2026-01-18.xlsx",
  "data/parking_analytics_report_2026-01-11_to_2026-01-18-2.xlsx"
];

const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
const MONDAY_FIRST = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

let loadedFiles = {};
let charts = {};

const datasetSelect = document.getElementById("datasetSelect");
const fileInput = document.getElementById("fileInput");
const statusMessage = document.getElementById("statusMessage");

const valueLabelPlugin = {
  id: "valueLabelPlugin",
  afterDatasetsDraw(chart) {
    if (chart.config.type !== "bar") return;

    const { ctx } = chart;
    ctx.save();
    ctx.fillStyle = "#24313d";
    ctx.font = "12px Georgia, serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";

    chart.data.datasets.forEach((dataset, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      meta.data.forEach((bar, index) => {
        const value = dataset.data[index];
        if (!Number.isFinite(value)) return;

        const isPercentChart = chart.canvas.id === "weekdayChart";
        const label = isPercentChart ? `${value.toFixed(1)}%` : String(Math.round(value));
        ctx.fillText(label, bar.x, bar.y - 4);
      });
    });

    ctx.restore();
  }
};

Chart.register(valueLabelPlugin);

datasetSelect.addEventListener("change", () => loadDataset(datasetSelect.value));
fileInput.addEventListener("change", handleManualFiles);

loadDataset(DATASETS[0]);

async function loadDataset(fileName) {
  const cacheKey = getFileName(fileName);
  statusMessage.textContent = `Loading ${cacheKey}...`;

  try {
    const workbook = loadedFiles[cacheKey] || await fetchWorkbook(fileName);
    loadedFiles[cacheKey] = workbook;

    const parsedData = parseWorkbook(workbook);
    renderDashboard(parsedData);

    statusMessage.textContent = `Showing ${cacheKey}`;
  } catch (error) {
    clearDashboard();
    statusMessage.textContent =
      "Could not automatically load the Excel file. Use the file chooser above, then select a dataset.";
    console.warn(error);
  }
}

async function fetchWorkbook(fileName) {
  const response = await fetch(fileName);
  if (!response.ok) throw new Error(`Unable to load ${fileName}`);

  const arrayBuffer = await response.arrayBuffer();
  return XLSX.read(arrayBuffer, { type: "array", cellDates: true });
}

async function handleManualFiles(event) {
  const files = Array.from(event.target.files);
  if (!files.length) return;

  for (const file of files) {
    const arrayBuffer = await file.arrayBuffer();
    loadedFiles[file.name] = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  }

  if (loadedFiles[getFileName(datasetSelect.value)]) {
    loadDataset(datasetSelect.value);
    return;
  }

  const firstLoadedName = Object.keys(loadedFiles)[0];
  const matchingDataset = DATASETS.find(dataset => getFileName(dataset) === firstLoadedName);
  if (matchingDataset) datasetSelect.value = matchingDataset;
  loadDataset(matchingDataset || firstLoadedName);
}

function getFileName(path) {
  return String(path || "").split("/").pop();
}

function parseWorkbook(workbook) {
  const sheetNames = workbook.SheetNames || [];
  const summarySheet = findSheet(workbook, ["summary"]);
  const rawSheet = findSheet(workbook, ["raw", "data"]);
  const dailySheets = sheetNames.filter(name => normalize(name).includes("daily"));

  const rawRows = sheetToObjects(rawSheet);
  const parsedRawRows = parseRawRows(rawRows);
  const rawData = parsedRawRows.length ? parsedRawRows : parseRawSheetByPosition(rawSheet);

  const dailyFromSummary = parseSummaryDailyOccupancy(summarySheet);
  const hourlyFromRaw = calculateHourlyOccupancyFromRaw(rawData);
  const hourlyFromDailySheets = parseDailySheets(workbook, dailySheets);
  const hourlyOccupancy = hourlyFromRaw.length ? hourlyFromRaw : hourlyFromDailySheets;

  const dailyFromHourly = averageDailyFromHourly(hourlyOccupancy);
  const finalDaily = dailyFromSummary.length ? dailyFromSummary : dailyFromHourly;

  return {
    dailyOccupancy: finalDaily,
    weekdayOccupancy: averageByWeekday(finalDaily),
    hourlyOccupancy,
    sensorActivity: countSensorActivity(rawData),
    insights: buildInsights(finalDaily, hourlyOccupancy, rawData)
  };
}

function findSheet(workbook, keywords) {
  const foundName = workbook.SheetNames.find(name => {
    const cleanedName = normalize(name);
    return keywords.every(keyword => cleanedName.includes(keyword));
  });

  return foundName ? workbook.Sheets[foundName] : null;
}

function normalize(value) {
  return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function sheetToObjects(sheet) {
  if (!sheet) return [];

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headerIndex = rows.findIndex(row => isHeaderRow(row));
  if (headerIndex === -1) return [];

  const headers = rows[headerIndex].map(cell => String(cell || "").trim());
  return rows.slice(headerIndex + 1)
    .filter(row => row.some(cell => cell !== ""))
    .map(row => {
      const object = {};
      headers.forEach((header, index) => {
        if (header) object[header] = row[index];
      });
      return object;
    });
}

function looksLikeHeader(cell) {
  const text = normalize(cell);
  return ["date", "hour", "recorded", "friendly", "occupancy", "status", "sensor"].some(word => {
    return text.includes(word);
  });
}

function isHeaderRow(row) {
  const filledCells = row.filter(cell => String(cell || "").trim() !== "");
  const headerLikeCells = filledCells.filter(cell => looksLikeHeader(cell));

  // Ignore decorative title rows such as "RAW SENSOR DATA LOG".
  return filledCells.length >= 2 && headerLikeCells.length >= 2;
}

function parseSummaryDailyOccupancy(sheet) {
  const rows = sheetToObjects(sheet);

  return rows.map(row => {
    const date = parseDateOnly(getColumn(row, ["date"]));
    const occupancy = parsePercent(getColumn(row, ["occupancy"]));

    if (!date || !Number.isFinite(occupancy)) return null;
    return { date, occupancy };
  }).filter(Boolean).sort((a, b) => a.date - b.date);
}

function parseRawRows(rows) {
  return rows.map(row => {
    const timestamp = parseTimestamp(getColumn(row, ["recorded", "timestamp", "time", "date"]));
    const sensor = cleanSensorName(getColumn(row, ["friendly", "sensor name", "sensor", "name"]));
    const status = parseStatus(getColumn(row, ["status", "occupancy"]));

    if (!timestamp || !sensor || status === null) return null;

    return {
      timestamp,
      date: dateAtStartOfDay(timestamp),
      weekday: WEEKDAYS[timestamp.getDay()],
      hour: timestamp.getHours(),
      sensor,
      status,
      occupied: status === "Occupied" ? 1 : 0
    };
  }).filter(Boolean);
}

function parseRawSheetByPosition(sheet) {
  if (!sheet) return [];

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headerIndex = rows.findIndex(row => {
    const text = row.map(normalize).join(" ");
    return text.includes("recorded") && text.includes("status");
  });

  if (headerIndex === -1) return [];

  return rows.slice(headerIndex + 1).map(row => {
    const timestamp = parseTimestamp(row[0]);
    const sensor = cleanSensorName(row[1]);
    const status = parseStatus(row[3]);

    if (!timestamp || !sensor || status === null) return null;

    return {
      timestamp,
      date: dateAtStartOfDay(timestamp),
      weekday: WEEKDAYS[timestamp.getDay()],
      hour: timestamp.getHours(),
      sensor,
      status,
      occupied: status === "Occupied" ? 1 : 0
    };
  }).filter(Boolean);
}

function parseDailySheets(workbook, dailySheetNames) {
  const hourlyItems = [];

  dailySheetNames.forEach(sheetName => {
    const date = parseDateOnly(sheetName);
    const rows = sheetToObjects(workbook.Sheets[sheetName]);
    if (!date || !rows.length) return;

    rows.forEach(row => {
      const hour = parseHour(getColumn(row, ["hour", "time"]));
      const occupancy = parsePercent(getColumn(row, ["occupancy", "rate"]));

      if (Number.isInteger(hour) && Number.isFinite(occupancy)) {
        hourlyItems.push({ date, weekday: WEEKDAYS[date.getDay()], hour, occupancy });
      }
    });
  });

  return hourlyItems;
}

function calculateHourlyOccupancyFromRaw(rawData) {
  const grouped = groupBy(rawData, row => `${row.weekday}-${row.hour}`);

  return Object.values(grouped).map(rows => ({
    date: rows[0].date,
    weekday: rows[0].weekday,
    hour: rows[0].hour,
    occupancy: average(rows.map(row => row.occupied)) * 100
  })).filter(item => Number.isFinite(item.occupancy));
}

function getColumn(row, possibleNames) {
  const normalizedNames = possibleNames.map(normalize);
  const entry = Object.entries(row).find(([key]) => {
    const cleanedKey = normalize(key);
    return normalizedNames.some(name => cleanedKey.includes(name));
  });

  return entry ? entry[1] : "";
}

function cleanSensorName(value) {
  const sensor = String(value || "").trim();
  return sensor || "Unknown Sensor";
}

function parseStatus(value) {
  const text = normalize(value);
  if (text.includes("occupied")) return "Occupied";
  if (text.includes("available") || text.includes("vacant") || text.includes("free")) return "Available";
  return null;
}

function parseTimestamp(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  const text = String(value || "").trim();
  if (!text) return null;

  // The workbook uses timestamps like "2026-01-14 16:17:29.231331+00:00".
  const match = text.match(/(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2})(?::(\d{2}))?(?::(\d{2}))?/);
  if (match) {
    const [, year, month, day, hour, minute = "0", second = "0"] = match;
    const date = new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function parseDateOnly(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return dateAtStartOfDay(value);
  }

  const text = String(value || "");
  const match = text.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return null;

  const [, year, month, day] = match;
  const date = new Date(Number(year), Number(month) - 1, Number(day));
  return Number.isNaN(date.getTime()) ? null : date;
}

function dateAtStartOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function parseHour(value) {
  if (typeof value === "number") return Math.floor(value);

  const text = String(value || "");
  const match = text.match(/\d{1,2}/);
  if (!match) return null;

  const hour = Number(match[0]);
  return hour >= 0 && hour <= 23 ? hour : null;
}

function parsePercent(value) {
  if (typeof value === "number") return value <= 1 ? value * 100 : value;

  const number = Number(String(value || "").replace("%", "").trim());
  return Number.isFinite(number) ? number : NaN;
}

function averageDailyFromHourly(hourlyItems) {
  const groups = groupBy(hourlyItems, item => formatDate(item.date));

  return Object.entries(groups).map(([dateText, items]) => ({
    date: parseDateOnly(dateText),
    occupancy: average(items.map(item => item.occupancy))
  })).filter(item => item.date).sort((a, b) => a.date - b.date);
}

function averageByWeekday(dailyItems) {
  const groups = {};

  dailyItems.forEach(item => {
    const weekday = WEEKDAYS[item.date.getDay()];
    if (!groups[weekday]) groups[weekday] = [];
    groups[weekday].push(item.occupancy);
  });

  return MONDAY_FIRST.map(day => ({
    weekday: day,
    occupancy: average(groups[day] || [])
  }));
}

function countSensorActivity(rawData) {
  const counts = {};

  rawData.forEach(row => {
    counts[row.sensor] = (counts[row.sensor] || 0) + 1;
  });

  return Object.entries(counts)
    .map(([sensor, changes]) => ({ sensor, changes }))
    .sort((a, b) => b.changes - a.changes);
}

function buildInsights(dailyOccupancy, hourlyOccupancy, rawData) {
  const sensorActivity = countSensorActivity(rawData);
  const weekdayAverages = averageByWeekday(dailyOccupancy);
  const hourlyAverages = averageHourlyDemand(hourlyOccupancy);

  const peakHour = hourlyAverages.reduce((best, item) => {
    return !best || item.occupancy > best.occupancy ? item : best;
  }, null);

  const weekdayValues = weekdayAverages.filter(item => MONDAY_FIRST.slice(0, 5).includes(item.weekday));
  const weekendValues = weekdayAverages.filter(item => ["Saturday", "Sunday"].includes(item.weekday));
  const weekdayAverage = average(weekdayValues.map(item => item.occupancy));
  const weekendAverage = average(weekendValues.map(item => item.occupancy));

  const busiestDay = weekdayAverages.reduce((best, item) => {
    return !best || item.occupancy > best.occupancy ? item : best;
  }, null);

  const topSensors = sensorActivity.slice(0, 3).map(item => item.sensor).join(", ");
  const dailyValues = dailyOccupancy.map(item => item.occupancy);
  const overallAverage = average(dailyValues);

  return {
    peakHour,
    weekdayAverage,
    weekendAverage,
    busiestDay,
    topSensors,
    overallAverage
  };
}

function averageHourlyDemand(hourlyOccupancy) {
  const groups = groupBy(hourlyOccupancy, item => item.hour);

  return Object.entries(groups).map(([hour, items]) => ({
    hour: Number(hour),
    occupancy: average(items.map(item => item.occupancy))
  })).sort((a, b) => a.hour - b.hour);
}

function groupBy(items, keyFunction) {
  return items.reduce((groups, item) => {
    const key = keyFunction(item);
    if (!groups[key]) groups[key] = [];
    groups[key].push(item);
    return groups;
  }, {});
}

function average(values) {
  const validValues = values.filter(Number.isFinite);
  if (!validValues.length) return 0;
  return validValues.reduce((sum, value) => sum + value, 0) / validValues.length;
}

function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function renderDashboard(data) {
  renderDailyTrend(data.dailyOccupancy);
  renderWeekdayChart(data.weekdayOccupancy);
  renderHourlyHeatmap(data.hourlyOccupancy);
  renderSensorChart(data.sensorActivity);
  renderInsights(data.insights);
}

function renderDailyTrend(items) {
  const labels = items.map(item => formatDate(item.date));
  const values = items.map(item => Number(item.occupancy.toFixed(1)));

  replaceChart("dailyTrendChart", {
    type: "line",
    data: {
      labels,
      datasets: [{
        label: "Average occupancy (%)",
        data: values,
        borderColor: "#149cf2",
        backgroundColor: "rgba(20, 156, 242, 0.12)",
        fill: true,
        tension: 0.25,
        spanGaps: false
      }]
    },
    options: baseChartOptions("Occupancy (%)", true)
  });
}

function renderWeekdayChart(items) {
  const values = items.map(item => Number(item.occupancy.toFixed(1)));

  replaceChart("weekdayChart", {
    type: "bar",
    data: {
      labels: items.map(item => item.weekday),
      datasets: [{
        label: "Average occupancy (%)",
        data: values,
        backgroundColor: "#5dcabb",
        borderColor: "#37a999",
        borderWidth: 1
      }]
    },
    options: baseChartOptions("Occupancy (%)", true, values)
  });
}

function renderHourlyHeatmap(items) {
  const container = document.getElementById("hourlyHeatmap");
  container.innerHTML = "";

  if (!items.length) {
    container.innerHTML = '<div class="empty-state">No hourly data was found for this dataset.</div>';
    return;
  }

  const grid = document.createElement("div");
  grid.className = "heatmap-grid";
  grid.appendChild(makeHeatmapLabel(""));

  for (let hour = 0; hour < 24; hour++) {
    grid.appendChild(makeHeatmapLabel(`${hour}:00`));
  }

  const grouped = groupBy(items, item => `${item.weekday}-${item.hour}`);

  MONDAY_FIRST.forEach(day => {
    grid.appendChild(makeHeatmapLabel(day.slice(0, 3)));

    for (let hour = 0; hour < 24; hour++) {
      const value = average((grouped[`${day}-${hour}`] || []).map(item => item.occupancy));
      const cell = document.createElement("div");
      cell.className = "heatmap-cell";
      cell.textContent = value ? Math.round(value) : "";
      cell.title = `${day} ${hour}:00 - ${value.toFixed(1)}% occupancy`;
      cell.style.backgroundColor = heatColor(value);
      grid.appendChild(cell);
    }
  });

  container.appendChild(grid);
}

function makeHeatmapLabel(text) {
  const label = document.createElement("div");
  label.className = "heatmap-label";
  label.textContent = text;
  return label;
}

function heatColor(value) {
  const percent = Math.max(0, Math.min(100, value)) / 100;
  const lightness = 95 - percent * 46;
  return `hsl(166, 55%, ${lightness}%)`;
}

function renderSensorChart(items) {
  const labels = items.map(item => item.sensor);
  const values = items.map(item => item.changes);

  replaceChart("sensorChart", {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: "Status records",
        data: values,
        backgroundColor: "#63ef96",
        borderColor: "#22b86c",
        borderWidth: 1
      }]
    },
    options: {
      ...baseChartOptions("Status records", false, values),
      scales: {
        x: {
          ticks: {
            maxRotation: 45,
            minRotation: 0
          }
        },
        y: {
          beginAtZero: true,
          suggestedMax: Math.max(...values, 1) * 1.18,
          title: {
            display: true,
            text: "Status records"
          }
        }
      }
    }
  });
}

function renderInsights(insights) {
  const crossInsight = document.getElementById("crossInsight");
  const insightsList = document.getElementById("insightsList");
  if (!crossInsight || !insightsList) return;

  const peakHourText = insights.peakHour
    ? `${insights.peakHour.hour}:00, with about ${insights.peakHour.occupancy.toFixed(1)}% average occupancy`
    : "not available from the current data";

  const weekdayWeekendText = insights.weekdayAverage >= insights.weekendAverage
    ? `Weekdays average ${insights.weekdayAverage.toFixed(1)}%, higher than weekends at ${insights.weekendAverage.toFixed(1)}%.`
    : `Weekends average ${insights.weekendAverage.toFixed(1)}%, higher than weekdays at ${insights.weekdayAverage.toFixed(1)}%.`;

  crossInsight.textContent =
    "Together, the daily trend, hourly heatmap, and sensor activity chart show when demand rises and which parking spaces create most of the recorded movement.";

  insightsList.innerHTML = `
    <li>Peak parking demand appears around ${peakHourText}.</li>
    <li>${weekdayWeekendText}</li>
    <li>The most active sensors are ${insights.topSensors || "not available"}.</li>
    <li>Overall average occupancy is about ${insights.overallAverage.toFixed(1)}%, with ${insights.busiestDay ? insights.busiestDay.weekday : "some days"} showing the strongest usage.</li>
  `;
}

function replaceChart(canvasId, config) {
  if (charts[canvasId]) charts[canvasId].destroy();

  const context = document.getElementById(canvasId);
  charts[canvasId] = new Chart(context, config);
}

function baseChartOptions(yTitle, isPercentAxis = false, values = []) {
  const maxValue = values.length ? Math.max(...values, 1) : 100;

  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        display: true
      },
      tooltip: {
        enabled: true,
        callbacks: {
          label(context) {
            const value = context.parsed.y;
            if (isPercentAxis) return `${context.dataset.label}: ${value.toFixed(1)}%`;
            return `${context.dataset.label}: ${value}`;
          }
        }
      }
    },
    scales: {
      y: {
        beginAtZero: true,
        suggestedMax: isPercentAxis ? Math.min(100, Math.max(maxValue * 1.18, 10)) : maxValue * 1.18,
        ticks: {
          callback(value) {
            return isPercentAxis ? `${value}%` : value;
          }
        },
        title: {
          display: true,
          text: yTitle
        }
      }
    }
  };
}

function clearDashboard() {
  Object.values(charts).forEach(chart => chart.destroy());
  charts = {};

  document.getElementById("hourlyHeatmap").innerHTML =
    '<div class="empty-state">Load a dataset to see hourly occupancy.</div>';

  const insightsList = document.getElementById("insightsList");
  if (insightsList) insightsList.innerHTML = "<li>Load a dataset to see parking insights.</li>";
}
