const REQUIRED_COLUMNS = [
  "이름",
  "학교",
  "학년",
  "리포트시점",
  "종합평가",
  "실기",
  "내신",
  "묘사력",
  "형태력",
  "사고력",
  "완성도",
  "멘탈관리",
  "발전 후 지원 가능한 학교",
];

const DRAWING_METRICS = ["형태력", "사고력", "멘탈관리", "완성도", "묘사력"];
const LINE_COLORS = ["#3c2c1f", "#b77a45", "#447487", "#68745b", "#9b6b7b"];

const fileInput = document.getElementById("inputFile");
const selectFileButton = document.getElementById("selectFileButton");
const printButton = document.getElementById("printButton");
const statusEl = document.getElementById("status");
const reportPage = document.getElementById("reportPage");

selectFileButton.addEventListener("click", () => fileInput.click());
printButton.addEventListener("click", () => window.print());
fileInput.addEventListener("change", handleFileSelection);

async function handleFileSelection(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    if (!window.XLSX) {
      throw new Error("XLSX 파서가 로드되지 않았습니다. 인터넷 연결 또는 CDN 접근을 확인하세요.");
    }

    setStatus(`선택한 파일: ${file.name}`);
    const rows = await readWorkbook(file);
    validateTemplate(rows);

    const normalizedRows = rows
      .map(normalizeRow)
      .filter((row) => row["리포트시점"] instanceof Date && !Number.isNaN(row["리포트시점"].getTime()));

    if (normalizedRows.length === 0) {
      throw new Error("읽을 수 있는 리포트 행이 없습니다.");
    }

    reportPage.hidden = false;
    renderReport(normalizedRows);
    printButton.disabled = false;

    const first = normalizedRows[0];
    const last = normalizedRows.at(-1);
    const studentName = valueWithFallback(last["이름"], first["이름"]) || "학생";
    setStatus(`${studentName} ${last["리포트시점"].getFullYear()}년 ${last["리포트시점"].getMonth() + 1}월 리포트를 생성했습니다.`);
  } catch (error) {
    reportPage.hidden = true;
    printButton.disabled = true;
    setStatus(`오류 발생: ${error.message}`);
    console.error(error);
  } finally {
    event.target.value = "";
  }
}

function setStatus(message) {
  statusEl.textContent = message;
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
}

function validateTemplate(rows) {
  if (!rows.length) {
    throw new Error("엑셀 파일에 데이터가 없습니다.");
  }

  const columns = Object.keys(rows[0]);
  if (columns.includes("성적")) {
    throw new Error("예중/예고 입시용 템플릿이 아닌 미대 입시용 템플릿을 불러들였습니다.");
  }

  const missing = REQUIRED_COLUMNS.filter((column) => !columns.includes(column));
  if (missing.length) {
    throw new Error(`필수 컬럼이 없습니다: ${missing.join(", ")}`);
  }
}

function normalizeRow(row) {
  const normalized = { ...row };
  normalized["리포트시점"] = parseExcelDate(row["리포트시점"]);

  ["학년", "실기", "내신", ...DRAWING_METRICS].forEach((key) => {
    normalized[key] = toNumber(row[key]);
  });

  DRAWING_METRICS.forEach((key) => {
    if (normalized[key] === 0) normalized[key] = null;
  });

  return normalized;
}

function parseExcelDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number") return excelSerialToDate(value);

  const text = String(value || "").trim();
  if (!text) return new Date(Number.NaN);

  const direct = new Date(text);
  if (!Number.isNaN(direct.getTime())) return direct;

  const numeric = Number(text);
  if (Number.isFinite(numeric)) return excelSerialToDate(numeric);

  return new Date(Number.NaN);
}

function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const date = new Date(utcValue * 1000);
  return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

function toNumber(value) {
  if (typeof value === "number") return value;
  const parsed = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(parsed) ? parsed : 0;
}

function hasValue(value) {
  return typeof value === "number" ? Number.isFinite(value) && value > 0 : String(value ?? "").trim() !== "";
}

function valueWithFallback(value, fallback) {
  return hasValue(value) ? value : fallback;
}

function renderReport(rows) {
  const first = rows[0];
  const last = rows.at(-1);
  const studentName = valueWithFallback(last["이름"], first["이름"]);
  const schoolName = valueWithFallback(last["학교"], first["학교"]);
  const grade = valueWithFallback(last["학년"], first["학년"]);

  document.getElementById("studentName").textContent = studentName || "";
  document.getElementById("schoolName").textContent = schoolName || "";
  document.getElementById("grade").textContent = formatGrade(grade);
  document.getElementById("studyPeriod").textContent = `${calendarMonthSpan(first["리포트시점"], last["리포트시점"])}개월 차`;
  document.getElementById("currentSchools").textContent = last["현재 상황 지원 가능한 학교"] || "기본 다지는 중";
  document.getElementById("futureSchools").textContent = last["발전 후 지원 가능한 학교"] || "";
  document.getElementById("evaluationText").textContent = last["종합평가"] || "";

  drawGradesChart(document.getElementById("gradesChart"), last);
  drawPentagonChart(document.getElementById("pentagonChart"), last);
  drawMonthlyChart(document.getElementById("monthlyChart"), rows);
}

function formatGrade(value) {
  return Number.isFinite(value) && value > 0 ? `${value}` : "";
}

function calendarMonthSpan(startDate, endDate) {
  return (endDate.getFullYear() - startDate.getFullYear()) * 12 + endDate.getMonth() - startDate.getMonth() + 1;
}

function setupCanvas(canvas) {
  const rect = canvas.getBoundingClientRect();
  const ratio = window.devicePixelRatio || 1;
  const cssWidth = rect.width || canvas.width;
  const cssHeight = rect.height || canvas.height;
  canvas.width = Math.round(cssWidth * ratio);
  canvas.height = Math.round(cssHeight * ratio);
  const ctx = canvas.getContext("2d");
  ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
  ctx.clearRect(0, 0, cssWidth, cssHeight);
  ctx.font = '12px "Malgun Gothic", Arial, sans-serif';
  ctx.fillStyle = "#111";
  ctx.strokeStyle = "#111";
  return { ctx, width: cssWidth, height: cssHeight };
}

function drawGradesChart(canvas, row) {
  const { ctx, width, height } = setupCanvas(canvas);
  const margin = { top: 42, right: 12, bottom: 36, left: 38 };
  const plotW = width - margin.left - margin.right;
  const plotH = height - margin.top - margin.bottom;
  const values = [row["실기"], row["내신"]];
  const labels = ["실기", "내신"];

  drawTitle(ctx, width / 2, 14, "나의 현재 입시 요소 상태");
  drawTitle(ctx, width / 2, 31, "(0점: 지원 불가, 100점: 지원 안정)");
  drawYAxis(ctx, margin, plotW, plotH, 0, 100);

  const barW = 42;
  values.forEach((value, index) => {
    const x = margin.left + plotW * (index + 0.5) / values.length - barW / 2;
    const y = margin.top + plotH - (clamp(value, 0, 100) / 100) * plotH;
    ctx.fillStyle = "#b77a45";
    ctx.fillRect(x, y, barW, margin.top + plotH - y);
    ctx.fillStyle = "#111";
    ctx.textAlign = "center";
    ctx.fillText(labels[index], x + barW / 2, height - 12);
  });
}

function drawPentagonChart(canvas, row) {
  const { ctx, width, height } = setupCanvas(canvas);
  const center = { x: width / 2, y: height * 0.58 };
  const radius = Math.min(width, height) * 0.27;
  const labelRadius = radius + 22;
  const labelPoints = pentagonPoints(center, labelRadius);

  drawTitle(ctx, width / 2, 18, "현재 나의 실기요소 상태");

  ctx.strokeStyle = "#111";
  ctx.lineWidth = 1;
  [1, 0.8, 0.6, 0.4, 0.2].forEach((scale, index) => {
    drawPolygonPath(ctx, center, radius * scale);
    if (index > 0) ctx.setLineDash([4, 4]);
    ctx.stroke();
    ctx.setLineDash([]);
  });

  ctx.font = '12px "Malgun Gothic", Arial, sans-serif';
  ctx.fillStyle = "#111";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  [
    ["형태력", labelPoints[0].x, labelPoints[0].y],
    ["사고력", Math.min(width - 26, labelPoints[1].x), labelPoints[1].y],
    ["멘탈관리", labelPoints[2].x + 8, Math.min(height - 18, labelPoints[2].y + 4)],
    ["완성도", labelPoints[3].x - 8, Math.min(height - 18, labelPoints[3].y + 4)],
    ["묘사력", Math.max(26, labelPoints[4].x), labelPoints[4].y],
  ].forEach(([text, x, y]) => {
    ctx.fillText(text, x, y);
  });

  const points = DRAWING_METRICS.map((metric, index) => {
    const value = row[metric];
    if (value === null || Number.isNaN(value)) return null;
    const angle = (-90 + index * 72) * Math.PI / 180;
    return {
      x: center.x + Math.cos(angle) * radius * clamp(value, 0, 100) / 100,
      y: center.y + Math.sin(angle) * radius * clamp(value, 0, 100) / 100,
    };
  });

  ctx.strokeStyle = "#3c2c1f";
  ctx.fillStyle = "#3c2c1f";
  ctx.lineWidth = 2;
  for (let index = 0; index < points.length; index += 1) {
    const current = points[index];
    const next = points[(index + 1) % points.length];
    if (!current || !next) continue;
    ctx.beginPath();
    ctx.moveTo(current.x, current.y);
    ctx.lineTo(next.x, next.y);
    ctx.stroke();
  }

  points.filter(Boolean).forEach((point) => {
    ctx.beginPath();
    ctx.arc(point.x, point.y, 3.5, 0, Math.PI * 2);
    ctx.fill();
  });
}

function pentagonPoints(center, radius) {
  return Array.from({ length: 5 }, (_, index) => {
    const angle = (-90 + index * 72) * Math.PI / 180;
    return {
      x: center.x + Math.cos(angle) * radius,
      y: center.y + Math.sin(angle) * radius,
    };
  });
}

function drawMonthlyChart(canvas, rows) {
  const { ctx, width, height } = setupCanvas(canvas);
  const margin = { top: 34, right: 94, bottom: 38, left: 38 };
  const plotW = width - margin.left - margin.right;
  const plotH = height - margin.top - margin.bottom;
  const months = rows.map((row) => startOfMonth(row["리포트시점"]));

  drawTitle(ctx, margin.left + plotW / 2, 18, "월별 실기요소 변화 그래프");
  drawYAxis(ctx, margin, plotW, plotH, 0, 100);

  ctx.textAlign = "center";
  ctx.fillStyle = "#111";
  months.forEach((date, index) => {
    const x = xAt(index, rows.length, margin.left, plotW);
    ctx.strokeStyle = "#d8d8d8";
    ctx.beginPath();
    ctx.moveTo(x, margin.top);
    ctx.lineTo(x, margin.top + plotH);
    ctx.stroke();
    ctx.fillText(monthLabel(date, index), x, height - 14);
  });

  DRAWING_METRICS.forEach((metric, metricIndex) => {
    const color = LINE_COLORS[metricIndex];
    ctx.strokeStyle = color;
    ctx.fillStyle = color;
    ctx.lineWidth = 2;
    ctx.beginPath();
    let isDrawing = false;

    rows.forEach((row, rowIndex) => {
      const value = row[metric];
      if (value === null || Number.isNaN(value)) {
        isDrawing = false;
        return;
      }
      const x = xAt(rowIndex, rows.length, margin.left, plotW);
      const y = margin.top + plotH - (clamp(value, 0, 100) / 100) * plotH;
      if (!isDrawing) {
        ctx.moveTo(x, y);
        isDrawing = true;
      } else {
        ctx.lineTo(x, y);
      }
    });
    ctx.stroke();
  });

  const endValues = DRAWING_METRICS.map((metric) => rows.at(-1)[metric]);
  const duplicateRank = markerRanks(endValues);
  DRAWING_METRICS.forEach((metric, metricIndex) => {
    const value = rows.at(-1)[metric];
    if (value === null || Number.isNaN(value)) return;
    const x = xAt(rows.length - 1, rows.length, margin.left, plotW);
    const y = margin.top + plotH - (clamp(value, 0, 100) / 100) * plotH;
    ctx.fillStyle = LINE_COLORS[metricIndex];
    ctx.beginPath();
    ctx.arc(x, y, 3 + duplicateRank[metricIndex] * 2, 0, Math.PI * 2);
    ctx.fill();
  });

  drawLegend(ctx, width - 84, margin.top + 22);
}

function drawTitle(ctx, x, y, text) {
  ctx.fillStyle = "#111";
  ctx.textAlign = "center";
  ctx.font = '13px "Malgun Gothic", Arial, sans-serif';
  ctx.fillText(text, x, y);
}

function drawYAxis(ctx, margin, plotW, plotH, min, max) {
  ctx.strokeStyle = "#111";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(margin.left, margin.top);
  ctx.lineTo(margin.left, margin.top + plotH);
  ctx.lineTo(margin.left + plotW, margin.top + plotH);
  ctx.stroke();

  ctx.font = '11px "Malgun Gothic", Arial, sans-serif';
  ctx.textAlign = "right";
  ctx.fillStyle = "#111";
  for (let tick = min; tick <= max; tick += 20) {
    const y = margin.top + plotH - ((tick - min) / (max - min)) * plotH;
    ctx.strokeStyle = "#d8d8d8";
    ctx.beginPath();
    ctx.moveTo(margin.left, y);
    ctx.lineTo(margin.left + plotW, y);
    ctx.stroke();
    ctx.fillText(String(tick), margin.left - 6, y + 4);
  }
}

function drawPolygonPath(ctx, center, radius) {
  ctx.beginPath();
  for (let index = 0; index < 5; index += 1) {
    const angle = (-90 + index * 72) * Math.PI / 180;
    const x = center.x + Math.cos(angle) * radius;
    const y = center.y + Math.sin(angle) * radius;
    if (index === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  }
  ctx.closePath();
}

function drawLegend(ctx, x, y) {
  ctx.font = '11px "Malgun Gothic", Arial, sans-serif';
  ctx.textAlign = "left";
  DRAWING_METRICS.forEach((metric, index) => {
    const yy = y + index * 18;
    ctx.strokeStyle = LINE_COLORS[index];
    ctx.fillStyle = LINE_COLORS[index];
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(x, yy - 4);
    ctx.lineTo(x + 18, yy - 4);
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(x + 9, yy - 4, 3, 0, Math.PI * 2);
    ctx.fill();
    ctx.fillStyle = "#111";
    ctx.fillText(metric, x + 24, yy);
  });
}

function xAt(index, count, left, width) {
  if (count <= 1) return left + width / 2;
  return left + (width * index) / (count - 1);
}

function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function monthLabel(date, index) {
  const month = `${date.getMonth() + 1}월`;
  if (index === 0 || date.getMonth() === 0) return `${date.getFullYear()}년 ${month}`;
  return month;
}

function markerRanks(values) {
  const ranks = new Array(values.length).fill(0);
  const groups = new Map();
  values.forEach((value, index) => {
    if (value === null || Number.isNaN(value)) return;
    const key = String(value);
    groups.set(key, [...(groups.get(key) || []), index]);
  });

  groups.forEach((indexes) => {
    indexes.slice().reverse().forEach((metricIndex, rank) => {
      ranks[metricIndex] = rank;
    });
  });
  return ranks;
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, Number(value) || 0));
}
