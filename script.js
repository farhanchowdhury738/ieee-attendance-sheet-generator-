const entries = [];

const dateInput = document.getElementById("dateInput");
const timeInput = document.getElementById("timeInput");
const nameInput = document.getElementById("nameInput");
const idInput = document.getElementById("idInput");
const titleInput = document.getElementById("titleInput");
const addBtn = document.getElementById("addBtn");
const clearBtn = document.getElementById("clearBtn");
const downloadBtn = document.getElementById("downloadBtn");
const downloadWordBtn = document.getElementById("downloadWordBtn");
const eventBtn = document.getElementById("eventBtn");
const eventBtnLabel = document.getElementById("eventBtnLabel");
const participantBtn = document.getElementById("participantBtn");
const dateControl = document.getElementById("dateControl");
const entryControls = document.getElementById("entryControls");
const entryActionButtons = document.getElementById("entryActionButtons");
const entriesSection = document.getElementById("entriesSection");
const entryList = document.getElementById("entryList");
const countBadge = document.getElementById("countBadge");
const tableWrap = document.getElementById("tableWrap");
const sheetTitle = document.getElementById("sheetTitle");
const brandTitle = document.getElementById("brandTitle");
const previewTitle = document.getElementById("previewTitle");
const participantDate = document.getElementById("participantDate");
const statusBox = document.getElementById("status");

const DEFAULT_BRAND_TITLE = "IEEE AIUB Student Branch";

/* =========================================================
   EXPORT GEOMETRY HELPERS

   The live preview is the SINGLE SOURCE OF TRUTH.
   Every exported file (PDF and Word) is rebuilt from the
   measurements and computed styles of the elements that are
   on screen right now, so nothing is hard-coded twice.

   These helpers are only used while generating a download.
   They never touch the live preview markup or styling.
   ========================================================= */

// The preview page is an exact A4 sheet at 96 DPI (794 x 1123 px).
const A4_WIDTH_PX = 794;
const A4_WIDTH_MM = 210;
const A4_HEIGHT_MM = 297;
const A4_WIDTH_TWIP = 11906; // 210mm
const A4_HEIGHT_TWIP = 16838; // 297mm

// 1 CSS px = 1/96 in = 0.75 pt = 15 twips = 1.5 half-points = 6 eighth-points
const pxToTwip = (px) => Math.round((Number(px) || 0) * 15);
const pxToHalfPoint = (px) => Math.max(2, Math.round((Number(px) || 0) * 1.5));
const pxToEighthPoint = (px) => Math.max(1, Math.round((Number(px) || 0) * 6));
const pxToMm = (px) => ((Number(px) || 0) * A4_WIDTH_MM) / A4_WIDTH_PX;

function cssNumber(value) {
  const n = parseFloat(value);
  return Number.isFinite(n) ? n : 0;
}

function cssColorToHex(value) {
  const m = String(value).match(/rgba?\(([^)]+)\)/i);
  if (!m) return "000000";

  const parts = m[1].split(",").map((n) => parseFloat(n));
  const [r, g, b] = parts;

  // Fully transparent borders should not be drawn as black.
  if (parts.length > 3 && parts[3] === 0) return "FFFFFF";

  return [r, g, b]
    .map((n) =>
      Math.max(0, Math.min(255, Math.round(Number(n) || 0)))
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase();
}

function cssFirstFontFamily(value) {
  const first = String(value || "")
    .split(",")[0]
    .trim()
    .replace(/^["']|["']$/g, "");

  return first || "Arial";
}

function cssIsBold(weight) {
  if (weight === "bold" || weight === "bolder") return true;
  return (parseInt(weight, 10) || 400) >= 600;
}

function cssAlignment(value, AlignmentType) {
  switch (String(value || "").toLowerCase()) {
    case "center":
      return AlignmentType.CENTER;
    case "right":
    case "end":
      return AlignmentType.RIGHT;
    case "justify":
      return AlignmentType.JUSTIFIED;
    default:
      return AlignmentType.LEFT;
  }
}

function cssVerticalAlign(value, VerticalAlign) {
  switch (String(value || "").toLowerCase()) {
    case "top":
      return VerticalAlign.TOP;
    case "bottom":
      return VerticalAlign.BOTTOM;
    default:
      return VerticalAlign.CENTER;
  }
}

function cssLineSpacing(computed) {
  const lineHeight = parseFloat(computed.lineHeight);
  if (!Number.isFinite(lineHeight)) return undefined;

  return { line: pxToTwip(lineHeight), lineRule: "atLeast" };
}

/*
 * Force the preview to its real A4 width while a file is being
 * generated, so a phone produces exactly the same geometry as a PC.
 * Returns a function that restores the original inline styles.
 */
function lockPreviewToExportWidth(page) {
  const previous = {
    width: page.style.width,
    minWidth: page.style.minWidth,
  };

  page.style.width = `${A4_WIDTH_PX}px`;
  page.style.minWidth = `${A4_WIDTH_PX}px`;

  return () => {
    page.style.width = previous.width;
    page.style.minWidth = previous.minWidth;
  };
}

function isVisibleBlock(el) {
  if (!el) return false;

  const cs = window.getComputedStyle(el);
  if (cs.display === "none" || cs.visibility === "hidden") return false;

  return el.textContent.trim().length > 0;
}

function cleanText(el) {
  return el.textContent.replace(/\s+/g, " ").trim();
}

/* ---------------------------------------------------------
   DOM -> docx builders
   --------------------------------------------------------- */

/*
 * Convert one <th>/<td> from the live preview into a Word table
 * cell that keeps its width, padding, borders, font, horizontal
 * alignment and vertical alignment.
 */
function domCellToDocxCell(cell, widthTwip, docxLib) {
  const {
    TableCell,
    Paragraph,
    TextRun,
    WidthType,
    VerticalAlign,
    AlignmentType,
    BorderStyle,
  } = docxLib;

  const cs = window.getComputedStyle(cell);

  const makeBorder = (widthValue, colorValue, styleValue) => {
    const width = cssNumber(widthValue);
    if (!width || styleValue === "none" || styleValue === "hidden") {
      return { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    }
    return {
      style: BorderStyle.SINGLE,
      size: pxToEighthPoint(width),
      color: cssColorToHex(colorValue),
    };
  };

  const text = cleanText(cell);
  const fontSizePx = cssNumber(cs.fontSize) || 12;

  return new TableCell({
    width: { size: widthTwip, type: WidthType.DXA },
    columnSpan: cell.colSpan > 1 ? cell.colSpan : undefined,
    rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
    verticalAlign: cssVerticalAlign(cs.verticalAlign, VerticalAlign),
    margins: {
      top: pxToTwip(cssNumber(cs.paddingTop)),
      bottom: pxToTwip(cssNumber(cs.paddingBottom)),
      left: pxToTwip(cssNumber(cs.paddingLeft)),
      right: pxToTwip(cssNumber(cs.paddingRight)),
    },
    borders: {
      top: makeBorder(cs.borderTopWidth, cs.borderTopColor, cs.borderTopStyle),
      bottom: makeBorder(
        cs.borderBottomWidth,
        cs.borderBottomColor,
        cs.borderBottomStyle,
      ),
      left: makeBorder(
        cs.borderLeftWidth,
        cs.borderLeftColor,
        cs.borderLeftStyle,
      ),
      right: makeBorder(
        cs.borderRightWidth,
        cs.borderRightColor,
        cs.borderRightStyle,
      ),
    },
    children: [
      new Paragraph({
        alignment: cssAlignment(cs.textAlign, AlignmentType),
        spacing: Object.assign(
          { before: 0, after: 0 },
          cssLineSpacing(cs) || {},
        ),
        children: [
          new TextRun({
            text,
            bold: cssIsBold(cs.fontWeight),
            size: pxToHalfPoint(fontSizePx),
            color: cssColorToHex(cs.color),
            font: cssFirstFontFamily(cs.fontFamily),
          }),
        ],
      }),
    ],
  });
}

/*
 * Rebuild the whole preview table in Word using the exact column
 * widths, row heights and cell styles that are on screen.
 */
function domTableToDocxTable(tableEl, page, docxLib) {
  const { Table, TableRow, WidthType, HeightRule, TableLayoutType } = docxLib;

  const headerCells = Array.from(
    tableEl.querySelectorAll("thead tr:first-child > *"),
  );

  const measuredCells = headerCells.length
    ? headerCells
    : Array.from(tableEl.querySelectorAll("tr:first-child > *"));

  const columnWidths = measuredCells.map((cell) =>
    pxToTwip(cell.getBoundingClientRect().width),
  );

  const totalWidthTwip = columnWidths.reduce((sum, w) => sum + w, 0);

  const pageRect = page.getBoundingClientRect();
  const tableRect = tableEl.getBoundingClientRect();
  const indentTwip = pxToTwip(Math.max(0, tableRect.left - pageRect.left));

  const columnCount = columnWidths.length;

  // Tracks how many rows are still covered by a rowspan started above,
  // so a cell that follows a merged cell keeps the correct column width.
  const rowSpanRemaining = new Array(columnCount).fill(0);

  const rows = Array.from(tableEl.rows).map((row) => {
    const isHeaderRow = row.parentElement.tagName === "THEAD";

    let columnIndex = 0;
    const cells = Array.from(row.cells).map((cell) => {
      // Skip the columns that a rowspan from a previous row still covers.
      while (columnIndex < columnCount && rowSpanRemaining[columnIndex] > 0) {
        columnIndex += 1;
      }

      const colSpan = cell.colSpan > 1 ? cell.colSpan : 1;
      const rowSpan = cell.rowSpan > 1 ? cell.rowSpan : 1;

      let widthTwip = 0;
      for (let i = 0; i < colSpan; i++) {
        widthTwip += columnWidths[columnIndex + i] || 0;
        rowSpanRemaining[columnIndex + i] = rowSpan;
      }
      columnIndex += colSpan;

      if (!widthTwip) {
        widthTwip = pxToTwip(cell.getBoundingClientRect().width);
      }

      return domCellToDocxCell(cell, widthTwip, docxLib);
    });

    for (let i = 0; i < columnCount; i++) {
      if (rowSpanRemaining[i] > 0) rowSpanRemaining[i] -= 1;
    }

    return new TableRow({
      tableHeader: isHeaderRow,
      cantSplit: true,
      height: {
        value: pxToTwip(row.getBoundingClientRect().height),
        rule: HeightRule.ATLEAST,
      },
      children: cells,
    });
  });

  return new Table({
    layout: TableLayoutType.FIXED,
    columnWidths,
    width: { size: totalWidthTwip, type: WidthType.DXA },
    indent: { size: indentTwip, type: WidthType.DXA },
    rows,
  });
}

/*
 * Convert a heading/date line from the preview into a Word paragraph
 * that keeps its font, weight, alignment, indents and the vertical
 * gap that sits above it in the preview.
 */
function domBlockToDocxParagraph(el, page, gaps, docxLib) {
  const { Paragraph, TextRun, AlignmentType } = docxLib;

  const cs = window.getComputedStyle(el);
  const pageRect = page.getBoundingClientRect();
  const rect = el.getBoundingClientRect();

  const indentLeft = Math.max(
    0,
    rect.left - pageRect.left + cssNumber(cs.paddingLeft),
  );
  const indentRight = Math.max(
    0,
    pageRect.right - rect.right + cssNumber(cs.paddingRight),
  );

  return new Paragraph({
    alignment: cssAlignment(cs.textAlign, AlignmentType),
    indent: {
      left: pxToTwip(indentLeft),
      right: pxToTwip(indentRight),
    },
    spacing: Object.assign(
      {
        before: pxToTwip(
          Math.max(0, (gaps && gaps.before) || 0) + cssNumber(cs.paddingTop),
        ),
        after: pxToTwip(
          Math.max(0, (gaps && gaps.after) || 0) + cssNumber(cs.paddingBottom),
        ),
      },
      cssLineSpacing(cs) || {},
    ),
    children: [
      new TextRun({
        text: cleanText(el),
        bold: cssIsBold(cs.fontWeight),
        size: pxToHalfPoint(cssNumber(cs.fontSize) || 12),
        color: cssColorToHex(cs.color),
        font: cssFirstFontFamily(cs.fontFamily),
      }),
    ],
  });
}

// Preview mode: "attendance" (default), "event", or "participant".
let previewMode = "attendance";

const PARTICIPANT_ROW_COUNT = 24;

// Default date = today.
const today = new Date();
dateInput.value = [
  today.getFullYear(),
  String(today.getMonth() + 1).padStart(2, "0"),
  String(today.getDate()).padStart(2, "0"),
].join("-");

function escapeHtml(value) {
  return String(value).replace(
    /[&<>"']/g,
    (c) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;",
      })[c],
  );
}

function parseTimeRange(raw) {
  const cleaned = raw.trim().replace(/\s+/g, "").replace(/[–—]/g, "-");

  const parts = cleaned.split("-");
  if (parts.length !== 2) return null;

  const parseOne = (part) => {
    const m = part.match(/^(\d{1,2})(?:(:|\.)(\d{1,2}))?(AM|PM)?$/i);

    if (!m) return null;

    let hour = Number(m[1]);
    let minute = Number(m[3] || 0);
    const mer = m[4] ? m[4].toUpperCase() : null;

    if (minute > 59) return null;

    if (mer) {
      if (hour < 1 || hour > 12) return null;

      if (mer === "AM" && hour === 12) hour = 0;
      if (mer === "PM" && hour !== 12) hour += 12;
    } else {
      if (hour > 23) return null;
    }

    return { hour, minute };
  };

  const start = parseOne(parts[0]);
  const end = parseOne(parts[1]);

  if (!start || !end) return null;

  return { start, end };
}

function format12(time) {
  let h = time.hour;
  const m = String(time.minute).padStart(2, "0");
  const suffix = h >= 12 ? "PM" : "AM";
  h = h % 12 || 12;
  return `${String(h).padStart(2, "0")}:${m} ${suffix}`;
}

function format24(time) {
  const hour = String(time.hour).padStart(2, "0");
  const minute = String(time.minute).padStart(2, "0");

  return `${hour}:${minute}`;
}

function formatTimeRange(parsed) {
  return `${format24(parsed.start)}-${format24(parsed.end)}`;
}

function formatSlot(raw) {
  return raw.trim().replace(/\s*-\s*/g, "-");
}

function formatDateLong(dateString) {
  if (!dateString) return "Attendance Sheet";
  const d = new Date(dateString + "T00:00:00");
  const weekday = d.toLocaleDateString("en-US", { weekday: "long" });
  const month = d.toLocaleDateString("en-US", { month: "long" });
  return `Attendance Sheet ${weekday} ${month} ${d.getDate()}, ${d.getFullYear()}`;
}

function formatDateSlash(dateString) {
  if (!dateString) return "";
  const [year, month, day] = dateString.split("-");
  if (!year || !month || !day) return "";
  return `${day}/${month}/${year}`;
}

function setStatus(message, type = "success") {
  statusBox.textContent = message;
  statusBox.className =
    "mt-4 rounded-xl px-4 py-3 text-sm font-medium " +
    (type === "error"
      ? "bg-red-50 text-red-700 border border-red-200"
      : "bg-emerald-50 text-emerald-700 border border-emerald-200");
}

function renderEntries() {
  countBadge.textContent = `${entries.length} ${entries.length === 1 ? "entry" : "entries"}`;

  if (!entries.length) {
    entryList.innerHTML = `<div class="rounded-xl border border-dashed border-slate-300 p-4 text-center text-sm text-slate-400">No entries yet.</div>`;
    return;
  }

  entryList.innerHTML = entries
    .map(
      (e, i) => `
      <div class="flex items-center gap-3 rounded-xl border border-slate-200 bg-slate-50 px-3 py-2.5">
        <div class="min-w-0 flex-1">
          <div class="truncate text-sm font-semibold text-slate-800">${escapeHtml(e.name)}${e.id ? ` <span class="font-normal text-slate-400">(${escapeHtml(e.id)})</span>` : ""}</div>
          <div class="text-xs text-slate-500">${escapeHtml(e.slot)} • ${escapeHtml(e.reporting)}</div>
        </div>
        <button data-delete="${i}" class="rounded-lg px-2.5 py-1.5 text-xs font-bold text-red-600 hover:bg-red-50">Delete</button>
      </div>
    `,
    )
    .join("");

  entryList.querySelectorAll("[data-delete]").forEach((btn) => {
    btn.addEventListener("click", () => {
      entries.splice(Number(btn.dataset.delete), 1);
      renderEntries();
      renderTable();
    });
  });
}

function renderTable() {
  if (previewMode === "participant") {
    sheetTitle.textContent = "";
    participantDate.style.display = "block";
    renderParticipantTable();
    return;
  }

  participantDate.style.display = "none";
  sheetTitle.textContent = formatDateLong(dateInput.value);

  if (!entries.length) {
    tableWrap.innerHTML = `<div class="empty-state">Add at least one entry from the left panel to build the ${
      previewMode === "event" ? "event" : "attendance"
    } table.</div>`;
    return;
  }

  if (previewMode === "event") {
    renderEventTable();
    return;
  }

  // Keep first-seen slot order, while grouping same slots together.
  const groups = [];
  const map = new Map();

  entries.forEach((e) => {
    if (!map.has(e.slot)) {
      const group = { slot: e.slot, people: [] };
      map.set(e.slot, group);
      groups.push(group);
    }
    map.get(e.slot).people.push(e);
  });

  let rows = "";
  groups.forEach((group) => {
    group.people.forEach((person, idx) => {
      rows += `<tr>`;
      if (idx === 0) {
        rows += `<td class="slot-cell" rowspan="${group.people.length}">${escapeHtml(group.slot)}</td>`;
      }
      rows += `
          <td>${escapeHtml(person.name)}</td>
          <td>${escapeHtml(person.reporting)}</td>
          <td></td>
          <td></td>
          <td></td>
        </tr>`;
    });
  });

  tableWrap.innerHTML = `
      <table class="attendance-table">
        <colgroup>
          <col class="slot"><col class="name"><col class="reporting">
          <col class="checkin"><col class="checkout"><col class="signature">
        </colgroup>
        <thead>
          <tr>
            <th>Slot Time</th>
            <th>Name</th>
            <th>Reporting Time</th>
            <th>Check In Time</th>
            <th>Check Out Time</th>
            <th>Signature</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    `;
}

function renderEventTable() {
  let rows = "";

  entries.forEach((person) => {
    rows += `
        <tr>
          <td>${escapeHtml(person.name)}</td>
          <td>${escapeHtml(person.id || "")}</td>
          <td></td>
          <td></td>
        </tr>`;
  });

  tableWrap.innerHTML = `
      <table class="attendance-table">
        <colgroup>
          <col class="event-name"><col class="event-id">
          <col class="event-reporting"><col class="event-signature">
        </colgroup>
        <thead>
          <tr>
            <th>Name</th>
            <th>ID</th>
            <th>Reporting Time</th>
            <th>Signature</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    `;
}

function renderParticipantTable() {
  let rows = "";
  for (let i = 0; i < PARTICIPANT_ROW_COUNT; i++) {
    rows += `
      <tr>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
      </tr>`;
  }

  const participantDateText = formatDateSlash(dateInput.value);
  participantDate.textContent = `Date: ${participantDateText}`;
  participantDate.style.display = "block";

  tableWrap.innerHTML = `
    <table class="attendance-table participant-table">
      <colgroup>
        <col class="participant-name">
        <col class="participant-email">
        <col class="participant-contact">
        <col class="participant-dep">
      </colgroup>
      <thead>
        <tr>
          <th>Name</th>
          <th>Email</th>
          <th>Contact No</th>
          <th>Dep</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function updateControlsForMode() {
  const participant = previewMode === "participant";

  dateControl.style.display = "block";
  entryControls.style.display = "";
  entryActionButtons.style.display = "";
  entriesSection.style.display = "";

  if (participant) {
    previewTitle.textContent = "Live Participant Preview";
    brandTitle.textContent = titleInput.value.trim() || DEFAULT_BRAND_TITLE;
    sheetTitle.style.display = "none";
    eventBtnLabel.textContent = "Event";
  } else if (previewMode === "event") {
    previewTitle.textContent = "Live Event Preview";
    sheetTitle.style.display = "block";
    brandTitle.textContent = titleInput.value.trim() || DEFAULT_BRAND_TITLE;
  } else {
    previewTitle.textContent = "Live Attendance Preview";
    sheetTitle.style.display = "block";
    brandTitle.textContent = DEFAULT_BRAND_TITLE;
  }
}

function addEntry() {
  if (previewMode === "participant") return;

  const rawTime = timeInput.value.trim();
  const name = nameInput.value.trim();
  const id = idInput.value.trim();
  const title = titleInput.value.trim();

  if (previewMode === "event") {
    if (!name || !id || !title) {
      setStatus("Please fill up Name, ID and Title.", "error");
      return;
    }
  } else {
    if (!rawTime || !name) {
      setStatus("Please enter both Time Range and Name.", "error");
      return;
    }
  }

  let slot = "";
  let reporting = "";

  if (rawTime) {
    const parsed = parseTimeRange(rawTime);
    if (!parsed) {
      setStatus("Invalid time range. Example: 9:00-10:00", "error");
      return;
    }

    slot = formatTimeRange(parsed);
    reporting = format12(parsed.start);
  }

  entries.push({ slot, name, id, reporting });

  timeInput.value = "";
  nameInput.value = "";
  idInput.value = "";
  timeInput.focus();

  renderEntries();
  renderTable();
  setStatus("Entry added successfully.");
}

addBtn.addEventListener("click", addEntry);

[timeInput, nameInput, idInput, titleInput].forEach((input) => {
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") addEntry();
  });
});

dateInput.addEventListener("change", renderTable);

clearBtn.addEventListener("click", () => {
  entries.length = 0;
  renderEntries();
  renderTable();
  setStatus("All entries cleared.");
});

eventBtn.addEventListener("click", () => {
  previewMode = previewMode === "event" ? "attendance" : "event";

  eventBtnLabel.textContent = previewMode === "event" ? "Attendance" : "Event";
  updateControlsForMode();
  renderTable();
});

participantBtn.addEventListener("click", () => {
  previewMode = previewMode === "participant" ? "attendance" : "participant";
  updateControlsForMode();
  renderTable();
});

// Keep the title in the participant/event preview in sync while typing.
titleInput.addEventListener("input", () => {
  if (previewMode === "event" || previewMode === "participant") {
    brandTitle.textContent = titleInput.value.trim() || DEFAULT_BRAND_TITLE;
  }
});

downloadBtn.addEventListener("click", async () => {
  if (!entries.length && previewMode !== "participant") {
    setStatus("Add at least one entry before downloading the PDF.", "error");
    return;
  }

  const originalText = downloadBtn.textContent;
  downloadBtn.disabled = true;
  downloadBtn.textContent = "Generating PDF...";

  const page = document.getElementById("pdfPage");

  // Save current mobile styles
  const originalWidth = page.style.width;
  const originalMinWidth = page.style.minWidth;

  page.classList.add("pdf-capture");

  try {
    /*
     * IMPORTANT:
     * When downloading from phone, temporarily make the PDF page
     * exactly the same width as the PC version.
     */
    page.style.width = "794px";
    page.style.minWidth = "794px";

    // Give browser time to apply the desktop-size layout
    await new Promise((resolve) => setTimeout(resolve, 200));

    // Load the footer image once and place it at the bottom of every A4 page.
    const footerDataUrl = await imageElementToDataUrl(
      document.getElementById("pdfFooter"),
    );

    const canvas = await html2canvas(page, {
      scale: 3,
      useCORS: true,
      backgroundColor: "#ffffff",
      logging: false,
      width: 794,
      windowWidth: 794,
    });

    const { jsPDF } = window.jspdf;

    const pdf = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
      compress: true,
    });

    const pageWidth = 210;
    const pageHeight = 297;

    const imgWidth = pageWidth;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;

    const toleranceMm = 1;

    // Keep the footer at the same A4 position as the reference PDF.
    // It is drawn separately so every page gets the footer, including
    // multi-page attendance sheets.
    const footerImage = new Image();
    footerImage.src = footerDataUrl;
    await new Promise((resolve, reject) => {
      footerImage.onload = resolve;
      footerImage.onerror = () =>
        reject(new Error("Could not load the footer image."));
    });

    /*
     * Use the footer size exactly as the live preview renders it, so
     * the exported footer keeps the same height and aspect ratio.
     */
    const footerEl = document.getElementById("pdfFooter");
    const footerRectHeight = footerEl
      ? footerEl.getBoundingClientRect().height
      : 0;

    const footerHeightMm = footerRectHeight
      ? pxToMm(footerRectHeight)
      : imgWidth * (footerImage.naturalHeight / footerImage.naturalWidth);

    const footerY = pageHeight - footerHeightMm;

    const addFooter = () => {
      pdf.addImage(
        footerDataUrl,
        "PNG",
        0,
        footerY,
        imgWidth,
        footerHeightMm,
        undefined,
        "FAST",
      );
    };

    if (imgHeight <= pageHeight + toleranceMm) {
      // One A4 page
      pdf.addImage(
        canvas.toDataURL("image/png"),
        "PNG",
        0,
        0,
        imgWidth,
        imgHeight,
      );
      addFooter();
    } else {
      // Multiple A4 pages
      const pxPerMm = canvas.width / imgWidth;
      const toleranceForLastSlicePx = toleranceMm * pxPerMm;
      const pageHeightPx = Math.round(pageHeight * pxPerMm);

      let renderedPx = 0;
      let first = true;

      while (canvas.height - renderedPx > toleranceForLastSlicePx) {
        const sliceHeightPx = Math.min(
          pageHeightPx,
          canvas.height - renderedPx,
        );

        const sliceCanvas = document.createElement("canvas");

        sliceCanvas.width = canvas.width;
        sliceCanvas.height = sliceHeightPx;

        const ctx = sliceCanvas.getContext("2d");

        ctx.drawImage(
          canvas,
          0,
          renderedPx,
          canvas.width,
          sliceHeightPx,
          0,
          0,
          canvas.width,
          sliceHeightPx,
        );

        const sliceHeightMm = sliceHeightPx / pxPerMm;

        if (!first) {
          pdf.addPage();
        }

        pdf.addImage(
          sliceCanvas.toDataURL("image/png"),
          "PNG",
          0,
          0,
          imgWidth,
          sliceHeightMm,
        );

        addFooter();

        renderedPx += sliceHeightPx;
        first = false;
      }
    }

    const d = dateInput.value || "attendance";
    const safeDate = d.replaceAll("-", "_");

    if (previewMode === "participant") {
      const safeTitle = (titleInput.value.trim() || "Participant_List").replace(
        /[^a-z0-9_-]+/gi,
        "_",
      );
      pdf.save(`IEEE_AIUB_Participant_${safeTitle}.pdf`);
    } else {
      pdf.save(`IEEE_AIUB_Attendance_${safeDate}.pdf`);
    }

    setStatus("PDF downloaded successfully.");
  } catch (error) {
    console.error(error);

    const isTainted = /tainted|SecurityError|insecure/i.test(
      String(error && error.message),
    );

    setStatus(
      isTainted
        ? "PDF generation failed: the browser blocked reading the page image."
        : `PDF generation failed: ${
            error && error.message ? error.message : "unknown error"
          }`,
      "error",
    );
  } finally {
    // Remove PDF mode
    page.classList.remove("pdf-capture");

    // Restore phone/mobile layout
    page.style.width = originalWidth;
    page.style.minWidth = originalMinWidth;

    downloadBtn.disabled = false;
    downloadBtn.textContent = originalText;
  }
});

async function imageElementToDataUrl(imageElement) {
  if (!imageElement) {
    throw new Error("Footer image element was not found.");
  }

  if (!imageElement.complete || !imageElement.naturalWidth) {
    await new Promise((resolve, reject) => {
      imageElement.onload = resolve;
      imageElement.onerror = () =>
        reject(new Error("Could not load the footer image."));
    });
  }

  const canvas = document.createElement("canvas");
  canvas.width = imageElement.naturalWidth;
  canvas.height = imageElement.naturalHeight;

  const ctx = canvas.getContext("2d");
  ctx.drawImage(imageElement, 0, 0);

  try {
    return canvas.toDataURL("image/png");
  } catch (error) {
    throw new Error("Could not read the footer image.");
  }
}

async function imageUrlToUint8Array(url) {
  const response = await fetch(url);

  if (!response.ok) {
    throw new Error("Could not load the header image.");
  }

  const blob = await response.blob();
  const buffer = await blob.arrayBuffer();

  return new Uint8Array(buffer);
}

downloadWordBtn.addEventListener("click", async () => {
  if (!entries.length && previewMode !== "participant") {
    setStatus(
      "Add at least one entry before downloading the Word file.",
      "error",
    );
    return;
  }

  const originalText = downloadWordBtn.textContent;
  downloadWordBtn.disabled = true;
  downloadWordBtn.textContent = "Generating Word...";

  const page = document.getElementById("pdfPage");
  let restoreLayout = null;

  try {
    if (!window.docx) {
      throw new Error(
        "Word library failed to load — check your internet connection and try again.",
      );
    }

    const docxLib = window.docx;
    const { Document, Packer, Paragraph, TextRun, ImageRun, Footer } = docxLib;

    /*
     * Measure the preview exactly as it is drawn for the PDF, so the
     * Word file and the PDF are built from the very same geometry.
     */
    page.classList.add("pdf-capture");
    restoreLayout = lockPreviewToExportWidth(page);

    // Give the browser time to apply the desktop-size layout.
    await new Promise((resolve) => setTimeout(resolve, 200));

    const pageStyle = window.getComputedStyle(page);
    const baseFont = cssFirstFontFamily(pageStyle.fontFamily);
    const baseSize = pxToHalfPoint(cssNumber(pageStyle.fontSize) || 12);
    const baseColor = cssColorToHex(pageStyle.color);

    const headerImgEl = page.querySelector('img[alt="AIUB IEEE Header"]');
    const footerImgEl = document.getElementById("pdfFooter");
    const tableEl = tableWrap.querySelector("table");

    if (!headerImgEl) {
      throw new Error("Header image element was not found.");
    }

    // ----- Header image (spans the full page width, like the preview)
    const headerRect = headerImgEl.getBoundingClientRect();
    const headerBytes = await imageUrlToUint8Array(headerImgEl.src);

    const headerParagraph = new Paragraph({
      spacing: { before: 0, after: 0, line: 240, lineRule: "auto" },
      children: [
        new ImageRun({
          data: headerBytes,
          transformation: {
            width: Math.round(headerRect.width),
            height: Math.round(headerRect.height),
          },
        }),
      ],
    });

    // ----- Text blocks between the header image and the table
    const blockElements = [participantDate, brandTitle, sheetTitle].filter(
      isVisibleBlock,
    );

    let previousBottom = headerRect.bottom;
    const blockGaps = blockElements.map((el) => {
      const rect = el.getBoundingClientRect();
      const before = rect.top - previousBottom;
      previousBottom = rect.bottom;
      return { el, before, after: 0 };
    });

    // The gap between the last text block and the table becomes the
    // "space after" of that block, because Word tables have no
    // spacing-before of their own.
    if (tableEl && blockGaps.length) {
      const tableTop = tableEl.getBoundingClientRect().top;
      blockGaps[blockGaps.length - 1].after = Math.max(
        0,
        tableTop - previousBottom,
      );
    }

    const blockParagraphs = blockGaps.map((entry) =>
      domBlockToDocxParagraph(
        entry.el,
        page,
        { before: entry.before, after: entry.after },
        docxLib,
      ),
    );

    // ----- The attendance / event / participant table itself
    const bodyChildren = [headerParagraph, ...blockParagraphs];

    if (tableEl) {
      bodyChildren.push(domTableToDocxTable(tableEl, page, docxLib));
    }

    // ----- Footer image, repeated at the bottom of every page
    let footerSection;
    let footerHeightTwip = 0;

    if (footerImgEl) {
      const footerRect = footerImgEl.getBoundingClientRect();
      const footerHeightPx = Math.round(footerRect.height);
      const footerWidthPx = Math.round(footerRect.width);
      footerHeightTwip = pxToTwip(footerHeightPx);

      const footerBytes = await imageUrlToUint8Array(footerImgEl.src);

      footerSection = new Footer({
        children: [
          new Paragraph({
            spacing: { before: 0, after: 0, line: 240, lineRule: "auto" },
            children: [
              new ImageRun({
                data: footerBytes,
                transformation: {
                  width: footerWidthPx,
                  height: footerHeightPx,
                },
              }),
            ],
          }),
        ],
      });
    }

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: { font: baseFont, size: baseSize, color: baseColor },
            paragraph: {
              spacing: { before: 0, after: 0, line: 240, lineRule: "auto" },
            },
          },
        },
      },
      sections: [
        {
          properties: {
            page: {
              size: { width: A4_WIDTH_TWIP, height: A4_HEIGHT_TWIP },
              margin: {
                top: 0,
                right: 0,
                bottom: footerHeightTwip,
                left: 0,
                header: 0,
                footer: 0,
                gutter: 0,
              },
            },
          },
          footers: footerSection ? { default: footerSection } : undefined,
          children: bodyChildren,
        },
      ],
    });

    // Every measurement is done: give the preview its own layout back.
    restoreLayout();
    restoreLayout = null;
    page.classList.remove("pdf-capture");

    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;

    if (previewMode === "participant") {
      a.download = `IEEE_AIUB_Participant_${(titleInput.value.trim() || "List").replace(/[^a-z0-9_-]+/gi, "_")}.docx`;
    } else {
      const d = dateInput.value || "attendance";
      a.download = `IEEE_AIUB_Attendance_${d.replaceAll("-", "_")}.docx`;
    }

    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    setStatus("Word document downloaded successfully.");
  } catch (error) {
    console.error(error);
    setStatus(
      `Word generation failed: ${error && error.message ? error.message : "unknown error"}`,
      "error",
    );
  } finally {
    if (restoreLayout) restoreLayout();
    page.classList.remove("pdf-capture");

    downloadWordBtn.disabled = false;
    downloadWordBtn.textContent = originalText;
  }
});

renderEntries();
updateControlsForMode();
renderTable();
