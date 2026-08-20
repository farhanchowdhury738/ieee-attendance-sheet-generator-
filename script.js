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
const entryList = document.getElementById("entryList");
const countBadge = document.getElementById("countBadge");
const tableWrap = document.getElementById("tableWrap");
const sheetTitle = document.getElementById("sheetTitle");
const brandTitle = document.getElementById("brandTitle");
const previewTitle = document.getElementById("previewTitle");
const statusBox = document.getElementById("status");

const DEFAULT_BRAND_TITLE = "IEEE AIUB Student Branch";

// Preview mode: "attendance" (default) or "event".
let previewMode = "attendance";

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

function addEntry() {
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

  if (previewMode === "event") {
    eventBtnLabel.textContent = "Attendance";
    previewTitle.textContent = "Live Event Preview";
    brandTitle.textContent = titleInput.value.trim() || DEFAULT_BRAND_TITLE;
  } else {
    eventBtnLabel.textContent = "Event";
    previewTitle.textContent = "Live Attendance Preview";
    brandTitle.textContent = DEFAULT_BRAND_TITLE;
  }

  renderTable();
});

// Keep the preview title in sync live while typing, but only while
// the Event view is showing.
titleInput.addEventListener("input", () => {
  if (previewMode === "event") {
    brandTitle.textContent = titleInput.value.trim() || DEFAULT_BRAND_TITLE;
  }
});

downloadBtn.addEventListener("click", async () => {
  if (!entries.length) {
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

    const footerHeightMm = Math.min(22, Math.max(12,
      imgWidth * (footerImage.naturalHeight / footerImage.naturalWidth)
    ));
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
        canvas.toDataURL("image/jpeg", 0.95),
        "JPEG",
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
          sliceCanvas.toDataURL("image/jpeg", 0.95),
          "JPEG",
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

    pdf.save(`IEEE_AIUB_Attendance_${safeDate}.pdf`);

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
  if (!entries.length) {
    setStatus(
      "Add at least one entry before downloading the Word file.",
      "error",
    );
    return;
  }

  const originalText = downloadWordBtn.textContent;
  downloadWordBtn.disabled = true;
  downloadWordBtn.textContent = "Generating Word...";

  try {
    if (!window.docx) {
      throw new Error(
        "Word library failed to load — check your internet connection and try again.",
      );
    }
    const {
      Document,
      Packer,
      Paragraph,
      TextRun,
      Table,
      TableRow,
      TableCell,
      ImageRun,
      AlignmentType,
      WidthType,
      HeadingLevel,
      VerticalAlign,
    } = window.docx;

    // Reuse the same header image already embedded on the page (base64).
    const headerImgEl = document.querySelector(
      '#pdfPage img[alt="AIUB IEEE Header"]',
    );
    const headerBytes = await imageUrlToUint8Array(headerImgEl.src);
    const headerNaturalWidth = headerImgEl.naturalWidth || 1274;
    const headerNaturalHeight = headerImgEl.naturalHeight || 294;
    const headerDisplayWidth = 600;
    const headerDisplayHeight = Math.round(
      headerDisplayWidth * (headerNaturalHeight / headerNaturalWidth),
    );

    const headerParagraph = new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new ImageRun({
          data: headerBytes,
          transformation: {
            width: headerDisplayWidth,
            height: headerDisplayHeight,
          },
        }),
      ],
    });

    const brandTitleParagraph = new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 200 },
      children: [
        new TextRun({
          text: brandTitle.textContent.trim() || DEFAULT_BRAND_TITLE,
          bold: true,
          size: 24,
        }),
      ],
    });

    const sheetTitleParagraph = new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [
        new TextRun({
          text: sheetTitle.textContent,
          bold: true,
          size: 22,
        }),
      ],
    });

    const colWidths = {
      slot: 15,
      name: 23,
      reporting: 17,
      checkin: 15,
      checkout: 15,
      signature: 15,
    };

    function headerCell(text, width) {
      return new TableCell({
        width: { size: width, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            children: [new TextRun({ text, bold: true })],
          }),
        ],
      });
    }

    function bodyCell(text, width, opts = {}) {
      return new TableCell({
        width: { size: width, type: WidthType.PERCENTAGE },
        verticalAlign: VerticalAlign.CENTER,
        rowSpan: opts.rowSpan,
        children: [new Paragraph({ text: text || "" })],
      });
    }

    const headerRow = new TableRow({
      tableHeader: true,
      children: [
        headerCell("Slot Time", colWidths.slot),
        headerCell("Name", colWidths.name),
        headerCell("Reporting Time", colWidths.reporting),
        headerCell("Check In Time", colWidths.checkin),
        headerCell("Check Out Time", colWidths.checkout),
        headerCell("Signature", colWidths.signature),
      ],
    });

    // Group entries by slot, same as the on-screen table (rowSpan merge).
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

    const bodyRows = [];
    groups.forEach((group) => {
      group.people.forEach((person, idx) => {
        const cells = [];
        if (idx === 0) {
          cells.push(
            bodyCell(group.slot, colWidths.slot, {
              rowSpan: group.people.length,
            }),
          );
        }
        cells.push(
          bodyCell(person.name, colWidths.name),
          bodyCell(person.reporting, colWidths.reporting),
          bodyCell("", colWidths.checkin),
          bodyCell("", colWidths.checkout),
          bodyCell("", colWidths.signature),
        );
        bodyRows.push(new TableRow({ children: cells }));
      });
    });

    const table = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [headerRow, ...bodyRows],
    });

    const doc = new Document({
      sections: [
        {
          children: [headerParagraph, brandTitleParagraph, sheetTitleParagraph, table],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);

    const d = dateInput.value || "attendance";
    const safeDate = d.replaceAll("-", "_");
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `IEEE_AIUB_Attendance_${safeDate}.docx`;
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
    downloadWordBtn.disabled = false;
    downloadWordBtn.textContent = originalText;
  }
});

renderEntries();
renderTable();
