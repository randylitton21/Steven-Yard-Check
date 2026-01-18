const ROW_COUNT = 25;

const yardCheckBody = document.getElementById("yardCheckBody");
const dateInput = document.getElementById("date");
const timeInput = document.getElementById("time");
const truckInput = document.getElementById("truck");
const tripInput = document.getElementById("trip");
const locationInput = document.getElementById("location");

const exportWordBtn = document.getElementById("exportWord");
const exportExcelBtn = document.getElementById("exportExcel");
const exportTxtBtn = document.getElementById("exportTxt");

function isMobileDevice() {
  return /Android|iPhone|iPad|iPod|webOS|BlackBerry|IEMobile|Opera Mini/i.test(
    navigator.userAgent
  );
}

function buildRows() {
  const fragment = document.createDocumentFragment();

  for (let i = 1; i <= ROW_COUNT; i += 1) {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td class="trailer-cell">
        <span class="row-number">${i}.</span>
        <input type="text" data-field="trailer" />
      </td>
      <td>
        <select data-field="fuel">
          <option value=""></option>
          <option value="F">F</option>
          <option value="7/8">7/8</option>
          <option value="3/4">3/4</option>
          <option value="5/8">5/8</option>
          <option value="1/2">1/2</option>
          <option value="3/8">3/8</option>
          <option value="1/4">1/4</option>
          <option value="1/8">1/8</option>
          <option value="E">E</option>
        </select>
      </td>
      <td class="combined-cell">
        <select data-field="status" class="combined-select">
          <option value=""></option>
          <option value="Loaded">Loaded</option>
          <option value="Empty">Empty</option>
          <option value="Red Tagged">Red Tagged</option>
        </select>
        <input
          type="text"
          data-field="issues"
          class="combined-issues combined-hidden"
        />
      </td>
      <td><input type="text" data-field="temp" /></td>
    `;
    fragment.appendChild(row);
  }

  yardCheckBody.appendChild(fragment);
}

function getRowsData() {
  return Array.from(yardCheckBody.querySelectorAll("tr")).map((row) => {
    const getValue = (field) =>
      row.querySelector(`[data-field="${field}"]`).value.trim();

    return {
      trailer: getValue("trailer"),
      fuel: getValue("fuel"),
      status: getValue("status"),
      issues: getValue("issues"),
      temp: getValue("temp"),
    };
  });
}

function getMetaData() {
  const date = dateInput.value.trim();
  const time = timeInput.value.trim();
  const dateTime = date || time ? `${date}${date && time ? " / " : ""}${time}` : "";

  return {
    date,
    time,
    dateTime,
    truck: truckInput.value.trim(),
    trip: tripInput.value.trim(),
    location: locationInput.value.trim(),
  };
}

function getStatusColumns(row) {
  if (row.status === "Red Tagged") {
    return {
      loadedEmpty: "",
      redTagged: row.issues || "",
    };
  }

  return {
    loadedEmpty: row.status,
    redTagged: "",
  };
}

async function getImageDataUrlFromElement(element) {
  if (!element) {
    return null;
  }

  const imageElement = await new Promise((resolve, reject) => {
    if (element.complete && element.naturalWidth > 0) {
      resolve(element);
      return;
    }
    element.addEventListener("load", () => resolve(element), { once: true });
    element.addEventListener(
      "error",
      () => reject(new Error("Image failed to load")),
      { once: true }
    );
  });

  try {
    const canvas = document.createElement("canvas");
    canvas.width = imageElement.naturalWidth;
    canvas.height = imageElement.naturalHeight;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(imageElement, 0, 0);
    return canvas.toDataURL("image/png");
  } catch (error) {
    // Fall back to direct file read if canvas is blocked.
  }

  try {
    const response = await fetch(imageElement.src);
    const blob = await response.blob();
    const dataUrl = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
    return dataUrl;
  } catch (error) {
    return null;
  }
}

async function getLogoDataUrl() {
  return getImageDataUrlFromElement(document.querySelector(".logo-image"));
}

function dataUrlToArrayBuffer(dataUrl) {
  if (!dataUrl) {
    return null;
  }
  const base64 = dataUrl.split(",")[1];
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

function getFilenameBase() {
  const now = new Date();
  const stamp = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(
    2,
    "0"
  )}-${String(now.getDate()).padStart(2, "0")}`;
  return `perdue_team_yard_check_${stamp}`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function shareFileOrDownload(blob, filename, mimeType) {
  const file = new File([blob], filename, { type: mimeType });

  if (navigator.share) {
    try {
      await navigator.share({
        files: [file],
        title: filename,
      });
      return;
    } catch (error) {
      // Continue to fallback below.
    }
  }

  // iOS Safari often ignores download; open a new tab so the user can save.
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.target = "_blank";
  link.rel = "noopener";
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}

async function exportToWord() {
  const {
    AlignmentType,
    Document,
    ImageRun,
    Packer,
    Paragraph,
    Table,
    TableRow,
    TableCell,
    TextRun,
    WidthType,
  } = window.docx;
  const meta = getMetaData();
  const rows = getRowsData();
  let logoData = null;
  try {
    const logoDataUrl = await getLogoDataUrl();
    logoData = dataUrlToArrayBuffer(logoDataUrl);
  } catch (error) {
    logoData = null;
  }

  const dateLine = meta.date || "____";
  const timeLine = meta.time || "____";
  const truckLine = meta.truck || "________";
  const tripLine = meta.trip || "________";
  const locationLine = meta.location || "________";

  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 100, type: WidthType.PERCENTAGE },
            children: [
              ...(logoData
                ? [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [
                        new ImageRun({
                          data: new Uint8Array(logoData),
                          transformation: { width: 300, height: 86 },
                        }),
                      ],
                    }),
                  ]
                : []),
            ],
            borders: {
              top: { size: 0, color: "FFFFFF" },
              bottom: { size: 0, color: "FFFFFF" },
              left: { size: 0, color: "FFFFFF" },
              right: { size: 0, color: "FFFFFF" },
            },
          }),
        ],
      }),
    ],
  });

  const headerInfoLine = new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: "Date/Time: ", bold: true }),
      new TextRun({ text: dateLine }),
      new TextRun({ text: " / " }),
      new TextRun({ text: timeLine }),
      new TextRun({ text: "    Truck: ", bold: true }),
      new TextRun({ text: truckLine }),
      new TextRun({ text: "    Trip: ", bold: true }),
      new TextRun({ text: tripLine }),
      new TextRun({ text: "    Location: ", bold: true }),
      new TextRun({ text: locationLine }),
    ],
  });
  const titleLine = new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({
        text: "Perdue Team Yard Check",
        bold: true,
        size: 24,
      }),
    ],
  });

  const tableRows = [
    new TableRow({
      height: { value: 420, rule: "exact" },
      children: [
        { text: "Trailer", width: 14 },
        { text: "Fuel", width: 6 },
        { text: "Loaded/Empty", width: 16 },
        {
          text: 'If "Red Tagged" Record issues here and report to R/R',
          width: 56,
        },
        { text: "Temp", width: 8 },
      ].map((header) =>
        new TableCell({
          width: { size: header.width, type: WidthType.PERCENTAGE },
          children: [
            new Paragraph({
              children: [new TextRun({ text: header.text, bold: true, size: 22 })],
            }),
          ],
        })
      ),
    }),
    ...rows.map(
      (row, index) => {
        const statusColumns = getStatusColumns(row);

        return new TableRow({
          height: { value: 420, rule: "exact" },
          children: [
            { text: `${index + 1}. ${row.trailer}`.trim(), width: 14 },
            { text: row.fuel, width: 6 },
            { text: statusColumns.loadedEmpty, width: 16 },
            { text: statusColumns.redTagged, width: 56 },
            { text: row.temp, width: 8 },
          ].map(
            (cell) =>
              new TableCell({
                width: { size: cell.width, type: WidthType.PERCENTAGE },
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: cell.text || "", size: 22 })],
                  }),
                ],
              })
          ),
        });
      }
    ),
  ];

  const table = new Table({
    width: {
      size: 100,
      type: WidthType.PERCENTAGE,
    },
    rows: tableRows,
  });

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Times New Roman",
            size: 22,
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 360, right: 360, bottom: 360, left: 360 },
          },
        },
        children: [headerTable, headerInfoLine, titleLine, table],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  downloadBlob(blob, `${getFilenameBase()}.docx`);
}

function exportToExcel() {
  exportToExcelWithLogo();
}

async function exportToExcelWithLogo() {
  const meta = getMetaData();
  const rows = getRowsData();
  const dateLine = meta.date || "____";
  const timeLine = meta.time || "____";
  const truckLine = meta.truck || "________";
  const tripLine = meta.trip || "________";
  const locationLine = meta.location || "________";
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Yard Check", {
    pageSetup: {
      paperSize: 1,
      orientation: "portrait",
      margins: { left: 0.25, right: 0.25, top: 0.25, bottom: 0.25 },
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 1,
    },
  });

  worksheet.properties.defaultRowHeight = 30;

  worksheet.columns = [
    { key: "trailer", width: 26 },
    { key: "fuel", width: 10 },
    { key: "loaded", width: 22 },
    { key: "issues", width: 68 },
    { key: "temp", width: 14 },
  ];

  const infoRow = worksheet.addRow([
    `Date/Time: ${dateLine} / ${timeLine}    Truck: ${truckLine}    Trip: ${tripLine}    Location: ${locationLine}`,
  ]);
  worksheet.mergeCells(`A${infoRow.number}:E${infoRow.number}`);
  infoRow.alignment = { horizontal: "left" };
  infoRow.font = { bold: true, size: 13 };
  infoRow.height = 22;

  const titleRow = worksheet.addRow(["Perdue Team Yard Check"]);
  worksheet.mergeCells(`A${titleRow.number}:E${titleRow.number}`);
  titleRow.alignment = { horizontal: "center" };
  titleRow.font = { bold: true, size: 13 };
  titleRow.height = 22;

  const headerRow = worksheet.addRow([
    "Trailer",
    "Fuel",
    "Loaded/Empty",
    'If "Red Tagged" Record issues here and report to R/R',
    "Temp",
  ]);
  headerRow.font = { bold: true, size: 12 };
  headerRow.alignment = { horizontal: "center" };
  headerRow.height = 22;

  rows.forEach((row, index) => {
    const statusColumns = getStatusColumns(row);
    worksheet.addRow([
      `${index + 1}. ${row.trailer}`.trim(),
      row.fuel,
      statusColumns.loadedEmpty,
      statusColumns.redTagged,
      row.temp,
    ]);
  });

  const tableStartRow = headerRow.number;
  const tableEndRow = tableStartRow + ROW_COUNT;
  for (let rowIndex = tableStartRow; rowIndex <= tableEndRow; rowIndex += 1) {
    const row = worksheet.getRow(rowIndex);
    row.height = 28;
    for (let colIndex = 1; colIndex <= 5; colIndex += 1) {
      const cell = row.getCell(colIndex);
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      if (rowIndex > tableStartRow && (rowIndex - tableStartRow) % 2 === 1) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFF2F2F2" },
        };
      }
    }
  }

  worksheet.pageSetup.printArea = `A1:E${tableEndRow}`;

  try {
    const logoDataUrl = await getLogoDataUrl();
    const logoBuffer = dataUrlToArrayBuffer(logoDataUrl);
    if (logoBuffer) {
      const imageId = workbook.addImage({
        buffer: logoBuffer,
        extension: "png",
      });
      worksheet.addImage(imageId, {
        tl: { col: 0.1, row: 0.1 },
        ext: { width: 220, height: 64 },
      });
    }
  } catch (error) {
    // Ignore logo issues for Excel export.
  }

  const arrayBuffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([arrayBuffer], {
    type:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8",
  });
  await shareFileOrDownload(
    blob,
    `${getFilenameBase()}.xlsx`,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
}

function exportToTxt() {
  const meta = getMetaData();
  const rows = getRowsData();

  const lines = [
    "Stevens Transport",
    "Driver Focused. People Driven.",
    "Perdue Team Yard Check",
    `Date/Time: ${meta.dateTime} | Truck: ${meta.truck} | Trip: ${meta.trip} | Location: ${meta.location}`,
    "",
    [
      "Trailer",
      "Fuel",
      "Loaded/Empty",
      'If "Red Tagged" Record issues here and report to R/R',
      "Temp",
    ].join("\t"),
    ...rows.map((row, index) =>
      (() => {
        const statusColumns = getStatusColumns(row);
        return [
          `${index + 1}. ${row.trailer}`.trim(),
          row.fuel,
          statusColumns.loadedEmpty,
          statusColumns.redTagged,
          row.temp,
        ].join("\t");
      })()
    ),
  ];

  const blob = new Blob([lines.join("\n")], { type: "text/plain" });
  downloadBlob(blob, `${getFilenameBase()}.txt`);
}

function clearForm() {
  dateInput.value = "";
  timeInput.value = "";
  truckInput.value = "";
  tripInput.value = "";
  locationInput.value = "";

  yardCheckBody.querySelectorAll("input, textarea, select").forEach((input) => {
    input.value = "";
  });
}

buildRows();

yardCheckBody.addEventListener("change", (event) => {
  const target = event.target;
  if (!target.matches("select[data-field='status']")) {
    return;
  }

  const row = target.closest("tr");
  const issuesInput = row.querySelector("input[data-field='issues']");

  if (target.value === "Red Tagged") {
    issuesInput.classList.remove("combined-hidden");
    target.classList.add("combined-hidden");
    issuesInput.value = "";
    issuesInput.focus();
  } else {
    issuesInput.classList.add("combined-hidden");
    issuesInput.value = "";
    target.classList.remove("combined-hidden");
  }
});

yardCheckBody.addEventListener(
  "blur",
  (event) => {
    const target = event.target;
    if (!target.matches("input[data-field='issues']")) {
      return;
    }

    if (target.value.trim() !== "") {
      return;
    }

    const row = target.closest("tr");
    const statusSelect = row.querySelector("select[data-field='status']");
    target.classList.add("combined-hidden");
    statusSelect.classList.remove("combined-hidden");
    statusSelect.value = "";
  },
  true
);

exportWordBtn.addEventListener("click", () => {
  exportToWord();
});

exportExcelBtn.addEventListener("click", () => {
  exportToExcel();
});

exportTxtBtn.addEventListener("click", () => {
  exportToTxt();
});

