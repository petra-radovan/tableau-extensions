"use strict";

console.log("downloadsheet.js loaded");

$(document).ready(function () {
  const $worksheetSelect = $("#worksheetSelect");
  const $downloadBtn = $("#downloadBtn");
  const $sendEmailBtn = $("#sendEmailBtn");
  const $emailInput = $("#emailInput");
  const $status = $("#status");

  let dashboard = null;
  let worksheets = [];

  const LOGO_PATH = "./awcl logo.png";
  const LOGO_SCALE_PERCENT = 70;
  const EMAIL_API_URL = "https://YOUR-BACKEND-URL/send-export-email";

  setStatus("Pokrećem ekstenziju...");

  if (typeof $ === "undefined") {
    setStatus("Greška: jQuery nije učitan.");
    return;
  }

  if (typeof ExcelJS === "undefined") {
    setStatus("Greška: ExcelJS nije učitan.");
    return;
  }

  if (typeof tableau === "undefined") {
    setStatus("Greška: Tableau Extensions API nije učitan.");
    return;
  }

  tableau.extensions.initializeAsync()
    .then(function () {
      dashboard = tableau.extensions.dashboardContent.dashboard;
      worksheets = dashboard.worksheets || [];

      if (!worksheets.length) {
        $worksheetSelect.html('<option value="">Nema worksheetova</option>');
        $downloadBtn.prop("disabled", true);
        $sendEmailBtn.prop("disabled", true);
        setStatus("Ekstenzija je inicijalizirana, ali ne vidi worksheetove.");
        return;
      }

      populateWorksheetDropdown(worksheets);
      $downloadBtn.prop("disabled", false);
      $sendEmailBtn.prop("disabled", false);

      setStatus(
        "Ekstenzija je uspješno pokrenuta.\n" +
        "Dashboard: " + safeText(dashboard.name) + "\n" +
        "Odaberi worksheet pa preuzmi ili pošalji mail."
      );
    })
    .catch(function (err) {
      setStatus("Inicijalizacija nije uspjela:\n" + getErrorText(err));
      $downloadBtn.prop("disabled", true);
      $sendEmailBtn.prop("disabled", true);
    });

  $downloadBtn.on("click", async function () {
    const worksheet = getSelectedWorksheet();
    if (!worksheet) return;

    try {
      toggleButtons(true);
      const result = await buildWorkbookBuffer(worksheet);
      triggerXlsxDownload(result.buffer, result.fileName);

      setStatus(
        "Preuzimanje je pokrenuto.\n" +
        "Worksheet: " + worksheet.name + "\n" +
        "Datoteka: " + result.fileName
      );
    } catch (err) {
      setStatus("Greška pri downloadu:\n" + getErrorText(err));
    } finally {
      toggleButtons(false);
    }
  });

  $sendEmailBtn.on("click", async function () {
    const worksheet = getSelectedWorksheet();
    if (!worksheet) return;

    const email = String($emailInput.val() || "").trim();

    if (!email) {
      setStatus("Upiši email adresu.");
      return;
    }

    if (!isValidEmail(email)) {
      setStatus("Email adresa nije valjana.");
      return;
    }

    try {
      toggleButtons(true);
      setStatus("Generiram XLSX i šaljem mail...");

      const result = await buildWorkbookBuffer(worksheet);
      const base64File = arrayBufferToBase64(result.buffer);

      const response = await fetch(EMAIL_API_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          to: email,
          subject: "Tableau export - " + worksheet.name,
          body: "U prilogu je XLSX export za worksheet: " + worksheet.name,
          fileName: result.fileName,
          fileBase64: base64File
        })
      });

      if (!response.ok) {
        const text = await response.text();
        throw new Error("Backend error: " + text);
      }

      setStatus(
        "Mail je uspješno poslan.\n" +
        "Primatelj: " + email + "\n" +
        "Worksheet: " + worksheet.name
      );
    } catch (err) {
      setStatus("Greška pri slanju maila:\n" + getErrorText(err));
    } finally {
      toggleButtons(false);
    }
  });

  function getSelectedWorksheet() {
    const selectedWorksheetName = $worksheetSelect.val();

    if (!selectedWorksheetName) {
      setStatus("Najprije odaberi worksheet.");
      return null;
    }

    const worksheet = worksheets.find(function (w) {
      return w.name === selectedWorksheetName;
    });

    if (!worksheet) {
      setStatus("Odabrani worksheet nije pronađen.");
      return null;
    }

    return worksheet;
  }

  function populateWorksheetDropdown(worksheetList) {
    $worksheetSelect.empty();

    worksheetList.forEach(function (worksheet) {
      const option = $("<option></option>")
        .val(worksheet.name)
        .text(worksheet.name);

      $worksheetSelect.append(option);
    });
  }

  async function buildWorkbookBuffer(worksheet) {
    const dataTable = await worksheet.getSummaryDataAsync();

    if (!dataTable || !dataTable.columns || !dataTable.data) {
      throw new Error("Worksheet nema dostupne podatke za export.");
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Tableau Extension";
    workbook.created = new Date();

    const safeSheetName = makeSafeWorksheetName(worksheet.name);
    const excelSheet = workbook.addWorksheet(safeSheetName);

    const transformed = transformDataForExcel(dataTable);

    excelSheet.addRow(transformed.headers);
    styleHeaderRow(excelSheet.getRow(1));

    transformed.rows.forEach(function (rowObject) {
      const excelRow = excelSheet.addRow(
        transformed.headers.map(function (header) {
          return rowObject[header];
        })
      );

      excelRow.eachCell(function (cell, colNumber) {
        const headerName = transformed.headers[colNumber - 1];
        applyExcelFormatting(cell, headerName);
      });
    });

    autoFitColumns(excelSheet, transformed.headers);
    excelSheet.views = [{ state: "frozen", ySplit: 1 }];

    await insertLogo(workbook, excelSheet);

    const buffer = await workbook.xlsx.writeBuffer();
    const fileName = sanitizeFileName(worksheet.name) + ".xlsx";

    return { buffer, fileName };
  }

  async function insertLogo(workbook, excelSheet) {
    try {
      const logoResponse = await fetch(LOGO_PATH);

      if (!logoResponse.ok) {
        throw new Error("Logo nije pronađen na putanji: " + LOGO_PATH);
      }

      const blob = await logoResponse.blob();

      const img = new Image();
      const objectUrl = URL.createObjectURL(blob);

      const dimensions = await new Promise(function (resolve, reject) {
        img.onload = function () {
          resolve({ width: img.width, height: img.height });
        };
        img.onerror = reject;
        img.src = objectUrl;
      });

      URL.revokeObjectURL(objectUrl);

      const scalePercent = Math.max(1, Math.min(100, LOGO_SCALE_PERCENT));
      const scaleFactor = scalePercent / 100;

      const finalWidth = Math.round(dimensions.width * scaleFactor);
      const finalHeight = Math.round(dimensions.height * scaleFactor);

      const arrayBuffer = await blob.arrayBuffer();

      const imageId = workbook.addImage({
        buffer: arrayBuffer,
        extension: "png"
      });

      excelSheet.addImage(imageId, {
        tl: { col: 6, row: 7 },
        ext: { width: finalWidth, height: finalHeight }
      });
    } catch (err) {
      console.warn("Logo nije umetnut:", err);
    }
  }

  function transformDataForExcel(dataTable) {
    const pivoted = pivotMeasureNamesData(dataTable);

    if (pivoted) {
      return pivoted;
    }

    const headers = (dataTable.columns || []).map(function (column, index) {
      return column.fieldName || column.caption || ("Column " + (index + 1));
    });

    const rows = (dataTable.data || []).map(function (row) {
      const result = {};
      headers.forEach(function (header, colIndex) {
        result[header] = convertCellValue(row[colIndex]);
      });
      return result;
    });

    return { headers, rows };
  }

  function pivotMeasureNamesData(dataTable) {
    const columnNames = (dataTable.columns || []).map(function (column, index) {
      return column.fieldName || column.caption || ("Column " + (index + 1));
    });

    const measureNameIndex = columnNames.findIndex(function (name) {
      return String(name).toLowerCase().includes("measure names");
    });

    const measureValueIndex = columnNames.findIndex(function (name) {
      return String(name).toLowerCase().includes("measure values");
    });

    if (measureNameIndex === -1 || measureValueIndex === -1) {
      return null;
    }

    const dimensionHeaders = columnNames.filter(function (_, index) {
      return index !== measureNameIndex && index !== measureValueIndex;
    });

    const rowMap = new Map();
    const discoveredMeasureHeaders = [];

    (dataTable.data || []).forEach(function (row) {
      const measureNameCell = row[measureNameIndex];
      const measureValueCell = row[measureValueIndex];

      const measureName = measureNameCell && measureNameCell.formattedValue
        ? String(measureNameCell.formattedValue)
        : "Measure";

      if (!discoveredMeasureHeaders.includes(measureName)) {
        discoveredMeasureHeaders.push(measureName);
      }

      const dimensionValues = [];
      const baseObject = {};

      columnNames.forEach(function (header, index) {
        if (index !== measureNameIndex && index !== measureValueIndex) {
          const value = row[index] && row[index].formattedValue !== undefined
            ? row[index].formattedValue
            : row[index] && row[index].value !== undefined
              ? row[index].value
              : "";

          dimensionValues.push(String(value));
          baseObject[header] = value;
        }
      });

      const key = dimensionValues.join("||");

      if (!rowMap.has(key)) {
        rowMap.set(key, baseObject);
      }

      const targetRow = rowMap.get(key);
      targetRow[measureName] = convertCellValue(measureValueCell);
    });

    const headers = dimensionHeaders.concat(discoveredMeasureHeaders);
    const rows = Array.from(rowMap.values());

    return { headers, rows };
  }

  function styleHeaderRow(headerRow) {
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFF00" }
    };
    headerRow.alignment = {
      vertical: "middle",
      horizontal: "center"
    };
    headerRow.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } }
    };
  }

  function convertCellValue(cell) {
    if (!cell) return "";

    const formatted = typeof cell.formattedValue !== "undefined" && cell.formattedValue !== null
      ? String(cell.formattedValue).trim()
      : "";

    const rawValue = typeof cell.value !== "undefined" ? cell.value : null;

    if (formatted.includes("%")) {
      const percentNumber = parseLocaleNumber(formatted.replace("%", ""));
      if (percentNumber !== null) {
        return percentNumber / 100;
      }
    }

    if (typeof rawValue === "number" && Number.isFinite(rawValue)) {
      return rawValue;
    }

    if (formatted) {
      const parsedFormatted = parseLocaleNumber(formatted);
      if (parsedFormatted !== null) {
        return parsedFormatted;
      }
      return formatted;
    }

    if (rawValue === null || rawValue === undefined) {
      return "";
    }

    return rawValue;
  }

  function applyExcelFormatting(cell, headerName) {
    const columnName = String(headerName || "").toLowerCase();

    cell.alignment = { vertical: "middle" };

    if (typeof cell.value === "number") {
      cell.alignment.horizontal = "right";
    }

    if (isPercentageColumn(columnName, cell.value)) {
      cell.numFmt = "0%";
      return;
    }

    if (typeof cell.value === "number") {
      cell.numFmt = "#,##0";
    }
  }

  function isPercentageColumn(columnName, value) {
    if (
      columnName.includes("%") ||
      columnName.includes("percent") ||
      columnName.includes("pct") ||
      columnName.includes("ratio") ||
      columnName.includes("total")
    ) {
      return true;
    }

    if (typeof value === "number" && value >= 0 && value <= 1) {
      return true;
    }

    return false;
  }

  function parseLocaleNumber(input) {
    if (input === null || input === undefined) {
      return null;
    }

    let text = String(input).trim();
    if (!text) return null;

    text = text.replace(/\s/g, "");

    const hasComma = text.includes(",");
    const hasDot = text.includes(".");

    if (hasComma && hasDot) {
      if (text.lastIndexOf(",") > text.lastIndexOf(".")) {
        text = text.replace(/\./g, "");
        text = text.replace(",", ".");
      } else {
        text = text.replace(/,/g, "");
      }
    } else if (hasComma) {
      text = text.replace(",", ".");
    }

    const parsed = Number(text);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function autoFitColumns(excelSheet, headers) {
    excelSheet.columns.forEach(function (column, columnIndex) {
      let maxLength = 10;

      const headerText = headers[columnIndex] ? String(headers[columnIndex]) : "";
      maxLength = Math.max(maxLength, headerText.length);

      column.eachCell({ includeEmpty: true }, function (cell) {
        let cellValue = "";
        if (cell.value !== null && cell.value !== undefined) {
          cellValue = String(cell.value);
        }
        maxLength = Math.max(maxLength, cellValue.length);
      });

      column.width = Math.min(maxLength + 2, 50);
    });
  }

  function triggerXlsxDownload(buffer, fileName) {
    const blob = new Blob(
      [buffer],
      { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
    );

    const blobUrl = URL.createObjectURL(blob);

    const link = document.createElement("a");
    link.href = blobUrl;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    URL.revokeObjectURL(blobUrl);
  }

  function arrayBufferToBase64(buffer) {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    const chunkSize = 0x8000;

    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode.apply(null, chunk);
    }

    return btoa(binary);
  }

  function isValidEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }

  function toggleButtons(disabled) {
    $downloadBtn.prop("disabled", disabled);
    $sendEmailBtn.prop("disabled", disabled);
  }

  function makeSafeWorksheetName(name) {
    let safeName = String(name)
      .replace(/[\\/*?:[\]]/g, "_")
      .trim();

    if (!safeName) {
      safeName = "Sheet1";
    }

    return safeName.substring(0, 31);
  }

  function sanitizeFileName(name) {
    return String(name)
      .replace(/[<>:"/\\|?*]+/g, "_")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function getErrorText(err) {
    if (!err) return "Unknown error";
    if (typeof err === "string") return err;
    if (err.message) return err.message;

    try {
      return JSON.stringify(err);
    } catch (jsonError) {
      return "Unserializable error";
    }
  }

  function setStatus(text) {
    $status.text(text);
  }

  function safeText(value) {
    if (value === null || value === undefined) return "";
    return String(value);
  }
});