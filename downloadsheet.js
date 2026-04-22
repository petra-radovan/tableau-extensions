"use strict";

console.log("downloadsheet.js loaded");

$(document).ready(function () {
  console.log("document ready");

  const $worksheetSelect = $("#worksheetSelect");
  const $downloadBtn = $("#downloadBtn");
  const $status = $("#status");

  let dashboard = null;
  let worksheets = [];

  const LOGO_PATH = "./awcl logo.png";
  const LOGO_SCALE_PERCENT = 70;

  $status.text("Pokrećem ekstenziju...");

  if (typeof tableau === "undefined") {
    console.error("tableau is undefined");
    $status.text("Greška: Tableau Extensions API nije učitan.");
    return;
  }

  tableau.extensions.initializeAsync({ configure: showConfigure })
    .then(function () {
      console.log("Tableau initialized successfully");

      dashboard = tableau.extensions.dashboardContent.dashboard;
      worksheets = dashboard.worksheets || [];

      if (!worksheets.length) {
        $worksheetSelect.html('<option value="">Nema worksheetova</option>');
        $downloadBtn.prop("disabled", true);
        $status.text("Nema worksheetova u ovom dashboardu.");
        return;
      }

      populateWorksheetDropdown(worksheets);
      $downloadBtn.prop("disabled", false);

      $status.text(
        "Ekstenzija je uspješno pokrenuta.\n" +
        "Dashboard: " + dashboard.name + "\n" +
        "Odaberi worksheet i klikni 'Preuzmi XLSX'."
      );
    })
    .catch(function (err) {
      const errorText = getErrorText(err);
      console.error("Tableau extension init error:", err);
      $status.text("Inicijalizacija nije uspjela: " + errorText);
      $downloadBtn.prop("disabled", true);
    });

  $downloadBtn.on("click", async function () {
    const selectedWorksheetName = $worksheetSelect.val();

    if (!selectedWorksheetName) {
      $status.text("Najprije odaberi worksheet.");
      return;
    }

    const worksheet = findWorksheetByName(selectedWorksheetName);

    if (!worksheet) {
      $status.text("Odabrani worksheet nije pronađen.");
      return;
    }

    try {
      $downloadBtn.prop("disabled", true);
      await downloadWorksheetAsXlsx(worksheet);
    } finally {
      $downloadBtn.prop("disabled", false);
    }
  });

  function showConfigure() {
    console.log("Configure clicked");
    $status.text("Kliknuto je Configure.");
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

  function findWorksheetByName(name) {
    return worksheets.find(function (worksheet) {
      return worksheet.name === name;
    });
  }

  async function downloadWorksheetAsXlsx(worksheet) {
    $status.text("Dohvaćam podatke za worksheet: " + worksheet.name + "...");

    try {
      const dataTable = await worksheet.getSummaryDataAsync();

      if (!dataTable || !dataTable.columns || !dataTable.data) {
        $status.text("Worksheet nema dostupne podatke za export.");
        return;
      }

      $status.text("Generiram XLSX datoteku...");

      const workbook = new ExcelJS.Workbook();
      workbook.creator = "Tableau Extension";
      workbook.created = new Date();

      const safeSheetName = makeSafeWorksheetName(worksheet.name);
      const excelSheet = workbook.addWorksheet(safeSheetName);

      const headers = (dataTable.columns || []).map(function (column, index) {
        return column.fieldName || column.caption || ("Column " + (index + 1));
      });

      excelSheet.addRow(headers);

      const headerRow = excelSheet.getRow(1);
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

      (dataTable.data || []).forEach(function (row) {
        const excelRow = excelSheet.addRow(
          row.map(function (cell, colIndex) {
            return convertCellValue(cell, dataTable.columns[colIndex]);
          })
        );

        excelRow.eachCell(function (cell, colNumber) {
          const columnMeta = dataTable.columns[colNumber - 1];
          applyExcelFormatting(cell, columnMeta);
        });
      });

      autoFitColumns(excelSheet, headers);

      excelSheet.views = [{ state: "frozen", ySplit: 1 }];

      let logoInserted = false;

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
            resolve({
              width: img.width,
              height: img.height
            });
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
          ext: {
            width: finalWidth,
            height: finalHeight
          }
        });

        logoInserted = true;
      } catch (logoError) {
        console.warn("Logo could not be inserted:", logoError);
      }

      const xlsxBuffer = await workbook.xlsx.writeBuffer();

      const fileName = sanitizeFileName(worksheet.name) + ".xlsx";
      triggerXlsxDownload(xlsxBuffer, fileName);

      $status.text(
        "Preuzimanje je pokrenuto.\n" +
        "Worksheet: " + worksheet.name + "\n" +
        "Redaka: " + dataTable.data.length + "\n" +
        "Datoteka: " + fileName + "\n" +
        "Logo umetnut: " + (logoInserted ? "da" : "ne")
      );
    } catch (err) {
      const errorText = getErrorText(err);
      console.error("Error creating XLSX:", err);
      $status.text(
        "Greška pri generiranju XLSX datoteke za worksheet '" +
        worksheet.name +
        "': " +
        errorText
      );
    }
  }

  function convertCellValue(cell) {
    if (!cell) {
      return "";
    }

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

  function applyExcelFormatting(cell, columnMeta) {
    const columnName = (columnMeta && (columnMeta.fieldName || columnMeta.caption) || "")
      .toLowerCase();

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

    if (!text) {
      return null;
    }

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

    if (Number.isFinite(parsed)) {
      return parsed;
    }

    return null;
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
    if (!err) {
      return "Unknown error";
    }

    if (typeof err === "string") {
      return err;
    }

    if (err.message) {
      return err.message;
    }

    try {
      return JSON.stringify(err);
    } catch (jsonError) {
      return "Unserializable error";
    }
  }
});