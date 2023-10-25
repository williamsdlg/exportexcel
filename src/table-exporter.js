import ExcelJS from "exceljs";
import saveAs from "file-saver";

export default function exportToExcel(htmlTable) {
  if (!checkIfTableFormatIsCorrect(htmlTable)) {
    console.error("The given table did not match the expected format");
    return;
  }

  console.log(htmlTable);

  const workbook = createWorkbook();
  const worksheet = workbook.addWorksheet("My Sheet");

  const htmlTableHead = htmlTable.querySelector("thead");
  const htmlTableHeadRows = htmlTableHead.querySelectorAll("tr");
  const htmlTableHeadRowNumber = htmlTableHeadRows.length;
  let currentRow = 0;
  let mergedCells = [];

  htmlTableHeadRows.forEach((headRow) => {
    let currentColumn = 0;

    const htmlHeadCells = headRow.querySelectorAll("th");
    htmlHeadCells.forEach((htmlCell) => {
      const rowspan = Number(htmlCell.getAttribute("rowspan") || "1");
      const colspan = Number(htmlCell.getAttribute("colspan") || "1");

      console.log(rowspan);

      if (rowspan > 1 || colspan > 1) {
        const mergedCell = {
          startRow: currentRow,
          startColumn: currentColumn,
          endRow: currentRow + rowspan,
          endColumn: currentColumn + colspan,
          rowspan,
          colspan
        };

        worksheet.mergeCells(
          mergedCell.startColumn,
          mergedCell.startColumn,
          mergedCell.endRow,
          mergedCell.endColumn
        );
        mergedCells.push(mergedCell);
      }

      currentColumn += colspan;
    });

    const headCells = Array.from(htmlHeadCells).map((cell) => cell.innerHTML);

    worksheet.addRow(headCells);
  });

  saveWorkbook(workbook);
}

function checkIfTableFormatIsCorrect(htmlTable) {
  const hasCorrectClass = htmlTable.classList.contains("pvtTable");
  const rowNumber = Number(htmlTable.getAttribute("data-numrows") || "0");
  const columnNumber = Number(htmlTable.getAttribute("data-numcols") || "0");

  return hasCorrectClass && rowNumber > 0 && columnNumber > 0;
}

function createWorkbook() {
  const workbook = new ExcelJS.Workbook();
  workbook.created = new Date();
  workbook.views = [
    {
      x: 0,
      y: 0,
      width: 10000,
      height: 20000,
      firstSheet: 0,
      activeTab: 1,
      visibility: "visible"
    }
  ];

  return workbook;
}

function saveWorkbook(workbook) {
  workbook.xlsx.writeBuffer().then((buffer) => {
    saveAs(
      new Blob([buffer], { type: "application/octet-stream" }),
      "MyWorkbook.xlsx"
    );
  });
}
