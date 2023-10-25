import React from "react";
import Papa from "papaparse";
import "./styles.css";
import tableToExcel from "@linways/table-to-excel";

import TableExporter from "./table-exporter";

var jQuery = require("jquery");
require("pivottable/dist/pivot.js");
require("./pivottable.es.js");
window.jQuery = jQuery;
require("jquery-ui-dist/jquery-ui");
require("pivottable/dist/pivot.min.css");

export default class App extends React.Component {
  componentDidMount() {
    Papa.parse(
      "https://raw.githubusercontent.com/aunitz/assets-for-codepen/master/alumnos-matriculados-fp-dual-por-sexo-nivel-fp-y-titularidad.csv",
      {
        download: true,
        skipEmptyLines: true,
        complete: (parsed) => {
          window.jQuery("#output").pivotUI(
            parsed.data,
            {
              rows: [
                "Territorio",
                "Nivel educativo",
                "Sexo",
                "Titularidad del centro"
              ],
              cols: ["Año/Curso"],
              vals: ["value"],
              hiddenFromDragDrop: ["value"],
              aggregatorName: "Suma de enteros",
              rendererName: "Tabla"
            },
            false,
            "es"
          );
        }
      }
    );
  }

  shouldComponentUpdate() {
    return false;
  }

  handleExportClick = () => {
    var htmlTable = document.querySelector(".pvtTable").cloneNode(true);
    //var htmlTable = jQuery(".pvtTable");

    console.log(htmlTable);

    const htmlTableHead = htmlTable.querySelector("thead");
    const htmlHeadRows = htmlTableHead.querySelectorAll("tr");
    htmlHeadRows.forEach((headRow) => {
      const htmlHeadCells = headRow.querySelectorAll("th");
      htmlHeadCells.forEach((htmlCell) => {
        const isAxisLabel = htmlCell.classList.contains("pvtAxisLabel");
        const isColLabel = htmlCell.classList.contains("pvtColLabel");
        const isTotalLabel = htmlCell.classList.contains("pvtTotalLabel");

        if (isAxisLabel) {
          htmlCell.setAttribute("data-a-h", "left");
          htmlCell.setAttribute("data-a-v", "middle");
        }
        if (isColLabel) {
          htmlCell.setAttribute("data-a-h", "center");
          htmlCell.setAttribute("data-a-v", "middle");
        }
        if (isTotalLabel) {
          htmlCell.setAttribute("data-exclude", "true");
        }
      });
    });

    const htmlTableBody = htmlTable.querySelector("tbody");
    const htmlBodyRows = htmlTableBody.querySelectorAll("tr");
    htmlBodyRows.forEach((bodyRow) => {
      const htmlBodyCells = bodyRow.querySelectorAll("th, td");
      htmlBodyCells.forEach((htmlCell) => {
        const isRowLabel = htmlCell.classList.contains("pvtRowLabel");
        const isValue = htmlCell.classList.contains("pvtVal");
        const isTotal = htmlCell.classList.contains("pvtTotal");
        const isTotalLabel = htmlCell.classList.contains("pvtTotalLabel");
        const isGrandTotal = htmlCell.classList.contains("pvtGrandTotal");

        if (isRowLabel) {
          htmlCell.setAttribute("data-a-h", "left");
          htmlCell.setAttribute("data-a-v", "middle");
        }
        if (isValue) {
          htmlCell.setAttribute("data-a-h", "right");
          htmlCell.setAttribute("data-a-v", "middle");
          htmlCell.setAttribute("data-t", "n");
        }
        if (isTotal || isTotalLabel || isGrandTotal) {
          htmlCell.setAttribute("data-exclude", "true");
        }
      });
    });

    tableToExcel.convert(htmlTable, { name: "mine.xlsx" });
  };

  render() {
    return (
      <div className="App">
        <button onClick={this.handleExportClick}>Export to excel</button>
        <div id="output" />
      </div>
    );
  }
}
