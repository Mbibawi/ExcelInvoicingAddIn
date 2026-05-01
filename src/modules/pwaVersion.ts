import * as m from "./main.js";
import {
  showUI,
  LawFirmUI,
  MarianneUI,
  IUserInterface,
  byID,
  populateSelectElement,
  splitter,
} from "./ui.js";

export const settingsNames = {
  invoices: {
    workBook: "invoicesWorkbook",
    tableName: "invoicesTable",
    wordTemplate: "invoicesTemplate",
    saveTo: "invoicesSaveTo",
  },
  letter: {
    workBook: "letterWorkbook",
    wordTemplate: "letterTemplate",
    saveTo: "letterSaveTo",
    tableName: "",
  },
  leases: {
    workBook: "leasesWorkbook",
    tableName: "leasesTable",
    wordTemplate: "leasesTemplate",
    saveTo: "leasesSaveTo",
  },
  Marianne: {
    workBook: "reportsWorkbook",
    tableName: "reportsTable",
    wordTemplate: "reportsTemplate",
    saveTo: "reportsSaveTo",
  },
};

interface IGraphAPI {
  addNewEntry(): Promise<void>;
  getContentControlsValues(...args: any[]): [string, string][];
  getWordTableRows(...args: any[]): {
    wordRows: string[][];
    totalsLabels?: string[];
  };
  updateTableRow(...args: any[]): Promise<void>;
}

interface ILawFirm {
  issueInvoice(...args: any[]): Promise<void>;
  issueLetter(...args: any[]): Promise<void>;
  issueLeaseLetter(...args: any[]): Promise<void>;
  searchFiles(...args: any[]): Promise<void>;
}

export class LawFirm implements ILawFirm {
  private UI: IUserInterface;
  private stored;
  private form;
  private tenantID;
  private settingsNames;

  constructor() {
    this.form = byID();
    this.UI = new LawFirmUI(this);
    this.stored = saveSettings(this.UI, undefined, true) || undefined;
    this.settingsNames = settingsNames;
    this.tenantID = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
  }

  getUI = () => this.UI;

  /**
   *
   * @param {boolean} add - If false, the function will only show a form containing input fields for the user to provide the data for the new row to be added to the Excel Table. If true, the function will parse the values from the input fields in the form, and will add them as a new row to the Excel Table. Its default value is false.
   * @param {boolean} display - If provided, the function will show the visible rows in the UI after the new row has been added.
   * @param {any[]} row - If provided, the function will add the row directly to the Excel Table without needing to retrieve the data from the inputs.
   */
  async addNewEntry(row?: any[]) {
    m.spinner(true); //We show the spinner
    const form = this.form ?? byID() ?? undefined;

    const UI = this.UI,
      inputOnChange = this.inputOnChange;

    const { workbookPath, tableName } = this.getConsts(
      this.settingsNames.invoices,
    );

    if ([this.stored, workbookPath, tableName].find((v) => !v))
      m.throwAndAlert("One of the constant values is not valid");

    const graph = new m.GraphAPI("", workbookPath);

    const TableRows = await graph.fetchExcelTable(tableName, true);
    if (!TableRows?.length) return alert("Failed to retrieve the Excel table");
    const tableTitles = TableRows[0];
    if (row) return await addEntry(TableRows, tableTitles, row);

    await showAddNewForm();

    async function showAddNewForm() {
      try {
        await createForm();
        m.spinner(false); //We hide the sinner
      } catch (error) {
        m.spinner(false); //We hide the sinner
        alert(error);
      }

      async function createForm() {
        const tableBody = TableRows!.slice(1, -1);
        const inputs: HTMLInputElement[] = [];
        const bound = (indexes: number[]) =>
          inputs
            .filter((input) => indexes.includes(m.getIndex(input)))
            .map((input) => [input, m.getIndex(input)]) as InputCol[];

        insertAddForm(tableTitles);

        function insertAddForm(titles: string[]) {
          if (!titles)
            throw new Error(
              "The table titles are missing. Check the console.log for more details",
            );

          if (!form) throw new Error("Could not find the form element");
          form.innerHTML = "";

          const divs = titles.map((title, index) => {
            const div = newDiv(index);
            if (![4, 7].includes(index))
              div.appendChild(createLable(title, index)); //We exclued the labels for "Total Time" and for "Year"
            div.appendChild(createInput(index));
            return div;
          });

          (function groupDivs() {
            [
              [11, 12, 13], //"Moyen de Paiement", "Compte", "Tiers"
              [9, 10], //"Montant", "TVA"
              [5, 6, 8], //"Start Time", "End Time", "Taux Horaire"
            ].forEach((group) =>
              newDiv(
                NaN,
                divs.filter((div) => group.includes(Number(div.dataset.block))),
              ),
            );
          })();

          (function addBtn() {
            const btnIssue = document.createElement("button");
            btnIssue.innerText = "Add Entry";
            btnIssue.classList.add("button");
            btnIssue.onclick = () => addEntry(TableRows!, tableTitles); //!We omit the row argument in order for addEntry() to parse the values from the inputs
            form.appendChild(btnIssue);
          })();

          (function homeBtn() {
            showUI(UI, true);
          })();

          function newDiv(
            i: number,
            divs?: HTMLDivElement[],
            css: string = "block",
          ) {
            if (divs) return groupDivs();
            else return create();

            function create() {
              const div = document.createElement("div");
              div.dataset.block = i.toString();
              form!.appendChild(div);
              div.classList.add(css);
              return div;
            }

            function groupDivs() {
              const div = newDiv(i, undefined, "group") as HTMLDivElement;
              divs?.forEach((el) => div.appendChild(el));
              form!.children[3]?.insertAdjacentElement("afterend", div);
              return div;
            }
          }

          function createLable(title: string, i: number) {
            const label = document.createElement("label");
            label.htmlFor = "input" + i.toString();
            label.innerHTML = title + ":";
            return label;
          }

          function createInput(index: number) {
            const css = "field";
            const input = document.createElement("input");
            const id = "input" + index.toString();

            (function append() {
              input.classList.add(css);
              input.id = id;
              input.name = id;
              input.autocomplete = "on";
              input.dataset.index = index.toString();
              input.type = "text";
              inputs.push(input);
            })();

            (function customize() {
              if ([8, 9, 10].includes(index)) input.type = "number";
              else if (index === 3) input.type = "date";
              else if ([5, 6].includes(index)) input.type = "time";
              else if ([4, 7].includes(index)) input.style.display = "none"; //We hide those 2 columns: 'Total Time' and the 'Year'

              (function addDataLists() {
                const updateNext = [0, 1, 8, 15]; //Those are the indexes of the inputs (i.e; the columns numbers) that need to get an onChange event in order to update the dataLists of the next inputs when the current input is changed: "Client"(0), "Affaire"(1), "Taux Horaire"(8), "Adresses"(15)

                if (updateNext.includes(index))
                  input.onchange = () =>
                    inputOnChange(index, bound(updateNext), tableBody, false);

                if (![0, 2, 11, 12, 13].includes(index)) return; //We will initially populate the "Client"(0), Nature(2), "Payment Method"(11), "Bank Account"(12), "Third Party"(13) lists only, the other inputs will be populate when the onChange function will be called
                populateSelectElement(
                  input,
                  m.getUniqueValues(index, tableBody),
                );
              })();

              (function addRestOnChange() {
                if (index < 5 || index > 10) return;
                //Only for the  "Start Time", "End Time", "Total Time", "Hourly Rate", "Amount", "VAT" columns . The "Total Time" input (7) is hidden, so it can't be changed by the user. We will add the onChange event to it by simplicity
                const reset = [5, 6, 7, 9, 10]; //Those are "Start Time" (5), "End Time" (6), "Total Time" (7, although it is hidden), "Amount" (9), "VAT" (10) columns. We exclude the "Hourly Rate" column (8). We let the user rest it if he wants
                input.onchange = () => resetInputs(bound(reset), index); //!We are passing the table[][] argument as undefined, and the invoice argument as false which means that the function will only reset the bound inputs without updating any data list
              })();
            })();

            return input;
          }

          function resetInputs(
            inputs: [HTMLInputElement, number][],
            index: number,
          ) {
            inputs
              .filter(([input, index]) => index > index)
              .forEach(([input, index]) => reset(input)); //We reset any input which dataset-index is > than the dataset-index of the input that has been changed

            if (index === 9)
              inputs
                .filter(([input, index]) => index < index)
                .forEach(([input], index) => reset(input)); //If the input is the input for the "Montant" column of the Excel table, we also reset the "Start Time" (5), "End Time" (6) and "Hourly Rate" (7) columns' inputs. We do this because we assume that if the user provided the amount, it means that either this is not a fee, or the fee is not hourly billed.

            function reset(input: HTMLInputElement) {
              if (!input) return;
              input.value = "";
              if (input.valueAsNumber) input.valueAsNumber = 0;
            }
          }
        }
      }
    }

    async function addEntry(
      tableRows: any[][],
      tableTitles: string[],
      row?: any[],
    ) {
      if (!row?.length) row = parseInputs() ?? undefined;
      try {
        const visibleCells = await addRow(row);
        if (visibleCells?.length) showFilteredRows(visibleCells);
        m.spinner(false); //We hide the spinner
      } catch (error) {
        m.spinner(false); //We hide the spinner
        alert(error);
      }

      function parseInputs() {
        const colNature = 2,
          colDate = 3,
          colStart = 5,
          colEnd = 6,
          colRate = 8,
          colAmount = 9,
          colVAT = 10;
        const stop = (missing: string) =>
          alert(
            `${missing} missing. You must at least provide the client, matter, nature, date and the amount. If you provided a time start, you must provide the end time and the hourly rate. Please review your iputs`,
          );

        const inputs = Array.from(
          document.getElementsByTagName("input"),
        ) as HTMLInputElement[]; //all inputs

        const nature = m.getInputByIndex(inputs, colNature)?.value;
        if (!nature) return stop("The matter is");
        const date = m.getInputByIndex(inputs, colDate)
          ?.valueAsDate as Date | null;
        if (!date) return stop("The invoice date is");
        const amount = m.getInputByIndex(inputs, colAmount) as HTMLInputElement;
        const rate = m.getInputByIndex(inputs, colRate)?.valueAsNumber || 0;

        const debit = [
          "Honoraire",
          "Débours/Dépens",
          "Débours/Dépens non facturables",
          "Rétrocession d'honoraires",
          "Charges déductibles",
        ].includes(nature); //We check if we need to change the value sign

        const row = inputs.map((input, index) => getInputValue(index)); //!CAUTION: The html inputs are not arranged according to their dataset.index values. If we follow their order, some values will be assigned to the wrong column of the Excel table. That's why we do not pass the input itself or the dataset.index of the input to getInputValue(), but instead we pass the index of the column for which we want to retrieve the value from the relevant input.

        if (missing()) return stop("Some of the required fields are");

        return row;

        function getInputValue(index: number) {
          const input = m.getInputByIndex(inputs, index) as HTMLInputElement;
          if ([colDate, colDate + 1].includes(index))
            return m.getISODate(date); //Those are the 2 date columns
          else if ([colStart, colEnd].includes(index))
            return m.getTime([input]); //time start and time end columns
          else if (index === 7) {
            //!This is a hidden input
            const timeInputs = [colStart, colEnd].map((i) =>
              m.getInputByIndex(inputs, i),
            );
            const totalTime = m.getTime(timeInputs); //Total time column
            if (totalTime && rate && !amount.valueAsNumber)
              amount.valueAsNumber = totalTime * 24 * rate; // making the amount equal the rate * totalTime
            return totalTime;
          } else if (debit && index === colAmount)
            return -input.valueAsNumber || 0; //This is the amount if negative
          else if ([colRate, colAmount, colVAT].includes(index))
            return input.valueAsNumber || 0; //Hourly Rate, Amount, VAT
          else return input.value;
        }

        function missing() {
          if (
            row.filter(
              (value, i) => (i < colDate + 1 || i === colAmount) && !value,
            ).length > 0
          )
            return true; //if client name, matter, nature, date or amount are missing

          if (row[colStart] === row[colEnd])
            return false; //If the total time = 0 we do not need to alert if the hourly rate is missing
          else if (row[colStart] && (!row[colEnd] || !row[colRate]))
            return true; //if startTime is provided but without endTime or without hourly rate
          else if (row[colEnd] && (!row[colStart] || !row[colRate]))
            return true; //if endTime is provided but without startTime or without hourly rate
        }
      }

      async function addRow(row: any[] | undefined) {
        if (!row) return m.throwAndAlert("The row is not valid");
        const visibleCells = await graph.addRowToExcelTable(
          row,
          tableRows!.length - 2,
          tableName!,
          tableTitles,
        );

        if (!visibleCells?.length)
          return m.throwAndAlert(
            "There was an issue with the adding or the filtering, check the console.log for more details",
          );

        alert("Row aded and the table was filtered");
        return visibleCells;
      }
      function showFilteredRows(visibleCells: string[][]) {
        const tableDiv = createDivContainer();
        if (!form) return m.throwAndAlert("The form element was not found");
        const table = document.createElement("table");
        table.classList.add("table");
        tableDiv.appendChild(table);

        const columns = [0, 1, 2, 3, 7, 8, 9, 10, 14]; //The columns that will be displayed in the table;
        const rowClass = "excelRow";
        (function insertTableHeader() {
          if (!tableTitles) return m.throwAndAlert("No Table Titles");
          const headerRow = document.createElement("tr");
          headerRow.classList.add(rowClass);
          const thead = document.createElement("thead");
          table.appendChild(thead);
          thead.appendChild(headerRow);
          tableTitles.forEach((cell, index) => {
            if (!columns.includes(index)) return;
            addTableCell(headerRow, cell, "th");
          });
        })();
        (function insertTableRows() {
          const tbody = document.createElement("tbody");
          table.appendChild(tbody);
          visibleCells.forEach((row, index) => {
            if (index < 1) return; //We exclude the header row
            if (!row) return;
            const tr = document.createElement("tr");
            tr.classList.add(rowClass);
            tbody.appendChild(tr);
            row.forEach((cell, index) => {
              if (!columns.includes(index)) return;
              addTableCell(tr, cell, "td");
            });
          });
        })();

        form.insertAdjacentElement("afterend", tableDiv);

        function createDivContainer() {
          const id = "retrieved";
          let tableDiv = byID(id);
          if (tableDiv) {
            tableDiv.innerHTML = "";
            return tableDiv;
          }
          tableDiv = document.createElement("div");
          tableDiv.classList.add("table-div");
          tableDiv.id = id;
          return tableDiv;
        }

        function addTableCell(parent: HTMLElement, text: string, tag: string) {
          const cell = document.createElement(tag);
          //   cell.classList.add(css);
          cell.textContent = text;
          parent.appendChild(cell);
        }
      }
    }
  }

  async issueInvoice() {
    const form = this.form;
    const { workbookPath, tableName, templatePath, saveTo } = this.getConsts(
      this.settingsNames.invoices,
    );
    const UI = this.UI,
      inputOnChange = this.inputOnChange,
      stored = this.stored;

    const addNewEntry = async (row: any[]) => await this.addNewEntry(row); //!We had to redefined it with arrow function  because addNewEntry() uses "this", if called from any sub function of issueInvoice(), this will be changed or undefined.
    await showInvoiceForm();

    async function showInvoiceForm() {
      m.spinner(true); //We show the spinner

      if (
        [stored, workbookPath, tableName, templatePath, saveTo].find((v) => !v)
      )
        m.throwAndAlert("One of the  constant values is not valid");

      const graph = new m.GraphAPI("", workbookPath);

      const sessionId = (await graph.createFileSession()) || "";
      if (!sessionId)
        return m.throwAndAlert(
          "There was an issue with the creation of the file cession. Check the console.log for more details",
        );

      const tableRows = await graph.fetchExcelTable(tableName, true);
      if (!tableRows?.length)
        return m.throwAndAlert("Failed to retrieve the Excel table");
      const tableTitles = tableRows[0];
      document.querySelector("table")?.remove();
      try {
        insertInvoiceForm(tableTitles);
        await graph.closeFileSession(sessionId);
        m.spinner(false); //We hide the spinner
      } catch (error) {
        m.throwAndAlert(`Error while showing the invoice user form: ${error}`);
        m.spinner(false); //We hide the spinner
      }

      function insertInvoiceForm(tableTitles: string[]) {
        if (!form) throw new Error("The form element was not found");
        const isNan = (index: number | string) => isNaN(Number(index));
        form.innerHTML = "";
        const tableBody = tableRows!.slice(1, -1);
        const boundInputs: InputCol[] = [];

        (function insertInputs() {
          insertInputsAndLables([0, 1, 2, 3, 3], "input"); //Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice
          insertInputsAndLables(["Discount"], "discount")[0].value = "0%"; //Inserting a discount percentage input and setting its default value to 0%
          insertInputsAndLables(["Français", "English"], "lang", true); //Inserting languages checkboxes
        })();

        (function customizeDateLabels() {
          const [from, to] = Array.from(
            document.getElementsByTagName("label"),
          )?.filter((label) => label.htmlFor.endsWith("3"));
          if (from) from.innerText += " From (included)";
          if (to) to.innerText += " To/Before (included)";
        })();

        (function addIssueInvoiceBtn() {
          const btnIssue = document.createElement("button");
          btnIssue.innerText = "Generate Invoice";
          btnIssue.classList.add("button");
          btnIssue.onclick = () =>
            createInvoice(tableName, tableTitles, templatePath!, saveTo, graph);
          form.appendChild(btnIssue);
        })();

        (function homeBtns() {
          showUI(UI, true);
        })();

        function insertInputsAndLables(
          indexes: (number | string)[],
          id: string,
          checkBox: boolean = false,
        ): HTMLInputElement[] {
          let css = "field";
          if (checkBox) css = "checkBox";
          return indexes.map((index) => {
            const div = newDiv(String(index));
            appendLable(index, div);
            return appendInput(index, div);
          });

          function appendInput(index: number | string, div: HTMLDivElement) {
            const input = document.createElement("input");
            input.classList.add(css);
            const isNaN = isNan(index);
            !isNaN ? (input.id = id + index.toString()) : (input.id = id);

            (function inputType() {
              if (checkBox) input.type = "checkbox";
              else if (isNaN || (index as number) < 3) input.type = "text";
              else input.type = "date";
            })();

            (function notCheckBox() {
              if (isNaN || checkBox) return; //If the index is not a number or the input is a checkBox, we return;
              index = Number(index);
              input.name = input.id;
              input.dataset.index = index.toString();
              if (index < 3) boundInputs.push([input, index]); //Fields "Client"(0), "Affaire"(1), "Nature"(2) are the inputs that will need to get their dataList created or updated each time the previous input is changed.
              if (index < 2)
                input.onchange = () =>
                  inputOnChange(index as number, boundInputs, tableBody, true); //We add onChange on "Client" (0) and "Affaire" (1) columns. We set combined = true in order to add to the dataList of the next column an option combining all the choices in the list
              if (index < 1)
                populateSelectElement(input, m.getUniqueValues(0, tableBody)); //We create a unique values dataList for the "Client" (0) input
            })();

            (function isCheckBox() {
              if (!checkBox) return;
              input.dataset.language = index
                .toString()
                .slice(0, 2)
                .toUpperCase();
              input.onchange = () =>
                Array.from(document.getElementsByTagName("input"))
                  .filter(
                    (checkBox: HTMLInputElement) =>
                      checkBox.dataset.language && checkBox !== input,
                  )
                  .forEach((checkBox) => (checkBox.checked = false));
            })();
            div.appendChild(input);
            return input;
          }

          function appendLable(index: number | string, div: HTMLDivElement) {
            const label = document.createElement("label");
            isNan(Number(index)) || checkBox
              ? (label.innerText = index.toString())
              : (label.innerText = tableTitles[Number(index)]);
            !isNan(Number(index))
              ? (label.htmlFor = id + index.toString())
              : (label.htmlFor = id);
            div?.appendChild(label);
          }

          function newDiv(i: string, css: string = "block") {
            const div = document.createElement("div");
            div.dataset.block = i;
            form!.appendChild(div);
            div.classList.add(css);
            return div;
          }
        }
      }
    }

    async function createInvoice(
      tableName: string,
      tableTitles: string[],
      templatePath: string,
      saveTo: string,
      graph: m.GraphAPI,
    ) {
      m.spinner(true); //We show the spinner
      try {
        await editInvoice(tableName, tableTitles, templatePath, saveTo, graph);
        m.spinner(false); //We hide the spinner
      } catch (error) {
        m.spinner(false); //We hide the sinner
        alert(error);
      }
    }

    async function editInvoice(
      tableName: string,
      tableTitles: string[],
      templatePath: string,
      saveTo: string,
      graph: m.GraphAPI,
    ) {
      const client = tableTitles[0],
        matter = tableTitles[1]; //Those are the 'Client' and 'Matter' columns of the Excel table
      const sessionId = (await graph.createFileSession(true)) || ""; //!persist must be = true. This means that if the session is closed, the changes made to the file will be saved.
      if (!sessionId)
        return m.throwAndAlert(
          "There was an issue with the creation of the file cession. Check the console.log for more details",
        );

      const inputs = Array.from(document.getElementsByTagName("input"));
      const criteria = inputs.filter((input) => m.getIndex(input) >= 0);

      const discount = parseInt(
        inputs.find((input) => input.id === "discount")?.value || "0%",
      );

      const lang =
        inputs.find((input) => input.dataset.language && input.checked === true)
          ?.dataset.language || "FR";

      const date = new Date(); //We need to generate the date at this level and pass it down to all the functions that need it
      const invoiceNumber = getInvoiceNumber(date);
      const data = await filterExcelData(
        criteria,
        discount,
        lang,
        invoiceNumber,
      );

      if (!data)
        return m.throwAndAlert("Could not retrieve the filtered Excel table");

      const { wordRows, totalsLabels, clientName, matters, adresses } = data;

      const invoice = {
        number: invoiceNumber,
        clientName: clientName,
        matters: matters,
        adress: adresses,
        lang: lang,
      };

      const contentControls = getContentControlsValues(invoice, date);

      const fileName = getInvoiceFileName(clientName, matters, invoiceNumber);
      let saveToPath = `${saveTo}/${fileName}`;

      saveToPath =
        prompt(
          `The file will be saved in ${saveTo}, and will be named : ${fileName}.\nIf you want to change the path or the name, provide the full file path and name of your choice without any sepcial characters`,
          saveToPath,
        ) || saveTo;

      (async function editInvoiceFilterExcelClose() {
        await graph.createAndUploadDocumentFromTemplate(
          templatePath!,
          saveToPath,
          lang,
          [["Invoice", wordRows, 1]],
          { nestedCtrls: contentControls },
          totalsLabels,
        );
        await graph.clearFilterExcelTable(tableName!, sessionId); //We unfilter the table;
        await graph.filterExcelTable(
          tableName,
          client,
          [clientName],
          sessionId,
        ); //We filter the table by the matters that were invoiced
        await graph.filterExcelTable(tableName, matter, matters, sessionId); //We filter the table by the matters that were invoiced
        await graph.closeFileSession(sessionId);
      })();

      /**
       * Filters the Excel table according to the values of each inputs, then returns the values of the Word table rows that will be added to the Word table in the invoice template document
       * @param {HTMLInputElement[]} inputs - the html inputs containing the values based on which the table will be filtered
       * @param {number} discount  - The discount percentage that will be applied to the amount of each invoiced row if any. It is a number between 0 and 100. If it is equal to 0, it means that no discount will be applied.
       * @param {string} lang - The language in which the invoice will be issued
       * @returns {Promise<[string[][], string[], string[], string[]]>} - The values of the rows that will be added to the Word table in the invoice template
       */
      async function filterExcelData(
        inputs: HTMLInputElement[],
        discount: number,
        lang: string,
        invoiceNumber: string,
      ): Promise<{
        wordRows: string[][];
        totalsLabels: string[];
        clientName: string;
        matters: string[];
        adresses: string[];
      } | void> {
        const clientCol = 0,
          matterCol = 1,
          dateCol = 3,
          addressCol = 15; //Indexes of the 'Matter' and 'Date' columns in the Excel table
        const clientNameInput = m.getInputByIndex(inputs, clientCol);
        const matterInput = m.getInputByIndex(inputs, matterCol);
        const clientName = clientNameInput!.value || "";
        const matters = m.getArray(matterInput!.value) || []; //!The Matter input may include multiple entries separated by ', ' not only one entry.

        if (!clientName || !matters?.length)
          m.throwAndAlert(
            "could not retrieve the client name or the matter/matters list from the inputs",
          );

        const excelTable = await graph.fetchExcelTable(tableName, true);

        let tableRows = excelTable?.slice(1, -1) || undefined; //We exclude the first and the last rows of the table. Since we are calling the "range" endpoint, we get the whole table including the headers. The first row is the header, and the last row is the total row.

        if (!tableRows)
          return m.throwAndAlert(
            "We could not retrieve the tableRows whie trying to issue the invoice",
          );

        //tableRows = _filterTableByInputsValues([[clientNameInput!, clientCol], [matterInput!, matterCol]], excelTable!);
        tableRows = filterTableByInputsValues(
          [
            [clientNameInput!, clientCol],
            [matterInput!, matterCol],
          ],
          excelTable!,
        );
        tableRows = filterByDate(tableRows!, dateCol);

        const adresses = m.getUniqueValues(addressCol, tableRows) as string[]; //!We must retrieve the adresses at this stage before filtering by "Matter" or any other column

        //const {wordRows, totalsLabels} = _getRowsData(tableRows, discount, lang, invoiceNumber);
        const { wordRows, totalsLabels } = await getWordTableRows(
          tableRows,
          discount,
          lang,
          invoiceNumber,
        );

        return { wordRows, totalsLabels, clientName, matters, adresses };

        function filterByDate(visible: string[][], dateCol: number) {
          const convert = (date: string | number) =>
            dateFromExcel(Number(date)).getTime();

          const [from, to] = inputs
            .filter((input) => m.getIndex(input) === dateCol)
            .map((input) => input.valueAsDate?.getTime());

          if (from && to)
            return visible.filter(
              (row) =>
                convert(row[dateCol]) >= from && convert(row[dateCol]) <= to,
            );
          //we filter by the date
          else if (from)
            return visible.filter((row) => convert(row[dateCol]) >= from); //we filter by the date
          else if (to)
            return visible.filter((row) => convert(row[dateCol]) <= to); //we filter by the date
          else
            return visible.filter(
              (row) => convert(row[dateCol]) <= new Date().getTime(),
            ); //we filter by the date
        }
      }
    }

    /**
     * Returns the Word file name by which the newly issued invoice will be saved on OneDrive
     * @param {string} clientName - The name of the client for which the invoice will be issued
     * @param {string} matters - The matters included in the invoice
     * @param {string} invoiceNumber - The invoice serial number
     * @returns {string} - The name of the Word file to be saved
     */
    function getInvoiceFileName(
      clientName: string,
      matters: string[],
      invoiceNumber: string,
    ): string {
      // return 'test file name for now.docx'
      return `${clientName}_Facture_${Array.from(matters).join("&")}_No.${invoiceNumber.replace("/", "@")}.docx`
        .replaceAll("/", "_")
        .replaceAll('"', "")
        .replaceAll("\\", "");
    }

    function getInvoiceNumber(date: Date): string {
      const padStart = (n: number) => n.toString().padStart(2, "0");

      return `${date.getFullYear() - 2000}${padStart(date.getMonth() + 1)}${padStart(date.getDate())}/${padStart(date.getHours())}${padStart(date.getMinutes())}`;
    }

    /**
     * Returns a string[][] representing the rows to be inserted in the Word table containing the invoice details
     * @param {string[][]} tableRows - The filtered Excel rows from which the data will be extracted and put in the required format
     * @param {string} lang - The language in which the invoice will issued
     * @returns {string[][]} - the rows to be added to the table. Each row has 4 elements
     */
    async function getWordTableRows(
      tableRows: any[][],
      discount: number = 0,
      lang: string,
      invoiceNumber: string,
    ): Promise<{ wordRows: string[][]; totalsLabels: string[] }> {
      const labels: { [index: string]: lable } = {
        totalFees: {
          nature: ["Honoraire"],
          FR: "Total honoraires",
          EN: "Total Fees",
        },
        totalExpenses: {
          nature: [
            "Débours/Dépens",
            "Rétrocession d'honoraires",
            "Débours/Dépens - Ackad Law Office",
            "Charges déductibles",
          ],
          FR: "Total débours et frais",
          EN: "Total Expenses",
        },
        totalPayments: {
          nature: ["Provision/Règlement"],
          FR: "Total provisions reçues",
          EN: "Total Downpayments",
        },
        totalTimeSpent: {
          nature: [],
          FR: "Total des heures facturables (hors prestations facturées au forfait) ",
          EN: "Total billable hours (other than lump-sum billed services)",
        },
        totalDue: {
          nature: [],
          FR: "Montant dû",
          EN: "Total Due",
        },
        totalReinbursement: {
          nature: [],
          FR: "A rembourser",
          EN: "Reimbursement",
        },
        totalDeduction: {
          nature: ["Remise"],
          FR: "Total des remises sur honoraires",
          EN: "Total fees' discounts",
        },
        netFees: {
          nature: [],
          FR: "Total honoraires après réduction",
          EN: "Total fee after discount",
        },
        discountDescription: {
          //This value is not used
          nature: [],
          FR: `XXX% de remise sur les honoraires`,
          EN: `XXX% discount on accrued fees`,
        },
        hourlyBilled: {
          nature: [],
          FR: "facturation au temps passé\u00A0:",
          EN: "hourly billed:",
        },
        hourlyRate: {
          nature: [],
          FR: "au taux horaire de\u00A0:",
          EN: "at an hourly rate of:",
        },
        decimal: {
          nature: [],
          FR: ",",
          EN: ".",
        },
        bankHoler: {
          nature: [],
          FR: "Titulaire du compte",
          EN: "Account holder",
        },
        bankName: {
          nature: [],
          FR: "Banque",
          EN: "Bank",
        },
        bankAdress: {
          nature: [],
          FR: "Adresse",
          EN: "Adress",
        },
      };
      const colDate = 3,
        colAmount = 9,
        colVAT = 10,
        colHours = 7,
        colRate = 8,
        colNature = 2,
        colDescr = 14; //Indexes of the Excel table columns from which we extract the date

      const totalsLabels: string[] = [];
      const wordRows: string[][] = tableRows.map((row) => {
        const date = dateFromExcel(Number(row[colDate]));
        const time = getTimeSpent(Number(row[colHours]));

        let description = `${String(row[colNature])} : ${String(row[colDescr])}`; //Column Nature + Column Description;

        //If the billable hours are > 0, we add to the description: time spent and hourly rate
        if (time)
          description += ` (${labels.hourlyBilled[lang as keyof lable]} ${time}, ${labels.hourlyRate[lang as keyof lable]} ${Math.abs(row[colRate]).toString()}\u00A0€).`;

        const rowValues: string[] = [
          m.getDateString(date), //Column Date
          description,
          getAmountString(row[colAmount] * -1), //Column "Amount": we inverse the +/- sign for all the values
          getAmountString(Math.abs(row[colVAT])), //Column VAT: always a positive value
        ];
        return rowValues;
      });
      await pushTotalsRows();
      return { wordRows, totalsLabels };

      async function pushTotalsRows() {
        //Adding rows for the totals of the different categories and amounts
        const total = (lable: lable) =>
          [colAmount, colVAT].map((col) =>
            sumColumn(col, lable.nature),
          ) as values; //!It always returns the absolute values of the total amount and the total VAT
        const amount = (v: values) => v[0];
        const totalFees = total(labels.totalFees);
        const feesDiscount = totalFees.map(
          (amount) => amount * (discount / 100),
        ); //This is an additional discount applied when the invoice is issued. The Excel table may already include other discounts registered as "Remise"
        const feesDeductions = total(labels.totalDeduction).map(
          (amount, index) => (amount += feesDiscount[index]),
        ) as values; //This is the total of the deductions from the fees: the "Remise" deductions, and the additional discount added at the time the invoice is issued
        const netFees = totalFees.map(
          (amount, index) => amount - feesDeductions[index],
        ) as values;
        const totalPayments = total(labels.totalPayments);
        const totalExpenses = total(labels.totalExpenses);
        const totalTimeSpent: values = [sumColumn(colHours), NaN]; //by omitting to pass the "natures" argument to sumColumn, we do not filter the "Total Time" column by any crieteria. We will get the sum of all the column. since the VAT = NaN, the VAT cell will end up empty.
        const totalDue = netFees.map(
          (amount, index) =>
            amount + totalExpenses[index] - totalPayments[index],
        ) as values;
        const percentage = (amount(feesDeductions) / amount(totalFees)) * 100;

        ["EN", "FR"].forEach(
          (lang) =>
          ((labels.totalDeduction[lang as keyof lable] as string) +=
            ` (${percentage}%)`),
        );

        (function pushTotalsRows() {
          pushRow(labels.totalFees, totalFees);
          pushRow(
            labels.totalDeduction,
            feesDeductions,
            !amount(feesDeductions),
          );
          pushRow(
            labels.netFees,
            netFees,
            !(amount(netFees) < amount(totalFees)),
          ); //We don't push this row if the there is no deduction applied on the fees or if the deduction is = 0
          pushRow(
            labels.totalTimeSpent,
            totalTimeSpent,
            !amount(totalTimeSpent),
          );
          pushRow(labels.totalExpenses, totalExpenses, !amount(totalExpenses));
          pushRow(labels.totalPayments, totalPayments, !amount(totalPayments));
          amount(totalDue) < 0
            ? pushRow(labels.totalReinbursement, totalDue)
            : pushRow(labels.totalDue, totalDue);
        })();
        await addDiscountRowToExcel();

        async function addDiscountRowToExcel() {
          if (!discount) return;
          const newRow = tableRows.find((row) =>
            labels.totalFees.nature.includes(row[colNature]),
          );
          if (!newRow) return;
          const [amount, vat] = feesDiscount; //!The discount must be added as a positive number. This is like a payment made by the client
          const descr = 
            prompt(
              "Provide a description for the discount",
              `Remise sur les honoraires de la facture n° ${invoiceNumber}`,
            ) || "";
          const date = m.getISODate(new Date());
          const cells: [number, string | number][] = [
            [colNature, "Remise"],
            [colAmount, amount],
            [colVAT, vat],
            [colDescr, descr],
            [colDate, date],
            [colDate + 1, date],
          ];

          cells.forEach(([col, value]) => (newRow[col] = value));
          await addNewEntry(newRow!);
        }

        function pushRow(
          rowLable: lable,
          [amount, vat]: values,
          ignore: boolean = false,
        ) {
          if (ignore || isNaN(amount)) return;
          const lable = (rowLable?.[lang as keyof lable] as string) || "";
          if (lable) totalsLabels.push(lable);
          const value =
            rowLable === labels.totalTimeSpent
              ? getTimeSpent(amount)
              : getAmountString(amount);
          wordRows.push([
            lable,
            "",
            value,
            getAmountString(vat), //VAT is always a positive value
          ]);
        }

        /**
         *
         * @param {number} col - the index of the column to be summed
         * @param {string[] | null} natures - the natures of the rows to be included in the sum. If null, we include all the rows regardless of their nature
         * @returns
         */
        function sumColumn(col: number, natures: string[] = []): number {
          let rows = tableRows;
          if (natures.length)
            rows = tableRows.filter((row) => natures.includes(row[colNature])); //If natures is specified, we filter the rows to include only the ones whose nature is included in the natures array
          return Math.abs(sumArray(rows.map((row) => Number(row[col])))); //!We return the absolute value of the total
        }
      }

      function sumArray(values: number[]) {
        let sum = 0;
        values.forEach((value) => (sum += value));
        return sum;
      }

      function getAmountString(value: number): string {
        if (isNaN(value)) return "";

        const amount = value.toLocaleString(
          `${lang.toLowerCase()}-${lang.toUpperCase()}`,
          {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          },
        );

        const versions = {
          FR: `${amount}\u00A0€`,
          EN: `€\u00A0${amount}`,
        };

        return versions[lang as keyof typeof versions];
      }

      /**
       * Convert the time as retrieved from an Excel cell into 'hh:mm' format
       * @param {number} time - The time as stored in an Excel cell
       * @returns {string} - The time as 'hh:mm' format
       */
      function getTimeSpent(time: number): string {
        if (!time || time <= 0) return "";
        time = time * (60 * 60 * 24); //84600 is the number in seconds per day. Excel stores the time as fraction number of days like "1.5" which is = 36 hours 0 minutes 0 seconds;
        const minutes = Math.floor(time / 60);
        const hours = Math.floor(minutes / 60);
        return [hours, minutes % 60, 0]
          .map((el) => el.toString().padStart(2, "0"))
          .join(":");
      }
    }

    function getContentControlsValues(
      arg: {
        number: string;
        clientName: string;
        matters: string[];
        adress: string[];
        lang: string;
      },
      date: Date,
    ): [string, string][] {
      const fields: InvoiceCtrls = {
        dateLabel: {
          title: "LabelParisLe",
          value: { FR: "Paris le ", EN: "Paris on " }[arg.lang] || "",
        },
        date: {
          title: "RTInvoiceDate",
          value: m.getDateString(date),
        },
        numberLabel: {
          title: "LabelInvoiceNumber",
          value:
            { FR: "Facture n°\u00A0:", EN: "Invoice No.:" }[arg.lang] || "",
        },
        number: {
          title: "RTInvoiceNumber",
          value: arg.number,
        },
        subjectLable: {
          title: "LabelSubject",
          value: { FR: "Affaires\u00A0: ", EN: "Matters: " }[arg.lang] || "",
        },
        subject: {
          title: "RTMatter",
          value: arg.matters.join(" & "),
        },
        fee: {
          title: "LabelTableHeadingHonoraire",
          value:
            { FR: "Honoraire/Débours", EN: "Fees/Expenses" }[arg.lang] || "",
        },
        amount: {
          title: "LabelTableHeadingMontantTTC",
          value:
            { FR: "Montant TTC", EN: "Amount VAT Included" }[arg.lang] || "",
        },
        vat: {
          title: "LabelTableHeadingTVA",
          value: { FR: "TVA", EN: "VAT" }[arg.lang] || "",
        },
        disclaimer: {
          title:
            "LabelDisclamer" +
            ["French", "English"].find(
              (el) => !el.toUpperCase().startsWith(arg.lang),
            ) || "English",
          value: "DELETECONTENTECONTROL", //!by setting text = "DELETECONTENTECONTROL", the contentControl will be deleted
        },
        clientName: {
          title: "RTClient",
          value: arg.clientName,
        },
        adress: {
          title: "RTClientAdresse",
          value: arg.adress.join(" & "),
        },
      };
      return Object.values(fields).map((RT) => [RT.title, RT.value as string]);
    }
  }

  async issueLetter() {
    const UI = this.UI,
      { templatePath, saveTo } = this.getConsts(this.settingsNames.letter);
    const form = this.form ?? byID() ?? undefined;
    if (!form) return;
    showForm();

    function showForm() {
      m.spinner(true); //We show the spinner
      document.querySelector("table")?.remove();
      form!.innerHTML = "";
      const input = document.createElement("textarea");
      (function inputAttributes() {
        input.id = "textInput";
        input.classList.add("field");
        form!.appendChild(input);
      })();

      (function generateBtn() {
        const btn = document.createElement("button");
        form!.appendChild(btn);
        btn.classList.add("button");
        btn.innerText = "Créer lettre";
        btn.onclick = () => generate();
      })();

      (function homeBtn() {
        showUI(UI, true);
        m.spinner(false); //We hide the spinner
      })();
    }

    async function generate() {
      try {
        await createLetter();
        m.spinner(false); //We hide the spinner
      } catch (error) {
        console.log(`There was an error: ${error}`);
        m.spinner(false); //We hide the spinner
      }

      async function createLetter() {
        m.spinner(true);
        const input = byID("textInput") as HTMLTextAreaElement;
        if (!input) return;

        const fileName = prompt(
          "Provide the file name without special characthers",
        );
        if (!fileName) return;
        const saveToPath = `${prompt("Provide the destination folder", saveTo || "NO SAVE TO PATH PROVIDED")}/${fileName}.docx`;

        if (!saveToPath) return;
        const contentControls: [string, string][] = [
          ["RTCoreText", input.value],
          ["RTReference", "Référence"],
          ["RTClientName", "Nom du Client"],
          ["RTEmail", "Email du client"],
        ];

        await new m.GraphAPI(
          "",
          saveToPath,
        ).createAndUploadDocumentFromTemplate(
          templatePath,
          saveToPath,
          "FR",
          undefined,
          { nestedCtrls: contentControls },
        );
      }
    }
  }

  async issueLeaseLetter() {
    const form = this.form ?? byID() ?? undefined;
    if (!form) return;
    const UI = this.UI,
      inputOnChange = this.inputOnChange;
    m.spinner(true); //We show the spinner
    const { workbookPath, tableName, templatePath, saveTo } = this.getConsts(
      this.settingsNames.leases,
    );

    if (
      [this.stored, workbookPath, tableName, templatePath, saveTo].find(
        (v) => !v,
      )
    )
      m.throwAndAlert("One of the  constant values is not valid");

    const graph = new m.GraphAPI("", workbookPath);
    const tableRows = await graph.fetchExcelTable(tableName, false); //We are calling the "/rows" endPoint, so we will get the tableBody without the headers

    const Ctrls: LeaseCtrls = {
      owner: {
        title: "RTBailleur",
        col: 0,
        label: "Nom du Bailleur",
        type: "select",
        value: "",
      },
      adress: {
        title: "RTAdresseDestinataire",
        label: "Adresse du bien loué",
        col: 1,
        type: "select",
        value: "",
      },
      tenant: {
        title: "RTLocataire",
        label: "Nom du Locataire",
        col: 2,
        type: "select",
        value: "",
      },
      leaseDate: {
        title: "RTDateBail",
        label: "Date du Bail",
        col: 3,
        type: "date",
        value: "",
      },
      leaseType: {
        title: "RTNature",
        label: "Nature du Bail",
        col: 4,
        type: "text",
        value: "",
      },
      initialIndex: {
        title: "RTIndiceInitial",
        label: "Indice initial",
        col: 5,
        type: "number",
        value: "",
      },
      indexQuarter: {
        title: "RTTrimestre",
        label: "Trimestre de l'indice",
        col: 6,
        type: "text",
        value: "",
      },
      initialIndexDate: {
        title: "RTIndiceInitialDate",
        label: "Date de l'indice initial",
        col: 7,
        type: "date",
        value: "",
      },
      baseIndex: {
        title: "RTIndiceBase",
        label: "Indice de référence",
        col: 8,
        type: "number",
        value: "",
      },
      baseIndexDate: {
        title: "RTDateIndiceBase",
        label: "Date de l'indice de référence",
        col: 9,
        type: "date",
        value: "",
      },
      index: {
        title: "RTIndice",
        label: "Indice de révision",
        col: 10,
        type: "number",
        value: "",
      },
      indexDate: {
        title: "RTDateIndice",
        label: "Date de l'indice de révision",
        col: 11,
        type: "date",
        value: "",
      },
      currentLease: {
        title: "RTLoyerActuel",
        label: "Loyer Actuel (ou révisé)",
        col: 12,
        type: "number",
        value: "",
      },
      revisionDate: {
        title: "RTDateRévision",
        label: "Date de la dernière Révision",
        col: 13,
        type: "date",
        value: "",
      },
      anniversaryDate: { title: "RTDateAnniversaire", value: "" },
      initialYear: { title: "RTIndiceInitialAnnée", value: "" },
      baseYear: { title: "RTIndiceBaseAnnée", value: "" },
      revisionYear: { title: "RTIndiceAnnée", value: "" },
      newLease: { title: "RTLoyerNouveau", value: "" },
      nextRevision: { title: "RTProchaineRevision", value: "" },
      startingMonth: { title: "RTMoisRévision", value: "" },
    };

    const ctrls = Object.values(Ctrls);

    const findRT = (id: string) => ctrls.find((RT) => RT.title === id);
    const fraction = (n: number) => Math.round(n * 100) / 100;

    let row: any[] | undefined,
      rowIndex: number = NaN;
    await showForm();

    async function showForm() {
      const inputs: InputCol[] = [];
      const findInput = (RT: RT) =>
        inputs.find(([input, col]) => input.id === RT.title)?.[0];
      if (!tableRows) return;
      document.querySelector("table")?.remove();
      form!.innerHTML = "";
      const divs: HTMLDivElement[] = [];

      (function insertInputs() {
        const unvalid = (values: (string | undefined)[]) =>
          values.find((value) => !value || isNaN(Number(value)));
        ctrls
          .filter((RT) => !isNaN(RT.col!))
          .map((RT) => inputs.push([createInput(RT), RT.col!] as const));

        const owner = findInput(Ctrls.owner);
        if (owner)
          populateSelectElement(
            owner,
            m.getUniqueValues(Ctrls.owner.col!, tableRows),
            false,
          );

        (function inputsOnChange() {
          const filled = inputs.filter(
            ([input, col]) => col <= Ctrls.tenant.col!,
          );
          filled.forEach(
            ([input, col]) =>
            (input.onchange = () =>
            ([row, rowIndex] = inputOnChange(
              col,
              inputs,
              tableRows,
              false,
            ) || [undefined, NaN])),
          );

          const index = findInput(Ctrls.index);
          const currentLeaseInput = findInput(Ctrls.currentLease);

          index!.onchange = () => {
            if (!row)
              return alert(
                "No single lease having owner name, property adress and tenant name as in the inputs was found",
              );
            const initial = row[Ctrls.initialIndex.col!]; //This is the value of the inital index
            const base = row[Ctrls.index.col!] || initial; //!For the base index, we will retrieve the value of the "Indice de Révision" (column 10) from the Excel row. We will not retrieve this value from the input but from the row itself. If this is the first time we are indexing the lease, we will fall back to the intial index (i.e., the value indicated in the lease agreement)
            const latestIndex = index!.valueAsNumber; //this is the latest index as provided by the user when the input.onChange() event was fired
            const currentLease = row[Ctrls.currentLease.col!]; //This is the value of the current lease
            if (unvalid([base, latestIndex, currentLease]))
              return alert(
                "Please make sure that the values of the current lease, the base indice and the new indice are all provided and valid numbers",
              );

            Ctrls.currentLease.value = currentLease; //!We immediately set the value of this control at this stage, because we will escape this Ctrl when we will update Ctrls values from the inputs, because the corresponding input will be showing the new lease value not the original value

            (function newLease() {
              const newLease = fraction(currentLease * (latestIndex / base)); //we get a 2 digits fractions from the value
              currentLeaseInput!.valueAsNumber = newLease; //!We only update the input value, NOT the value of the Excel row (row). We need to keep the initial value in case the user wants to correct  the value in the index input which means we will need to recalculate the newLease value based on the current lease value. We will hence keep the current lease value unchanged until the generate() function is called.
              Ctrls.newLease.value = newLease; //We update the new lease RT
              Ctrls.baseIndex.value = latestIndex; //We update  the value of the base index with the latest index
            })();
          };
        })();
      })();

      (function groupDivs() {
        [
          [0, 1, 2], //"Bailleur"(0), "Adresse"(1), "Locataire"(2)
          [3, 4], //"Date du Bail"(3), "Nature du Bail"(4)
          [5, 6, 7], //"Indice Initial"(5), "Date de l'indice initial"(6), "Trimestre de l'indice"(7)
          [8, 9], //"Indice de référence"(8), "Date de l'indice de référence"(9)
          [10, 11], //"Indice de révision"(10), "Date de l'indice de révision"(11)
          [12, 13], //"Loyer Actuel (ou révisé)"(12), "Date de la dernière Révision"(13)
        ].forEach((group, index) =>
          groupDivs(
            divs.filter((div) => group.includes(m.getIndex(div))),
            index,
          ),
        );

        function groupDivs(divs: HTMLDivElement[], i: number) {
          const div = document.createElement("div");
          div.classList.add("group");
          div.dataset.block = i.toString();
          divs?.forEach((el) => div.appendChild(el));
          form!.appendChild(div);
          return div;
        }
      })();

      (function generateBtn() {
        const btn = document.createElement("button");
        form!.appendChild(btn);
        btn.classList.add("button");
        btn.innerText = "Créer lettre";
        btn.onclick = () => generate(inputs, row);
      })();

      (function homeBtn() {
        showUI(UI, true);
        m.spinner(false); //We hide the spinner
      })();

      function createInput(RT: RT, className: string = "field") {
        const id = RT.title;
        const div = document.createElement("div");
        form!.appendChild(div);
        const append = (el: HTMLElement) => div.appendChild(el);

        (function appendLabel() {
          if (!RT.label) return;
          const label = document.createElement("label");
          label.htmlFor = id;
          label.innerText = RT.label;
          append(label);
        })();
        return appendInput();
        function appendInput() {
          const input = document.createElement("input") as HTMLInputElement;
          input.type = RT.type || "text";
          input.id = id;
          input.classList.add(className);
          const col = RT.col!.toString();
          input.dataset.index = col;
          div.dataset.index = col;
          append(input);
          divs.push(div);
          return input as HTMLInputElement;
        }
      }
    }
    async function generate(inputs: InputCol[], row: any[] | undefined) {
      if (!inputs.length)
        return m.throwAndAlert("The inputs collection is missing");

      const date = new Date();
      const fileName =
        prompt("Provide the file name without special characthers") ??
        "NO VALID FILE NAME WAS PROVIDED";
      const savePath = prompt(
        "Provide the destination folder",
        `${saveTo}/${fileName}_${date.getFullYear()}${date.getMonth() + 1}${date.getDate()}@${date.getHours()}-${date.getMinutes()}.docx`,
      );
      if (!savePath) return alert("The path for saving the file is not valid");

      inputs.map(([input, col]) => {
        const RT = findRT(input.id) as RT;
        if (RT === Ctrls.currentLease) return; //! We DO NOT update the value of the current lease from the input because the value in the input is the new lease value after revision, not the original value. We need to keep the original value

        if (RT.type === "number") RT.value = fraction(input.valueAsNumber);
        else RT.value = input.value; //If the input.type is "date", the input.value is an ISO date string. So we do not need to make any conversions
      });

      (function setMissingValues() {
        const anniversary = (year: number, date: Date) => {
          date.setFullYear(year);
          return m.getDateString(date);
        };

        const leaseDate = new Date(Ctrls.leaseDate.value); //The value was set to an ISO Date when the Ctrls were updated from the inputs (since the input asscoiated with this Ctrl is of type "date", the input.value is an ISO Date)
        const year = date.getFullYear();

        Ctrls.revisionDate.value = m.getISODate(date); //!This Ctrl is associated with a column in the table, that's why we are setting its value to ISO date in order to update the excel table later with a valid date format

        (function withNoColumn() {
          Ctrls.initialYear.value = getIndexYear(
            Ctrls.initialIndexDate.value as string,
          );
          Ctrls.baseYear.value = getIndexYear(
            Ctrls.baseIndexDate.value as string,
          );
          Ctrls.revisionYear.value = getIndexYear(
            Ctrls.indexDate.value as string,
          );
          Ctrls.anniversaryDate.value = anniversary(year, leaseDate);
          Ctrls.nextRevision.value = anniversary(year + 1, leaseDate);
          Ctrls.startingMonth.value = `${new Intl.DateTimeFormat("fr-FR", { month: "long" }).format(date)} ${year.toString()}`;
        })();

        function getIndexYear(isoDate: string) {
          //!the date passed at this stage is an ISO date formated as "YYYY-MM-DD" (The conversion was  done when the Ctrls values were updated from the inputs). We do not need to convert it as a date from Excel.
          const newDate = new Date(isoDate);
          const month = newDate.getMonth();
          if (month < 3) {
            //if the date of publication of the index is within the 1st quarter of the year, it means the index is the index of Q4 of the previous year
            return newDate.getFullYear() - 1;
          } else {
            //The year of the index is the same year as the year of its publication
            return newDate.getFullYear();
          }
        }
      })();

      const decimals = [
        Ctrls.initialIndex,
        Ctrls.index,
        Ctrls.baseIndex,
        Ctrls.currentLease,
        Ctrls.newLease,
      ]; //Those are the ctrls for which we will replace the '.' decimal with a ',' decimal mark

      const contentControls: [string, string][] = ctrls.map((RT) => {
        if (RT.type === "date")
          return [RT.title, m.getDateString(new Date(RT.value) || null)];
        else if (decimals.includes(RT))
          return [RT.title, (RT.value as number).toFixed(2).replace(".", ",")]; //!We must NOT do this on the Ctrls object directly. We need the values of these Ctrls to remain numbers in order to update the Excel table.
        else return [RT.title, RT.value.toString()];
      });

      try {
        await graph.createAndUploadDocumentFromTemplate(
          templatePath,
          savePath,
          "FR",
          undefined,
          { nestedCtrls: contentControls },
        );
        await updateExcelTable();
        m.spinner(false); //We hide the spinner
      } catch (error) {
        console.log(error);
        alert(error);
        m.spinner(false); //We hide the spinner
      }

      async function updateExcelTable() {
        Ctrls.currentLease.value = Ctrls.newLease.value; //!This must be done at this stage NOT EARLIER, otherwise, we will lose the value of the original lease when editing the Word template. Notice  that Ctrls.newLease is not associated with a column of the Excel table, (i.e., Ctrls.newLease.col property is undefined), which means the row[] will not updated from this Ctrl.
        if (row && rowIndex) {
          await graph.updateExcelTableRow(tableName, rowIndex, update(row));
        } else {
          row = Array(inputs.length);
          await graph.addRowToExcelTable(update(row), rowIndex, tableName);
        }
        function update(row: any[]) {
          ctrls
            .filter((ctrl) => ctrl.col)
            .forEach(({ value, col }) => (row[col!] = value));
          return row;
        }
      }
    }
  }

  async searchFiles() {
    m.spinner(true); //We show the spinner
    const form = this.form;
    if (!form) return;
    const graph = new m.GraphAPI("");
    (function showForm() {
      form.innerHTML = "";
      if (localStorage.folderPath)
        fetchAllDriveFiles(form, localStorage.folderPath); //We will delete the record for this folder path from the database
      (function RegExpInput() {
        const regexp = document.createElement("input");
        regexp.id = "search";
        regexp.classList.add("field");
        regexp.placeholder =
          "Enter your file name search as a regular expression";
        regexp.onkeydown = (e) =>
          e.key === "Enter" ? fetchAllDriveFiles(form!) : e.key;
        form.appendChild(regexp);
      })();
      (function dateAfterInput() {
        const after = document.createElement("input");
        after.type = "date";
        after.id = "after";
        after.classList.add("field");
        after.title =
          "You can proivde the date after which the file was created";
        form.appendChild(after);
      })();
      (function dateAfterInput() {
        const before = document.createElement("input");
        before.type = "date";
        before.id = "before";
        before.title =
          "You can provide the date before which the file was created";
        before.classList.add("field");
        form.appendChild(before);
      })();
      (function fileTypeInput() {
        const mime = document.createElement("input");
        mime.classList.add("field");
        mime.placeholder = "Enter the mime type of the file";
        form.appendChild(mime);
      })();
      (function folderPathInput() {
        const folder = document.createElement("input");
        folder.id = "folder";
        folder.placeholder = "Proide the path for the folder";
        folder.classList.add("field");
        if (localStorage.folderPath) folder.value = localStorage.folderPath;
        form.appendChild(folder);
      })();

      (function searchBtn() {
        const btn = document.createElement("button");
        form.appendChild(btn);
        btn.classList.add("button");
        btn.innerText = "Search";
        btn.onclick = () => fetchAllDriveFiles(form!);
      })();

      (function insertTable() {
        document.querySelector("table")?.remove();
        const table = document.createElement("table");
        form.insertAdjacentElement("afterend", table);
      })();
    })();

    async function fetchAllDriveFiles(form: HTMLElement, record?: string) {
      if (record) return manageFilesDatabase([], record, true); //We delete the record for the folder path

      try {
        await fetchAndFilter();
        m.spinner(false); //Hide the spinner
      } catch (error) {
        m.spinner(false); //Hide the spinner
        console.log(error);
        alert(error);
      }

      async function fetchAndFilter() {
        const files = await fetchAllFilesByBatches();
        if (!files)
          throw new Error("Could not fetch the files list from onedrive");
        const search = form.querySelector("#search") as HTMLInputElement;
        if (!search) throw new Error("Did not find the serch input");
        // Filter files matching regex pattern
        const matchingFiles = filterFiles(files, search.value);

        // Get reference to the table

        const table = document.querySelector("table");
        if (!table) throw new Error("The table element was not found");
        table.innerHTML =
          '<tr class ="fileTitle"><th>File Name</th><th>Created Date</th><th>Last Modified</th></tr>'; // Reset table
        const docFragment = new DocumentFragment();
        docFragment.appendChild(table); //We move the table to the docFragment in order to avoid the slow down related to the insertion of the rows directly in the DOM

        for (const file of matchingFiles) {
          // Populate table with matching files
          const row = table.insertRow();
          row.classList.add("fileRow");
          row.insertCell(0).textContent = file.name;
          row.insertCell(1).textContent = new Date(
            file.createdDateTime,
          ).toLocaleString();
          row.insertCell(2).textContent = new Date(
            file.lastModifiedDateTime,
          ).toLocaleString();
          const link = await getDownloadLink(file.id);
          // Add double-click event listener to open file
          row.addEventListener("dblclick", () => {
            window.open(link, "_blank");
          });
        }

        form.insertAdjacentElement("afterend", table);

        console.log(
          `Fetched ${files.length} items, displaying ${matchingFiles.length} matching files.`,
        );
      }

      async function getDownloadLink(fileId: string) {
        const data = await JSONFromGETRequest(
          `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`,
        );

        return data?.webUrl;
      }

      async function fetchAllFilesByBatches() {
        const path = (byID("folder") as HTMLInputElement)?.value;
        if (!path) throw new Error("The file path could not be retrieved");

        const allFiles: fileItem[] = [];
        const existing = await manageFilesDatabase(allFiles, path);
        if (existing.length) return existing as fileItem[];

        localStorage.folderPath = path;
        const select =
          "$select=name,id,folder,file,createdDateTime,lastModifiedDateTime";
        const top = "$top=900";
        await fetchAllFilesByPath(path);

        return await manageFilesDatabase(allFiles, path);

        async function fetchAllFilesByPath(path: string) {
          // Step 1: Get root-level files & folders
          path = path.replace("\\", "/");
          const topLevelItems = await fetchTopLevelFiles(path);
          const [files, folders] = getFilesAndFolders(topLevelItems);
          allFiles.push(...files);

          // Step 2: Filter folders & fetch their contents using $batch
          const folderIds: string[] = folders.map((f) => f.id);

          await fetchSubfolderContents(folderIds);

          console.log(`Fetched ${allFiles.length} files.`);
          return allFiles;
        }

        async function fetchTopLevelFiles(path: string) {
          const id = await getFolderIdByPath(path);
          const url = `https://graph.microsoft.com/v1.0/me/drive/items/${id}/children?${top}&${select}`;

          const data = await JSONFromGETRequest(url);
          return data?.value as (fileItem | folderItem)[]; // Returns an array of files & folders

          async function getFolderIdByPath(path: string): Promise<string> {
            const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${path}`;
            const data = await JSONFromGETRequest(endpoint);
            return data?.id; // Folder ID
          }
        }

        async function fetchSubfolderContents(folderIds: string[]) {
          const batchUrl = "https://graph.microsoft.com/v1.0/$batch";

          // Create batch request for each folder
          const batchRequests = folderIds.map((folderId, index) => ({
            id: `${index + 1}`,
            method: "GET",
            url: `/me/drive/items/${folderId}/children?${top}&${select}`,
          }));
          const limit = 20;
          for (let i = 0; i < batchRequests.length; i += limit) {
            const batchData = await fetchRequests(
              batchRequests.slice(i, i + limit),
            );
            await processItems(batchData);
          }

          async function fetchRequests(requests: any[]) {
            const body = { requests: requests };
            const response = await graph.sendRequest(
              batchUrl,
              "POST",
              body,
              undefined,
              "application/json",
              "Error fetching subfolders",
            );
            if (!response?.ok) return;
            return await response?.json();
          }

          async function processItems(data: any) {
            // Extract file lists from batch responses
            const items = data.responses
              .map((res: any) => res.body.value)
              .flat() as (fileItem | folderItem)[];
            const [files, folders] = getFilesAndFolders(items);
            allFiles.push(...files);
            const subfolderIds = folders.map((f) => f.id);
            await fetchSubfolderContents(subfolderIds);
          }
        }
      }

      function getFilesAndFolders(
        items: (fileItem | folderItem)[],
      ): [fileItem[], folderItem[]] {
        return [getFiles(items), subFolders(items)];
      }
      function subFolders(items: (fileItem | folderItem)[]) {
        return items.filter(
          (item) => (item as folderItem)?.folder,
        ) as folderItem[];
      }
      function getFiles(items: (fileItem | folderItem)[]) {
        return items.filter((item) => (item as fileItem)?.file) as fileItem[];
      }
      async function JSONFromGETRequest(url: string) {
        const response = await graph.sendRequest(
          url,
          "GET",
          undefined,
          undefined,
          undefined,
          "Error fetching items from endpoint",
        );
        if (!response?.ok) return;
        return await response.json();
      }

      function filterFiles(files: fileItem[], search: string) {
        const byName = files.filter((item: any) =>
          RegExp(search, "i").test(item.name),
        );
        const created = (file: fileItem) => new Date(file.createdDateTime);

        const after = (form.querySelector("#after") as HTMLInputElement)
          ?.valueAsDate;
        const before = (form.querySelector("#before") as HTMLInputElement)
          ?.valueAsDate;

        if (after && before)
          return byName.filter(
            (file) =>
              created(file).getTime() > after.getTime() &&
              created(file).getTime() < before.getTime(),
          );
        else if (before)
          return byName.filter(
            (file) => created(file).getTime() < before.getTime(),
          );
        else if (after)
          return byName.filter(
            (file) => created(file).getTime() > after.getTime(),
          );
        else return byName;
      }

      async function manageFilesDatabase(
        files: fileItem[],
        path: string,
        deleteRecord: boolean = false,
      ): Promise<fileItem[]> {
        const dbName = "FileDatabase";
        const storeName = "Files";
        const dbVersion = 1;

        // Open (or create) the database
        const db = await new Promise((resolve, reject) => {
          const request = indexedDB.open(dbName, dbVersion);

          request.onupgradeneeded = function (event) {
            const db = (event.target as IDBOpenDBRequest)?.result;
            if (db.objectStoreNames.contains(storeName)) return;
            db.createObjectStore(storeName, { keyPath: "path" });
            console.log("Object store created successfully.");
          };

          request.onsuccess = function (event) {
            resolve((event.target as IDBOpenDBRequest)?.result);
          };

          request.onerror = function (event) {
            reject(
              "Failed to open database: " +
              (event.target as IDBOpenDBRequest)?.error,
            );
          };
        });

        // Retrieve or add the entry
        return new Promise((resolve, reject) => {
          const transaction = (db as IDBDatabase).transaction(
            storeName,
            "readwrite",
          );
          const store = transaction.objectStore(storeName);

          // Check if an entry with the given path exists
          const getRequest = store.get(path);

          getRequest.onsuccess = function (event) {
            const existingEntry = (event.target as IDBRequest)?.result;
            if (existingEntry && deleteRecord) {
              const deleteRequest = store.delete(path);
              deleteRequest.onsuccess = function () {
                console.log("successfuly deleted the record");
                resolve(files);
              };
              deleteRequest.onerror = function () {
                reject(
                  "Failed to delete the specified record: " +
                  (event.target as IDBRequest)?.error,
                );
              };
            } else if (existingEntry) {
              console.log("Entry found for path:", path);
              resolve(existingEntry.files as fileItem[]); // Return the existing files array
            } else if (!files.length) {
              resolve(files); //We return the empty array
            } else {
              // Add a new entry if it doesn't exist
              const data = { path: path, files: files };
              const addRequest = store.put(data);

              addRequest.onsuccess = function () {
                console.log("New entry added for path:", path);
                resolve(files); // Return the newly added files array
              };

              addRequest.onerror = function (event) {
                reject(
                  "Failed to add new entry: " +
                  (event.target as IDBRequest)?.error,
                );
              };
            }
          };

          getRequest.onerror = function (event) {
            reject(
              "Failed to retrieve entry: " +
              (event.target as IDBRequest)?.error,
            );
          };
        });
      }
    }
  }

  async updateTableRow() {
    return;
  }

  private findSetting = (name: string, settings: settingInput[] | undefined) =>
    settings?.find((setting) => setting.name === name);

  private getConsts(setting: {
    workBook: string;
    tableName: string;
    wordTemplate: string;
    saveTo: string;
  }) {
    const workbookPath =
      this.findSetting(setting.workBook, this.stored)?.value ||
      prompt("Provide the Excel workbook path") ||
      "";
    const tableName =
      this.findSetting(setting.tableName, this.stored)?.value ||
      prompt("Provide the name of the Excel table containing the data") ||
      "";
    const templatePath =
      this.findSetting(setting.wordTemplate, this.stored)?.value ||
      prompt("Provide the path for the Word invoice template") ||
      "MISSING TEMPLATE PATH";
    const saveTo =
      this.findSetting(setting.saveTo, this.stored)?.value ||
      prompt(
        "Provide teh path for the folder where the invoice should be saved",
      ) ||
      "MISSING SAVETO PATH";
    return { workbookPath, tableName, templatePath, saveTo };
  }

  /**
   * Updates the data list or the value of bound inputs according to the value of the input that has been changed
   * @param {number} index - the dataset.index of the input that has been changed
   * @param {any[][]} table - The table that will be filtered to update the data list of the button. If undefined, it means that the data list will not be updated.
   * @param {boolean} combine - If true, it means that the dataList of the next bound input, will include an additional option combining all the options in the dataList
   * @returns
   */
  private inputOnChange(
    index: number,
    inputs: InputCol[],
    table: any[][] | undefined,
    combine: boolean,
  ): [any[], number] | void {
    if (!table?.length) return;

    const filledInputs = inputs.filter(
      ([input, col]) => input.value && col <= index,
    ); //Those are all the inputs that the user filled with data

    const filtered = filterTableByInputsValues(filledInputs, table); //We filter the table based on the filled inputs

    if (!filtered.length) return;

    const boundInputs = inputs.filter(([input, col]) => col > index); //Those are the inputs for which we want to create  or update their data lists

    for (const [input, col] of boundInputs) {
      input.value = ""; //We reset the value of all bound inputs.
      const list = m.getUniqueValues(col, filtered);
      const row = fillBound(list, input);
      if (row) return [row, table.indexOf(row)]; //!If the function returns true, it means that we filled the value of all the bound inputs, so we break the loop. If it returns false, it means that there is more than one value in the list, so we need to create or update the data list of the input.
      //if (fillBound(list, input)) break;//!If the function returns true, it means that we filled the value of all the bound inputs, so we break the loop. If it returns false, it means that there is more than one value in the list, so we need to create or update the data list of the input.
      populateSelectElement(input, list, combine);
    }

    function fillBound(list: any[], input: HTMLInputElement): void | any[] {
      if (list.length > 1) return;
      const value = list[0],
        found = filtered.length < 2;
      if (!found) return setValue(input, value); //If the filtered array contains more than one row with the same unique value in the corresponding column, we will not fill the next inputs
      const row = filtered[0]; //This is the unique row in the filtered list, we will use it to fill all the other inputs
      boundInputs.forEach(([input, col]) => setValue(input, row[col]));
      return row;
    }

    function setValue(input: HTMLInputElement, value: any) {
      if (input.type === "date")
        input.value = m.getISODate(dateFromExcel(value)); //!We must convert the dates from Excel, and pass the ISO date to the input value (NOT to the input.valueAsDate) in order to avoid the timezone offset issue when using input.valueASDate
      else input.value = value?.toString() || "";
    }
  }
}

export class Marianne implements IGraphAPI {
  private UI: IUserInterface;
  private report: setting = {};
  private stored;
  private form;
  private tenantID;
  private settingsNames;
  private workbookPath;
  private graph;
  private Ctrls: ReportsCtrls = {
    monthly: [
      {
        title: "",
        value: "",
        label: "",
        col: undefined,
        type: "",
      },
    ],
    annual: [
      {
        title: "",
        value: "",
        label: "",
        col: undefined,
        type: "",
      },
    ],
    production: [
      {
        title: "",
        value: "",
        label: "",
        col: undefined,
        type: "",
      },
    ],
    returned: [
      {
        title: "",
        value: "",
        label: "",
        col: undefined,
        type: "",
      },
    ],
  };
  private datesColumns = [3, 4, 16];

  constructor() {
    this.form = byID();
    this.UI = new MarianneUI(this);
    this.stored = saveSettings(this.UI, undefined, true) || undefined;
    this.settingsNames = settingsNames;
    this.workbookPath =
      this.findSetting(this.settingsNames.invoices.wordTemplate, this.stored)
        ?.value ??
      prompt("The path for the Excel workbook is missing") ??
      alert("the workbook path is missing");
    this.graph = new m.GraphAPI(undefined, this.workbookPath ?? "");
    this.tenantID = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
  }

  getUI = () => this.UI;

  async reportFactory(columns: number[], callBack: Function) {
    const datesColumns = this.datesColumns,
      appendInputsAndLabels = this.appendInputsAndLabels,
      form = this.form,
      findSetting = this.findSetting,
      stored = this.stored,
      settingsNames = this.settingsNames;
    const tableName =
      this.findSetting(this.settingsNames.Marianne.tableName, this.stored)
        ?.value ??
      prompt("Provide the name of the Excel table") ??
      "";
    const tableRows = await this.graph.fetchExcelTable(tableName, true);
    if (!tableRows)
      return m.throwAndAlert("Could not retrieve the Excel table");
    const tableTitles = tableRows![0];
    showInputs();

    function showInputs() {
      //Show the user form for filtering the table. The user from will show inputs for the relevant columns for the report
      const inputs = columns.map((col) => {
        let type = "text";
        if (datesColumns.includes(col)) type = "date";
        const input = appendInputsAndLabels(
          col.toString(),
          tableTitles[col],
          type,
          form!,
        );
        return [input, col] as InputCol;
      });

      (function addCreateInput() {
        const button = document.createElement("button");
        button.innerText = "Issue Report";
        button.onclick = () => prepareData(inputs);
      })();

      (function dateFromTo() {
        //!date inputs to fix the period fo the report
      })();

      (function homeBtn() {
        //not sure we need it
      })();
    }

    async function prepareData(inputs: InputCol[]) {
      const templatePath =
        findSetting(settingsNames.Marianne.wordTemplate, stored)?.value ??
        prompt(
          "Provide the name of the path for the word document used as a template for the report",
        ) ??
        "";
      const saveTo =
        findSetting(settingsNames.Marianne.saveTo, stored)?.value ??
        prompt(
          "Provide the name of the destination path for saving the report",
        ) ??
        "";

      (async function filterTableRowsLogic() {
        //Use the filter table by inputs value in the LawFirm class
        let filtered = tableRows?.filter((r) =>
          r.map((cell, index) => columns.includes(index)),
        );
        filtered = filterTableByInputsValues(inputs, filtered!);

        callBack(filtered, { templatePath, saveTo, tableName });
      })();
    }
  }

  async monthlyReport(
    filtered?: any[][],
    args?: { templatePath: string; saveTo: string; tableName: string },
  ) {
    if (!filtered)
      return await this.reportFactory([1, 5, 9], this.monthlyReport);
    if (!args) return;
    const { templatePath, saveTo, tableName } = args;

    const wordRows = customize(filtered);
    const contentControls = this.getContentControlsValues(this.Ctrls.monthly);
    await this.issueReport(
      templatePath,
      saveTo,
      [[tableName, wordRows, 0]],
      contentControls,
    );

    function customize(filtered: any[][]) {
      //This function will customize the order of the filtered rows before editing the report
      const wordRows: any[][] = [];
      return wordRows;
    }
  }

  async annualReport() {
    const ctrls = this.Ctrls.annual;
  }

  async returnedReport() {
    const ctrls = this.Ctrls.returned;
  }

  async productionReport() {
    const ctrls = this.Ctrls.production;
  }

  async noticesReport() {
    const ctrls = this.Ctrls.production;
  }

  async addNewEntry(): Promise<void> { }

  getContentControlsValues(ctrls: RT[]): [string, string][] {
    return [["", ""]];
  }

  /**
   * Returns an array containing the values for filling the new rows that will be added to the word table of the report
   * @param args
   * @returns
   */
  getWordTableRows(filtered: any[][]): {
    wordRows: string[][];
    totalsLabels?: string[];
  } {
    return {
      wordRows: [[""]],
    };
  }

  async updateTableRow(...args: any[]): Promise<void> { }

  private async issueReport(
    templatePath: string,
    saveTo: string,
    tables: [string, string[][], number][],
    contentControls: [string, string][],
  ) {
    const form = this.form;

    await this.graph.createAndUploadDocumentFromTemplate(
      templatePath,
      saveTo,
      "AR",
      tables,
      { nestedCtrls: contentControls },
      undefined,
    );
  }

  findSetting(name: string, settings: settingInput[] | undefined) {
    return settings?.find((setting) => setting.name === name);
  }

  appendInputsAndLabels(
    id: string,
    label: string,
    type: string,
    form: HTMLElement,
  ): HTMLInputElement {
    console.log(id, label, type);
    const div = document.createElement("div");
    form.appendChild(div);
    appendLable();
    return appendInput();
    function appendLable() { }
    function appendInput() {
      const input = document.createElement("input");
      input.id = id;
      input.dataset.index = id;
      div.appendChild(input);
      return input;
    }
  }
  /**
   * UNUSED: This function isn't used in Marianne Class
   * @returns
   */
  async issueLetter(): Promise<void> {
    console.log("this is not a valid method in Marianne Class");
    return;
  }
  /**
   * UNUSED: This function isn't used in Marianne Class
   * @returns
   */
  async issueLeaseLetter(): Promise<void> {
    return console.log("this is not a valid method in Marianne Class");
  }
  /**
   * UNUSED: This function isn't used in Marianne Class
   * @returns
   */
  async issueInvoice(): Promise<void> {
    return console.log("this is not a valid method in Marianne Class");
  }

  /**
   * UNUSED: This function isn't used in Marianne Class
   * @returns
   */
  async searchFiles(): Promise<void> {
    return console.log("this is not a valid method in Marianne Class");
  }
}
class Reports {
  private today;
  private cols = {
    ReceiptNumber: 0,
    ReceiptDate: 1,
    RegisterNumber: 2,
    FileNumber: 3,
    CaseNumber: 4,
    CaseYear: 5,
    JudiciaryCaseYear: 6,
    CaseCourt: 7,
    CaseType: 8,
    ClaimantName: 9,
    DefendantName: 10,
    TransferDate: 11,
    ReceiptionDate: 12,
    FirstMeetingDate: 13,
    CurrrentStatus: 14,
    LastMeetingDate: 15,
    EndOfTreatmentDate: 16,
    AchievementMonth: 17,
    AchievementType: 18,
    ReturnedPreviousExpertName: 19,
    ReturnedRegisterNbr: 20,
    ReturnedDateOfPreviousReport: 21,
    ReturnedAchievementType: 22,
    PartyName: 23,
    PartyAddress: 24,
    Observations: 25,
  };
  private status = {
    notStarted: "لم يبدأ",
    ongoing: "جاري",
    achieved: "منتهي",
    issued: "منتهي وصدر",
  };
  private templates = {
    returned: {
      fileName: "Returned Report - Form 4 Template.docx",
      tableName: "????",
    },
    notices: {
      fileName: "Notices Template.docx",
      tableName: "????",
    },
    noticesReport: {
      fileName: "Notices Report Template.docx",
      tableName: "????",
    },
    monthly: {
      fileName: "Monthly Production Report Template.docx",
      tableName: "MonthlyReport",
    },
    followUp: {
      fileName: "Follow Up Report - Form 5 Template.docx",
      tableName: "????",
    },
    pending: {
      fileName: "Form 2 Template.docx",
      tableName: "????",
    },
  };
  private form = byID();

  private idCols = [
    this.cols.CaseNumber,
    this.cols.ReceiptDate,
    this.cols.RegisterNumber,
    this.cols.CaseYear,
    this.cols.CaseCourt,
    this.cols.ClaimantName,
    this.cols.DefendantName,
  ]; //! those will just serve to avoid adding duplicate cases to the Set()

  constructor() {
    this.today = new Date();
  }

  async monthlyProductionReport() {
    const excelData = await this.fetchExelData();
    if (!excelData) return;
    const titles = excelData[0];
    this.form!.innerHTML = "";

    const cols = this.cols,
      status = this.status,
      rowID = this.rowID,
      editReport = this.editReport,
      createInput = this.createInput,
      createBtn = this.createBtn,
      template = this.templates.monthly;

    function userForm() {
      const dateInput = createInput("month", "Date", "date");
      const reportBtn = createBtn();

      reportBtn.onclick = () => report(dateInput.valueAsDate);
    }

    async function report(date: Date | null) {
      if (!date) return;
      const getMonth = (date: Date) => date.getMonth();
      const issued = excelData
        .filter((row) => row[cols.CurrrentStatus] === status.issued)
        .filter(
          (row) =>
            getMonth(dateFromExcel(row[cols.AchievementMonth])) ===
            getMonth(date),
        );

      const contentControls: [string, string][] = [["????", "????"]];

      const unique = new Set<string>(); //!We must recreate it for each type.

      issued.forEach((row) => unique.add(rowID(row)));

      const cases = Array.from(unique).map((id) =>
        issued.find((row) => rowID(row) === id),
      );

      const wdRows = cases.map((row, i) => {
        return [
          (i + 1).toString(), //Serial Number
          `${row![cols.ReceiptNumber]}/${dateFromExcel(row![cols.ReceiptDate]).getFullYear()}`,
          `${row![cols.CaseNumber]} لسنة ${row![cols.CaseYear] || row![cols.JudiciaryCaseYear]} محكمة ${row![cols.CaseCourt]}`,
          `${row![cols.ClaimantName]} x ${row![cols.DefendantName]}`,
          row![cols.CaseType],
          row![cols.AchievementType],
        ];
      });
      await editReport(wdRows, template, contentControls);
    }
  }

  async returnedCasesReport() {
    const excelData = await this.fetchExelData();
    if (!excelData) return;
    const titles = excelData[0];
    const cols = this.cols,
      status = this.status,
      rowID = this.rowID,
      editReport = this.editReport,
      getFirst = this.findCaseFirstRow,
      template = this.templates.returned,
      tableName = "???????";

    function userForm() { }

    async function report() {
      const returned = excelData.filter(
        (row) =>
          row[cols.ReturnedRegisterNbr] &&
          row[cols.CurrrentStatus] !== status.issued,
      );

      const contentControls: [string, string][] = [["????", "????"]];

      const unique = new Set<string>(); //!We must recreate it for each type.

      returned.forEach((row) => unique.add(rowID(row)));

      const cases = Array.from(unique).map((id) =>
        returned.find((row) => rowID(row) === id),
      );

      const wdRows = cases.map((row, i) => {
        return [
          (i + 1).toString(), //Serial Number
          `${row![cols.ReceiptNumber]}/${dateFromExcel(row![cols.ReceiptDate]).getFullYear()}`,
          `${row![cols.CaseNumber]} لسنة ${row![cols.CaseYear] || row![cols.JudiciaryCaseYear]} محكمة ${row![cols.CaseCourt]}`,
          `${row![cols.ClaimantName]} x ${row![cols.DefendantName]}`,
          m.getDateString(row![cols.ReceiptionDate]),
          row![cols.ReturnedPreviousExpertName],
          row![cols.ReturnedAchievementType],
          row![cols.CurrrentStatus],
          row![cols.ReturnedDateOfPreviousReport],
        ];
      });
      await editReport(wdRows, template, contentControls);
    }
  }

  async generateNotices() {
    const excelData = await this.fetchExelData();
    if (!excelData)
      return m.throwAndAlert("Failed to fetch the data from the Excel table");

    this.form!.innerHTML = "";
    const titles = excelData[0];
    const cols = this.cols,
      rowID = this.rowID,
      today = this.today,
      createInput = this.createInput,
      createBtn = this.createBtn;

    function userForm() {
      //The user form will send an array where for each case = [parties:string[], {date:string, hour:string, AM/PM: string}]
      let filtered: any[][]; //These are the rows that would have been found by filtering the table from the UI inputs

      const casesArray: string[][] = [];

      const meetingDate: HTMLInputElement = createInput(
        "meetingDate",
        "Meeting Date",
      ),
        meetingHour: HTMLInputElement = createInput(
          "meetingHour",
          "Meeting Hour",
        );

      const [claimant, defendant, caseNbr, caseCourt, caseType] = [
        cols.ClaimantName,
        cols.DefendantName,
        cols.CaseCourt,
        cols.CaseType,
      ].map((col) => createInput(`input${col}`, titles[col]));

      const addCaseBtn = createBtn();

      const inputs: InputCol[] = [
        [claimant, cols.ClaimantName],
        [defendant, cols.DefendantName],
        [caseNbr, cols.CaseNumber],
        [caseCourt, cols.CaseCourt],
        [caseType, cols.CaseType],
      ];

      const getInputCol = (next: HTMLInputElement) =>
        inputs.find(([input, col]) => input === next);
      claimant.onchange = () =>
        filterOnChange(getInputCol(claimant)!, excelData);

      //issueNotices.onchange = () => report([[filtered, { date, hour }]]);

      function addCase() {
        const addCaseBtn = createBtn();
        addCaseBtn.onclick = () => {
          //inputs.map()
        };
      }

      function filterOnChange(inputCol: InputCol, data: any[][]) {
        const [input, col] = inputCol;

        //const relevant = (row: any[]) => inputs.map(([input, col]) => row[col]);

        const _filtered = data.filter((row) => row[col] === input.value);

        if (!_filtered.length)
          return m.throwAndAlert(
            "The value provided did not correspond to any result",
          );

        filtered = _filtered;

        const filled = inputs.filter(
          (el, index) => index <= inputs.indexOf(inputCol),
        );

        const bound = inputs.filter(
          (el, index) => index > inputs.indexOf(inputCol),
        );

        if (found()) {
          const row = filtered[0];
          bound.forEach(([input, col]) => (input.value = row[col]));
          return showParties([row]);
        }

        bound.forEach(([input, col]) => {
          const uniqueValues = m.getUniqueValues(col, filtered);
          populateSelectElement(input, uniqueValues);
          if (bound.indexOf([input, col]) > bound.length - 2) return;
          input.onchange = () => filterOnChange([input, col], filtered);
        });

        function found() {
          const cells = (row: any[][]) =>
            filled.map(([input, col]) => row[col]);
          const first = filtered[0];
          return filtered.every((row) => cells(row) === cells(first));
        }
      }

      function showParties(relevant: any[][]) {
        //Show the name of the case with a checkbox in a div

        const container = byID("casesList");
        relevant.forEach((row) => {
          const [name, adress] = [
            relevant[cols.ClaimantName],
            relevant[cols.PartyAddress],
          ];
          const div = document.createElement("div");
          container?.appendChild(div);
          const checkBox = document.createElement("input");
          checkBox.type = "checkbox";
          div.appendChild(checkBox);
          const label = document.createElement("label");
          label.innerText = `${name} : ${adress}`;
          div.appendChild(label);

          checkBox.onchange = () => {
            if (checkBox.checked) casesArray.push(row);
            else casesArray.splice(casesArray.indexOf(row), 1);
          };
        });
      }
    }

    function report(
      parties: [
        any[],
        { meetingDate: HTMLInputElement; meetingHour: HTMLInputElement },
      ][],
    ) {
      const noticeCtrl = "RTNotice";

      const ctrls = {
        partyName: {
          title: "RTPartyName",
          col: cols.PartyName,
        },
        partyAdress: {
          title: "RTPartyAdress",
          col: cols.PartyAddress,
        },
        registreNbr: {
          title: "RTRegisterNbr",
          col: cols.RegisterNumber,
        },
        receiptDate: {
          title: "RTReceiptDate",
          col: cols.ReceiptDate,
        },
        caseNbr: {
          title: "RTCaseNumbr",
          col: cols.CaseNumber,
        },
        caseYear: {
          title: "RTCaseYear",
          col: cols.CaseYear,
        },
        caseCourt: {
          title: "RTCaseCourt",
          col: cols.CaseCourt,
        },
        claimantName: {
          title: "RTClaimant",
          col: cols.ClaimantName,
        },
        defendantName: {
          title: "RTDefendant",
          col: cols.DefendantName,
        },
        meetingDate: {
          title: "RTMeetingDate",
          col: undefined,
        },
        meetingHour: {
          title: "RTMeetingHour",
          col: undefined,
        },
        meetingAmPm: {
          title: "RTAmPm",
          col: undefined,
        },
        noticeDate: {
          title: "RTToday",
          col: undefined,
        },
      };
      const _ctrls = Object.values(ctrls);

      parties.forEach(([party, { meetingDate, meetingHour }]) => {
        const contentControls = _ctrls.map((ctrl) => {
          const date = meetingDate.valueAsDate,
            hour = meetingHour.valueAsNumber;
          let value = "";
          if (ctrl.col) value = party[ctrl.col];
          else if (ctrl === ctrls.meetingDate) value = m.getDateString(date);
          else if (ctrl === ctrls.meetingHour) value = hour.toString();
          else if (ctrl === ctrls.meetingAmPm)
            value = hour >= 1 && hour < 8 ? (value = "مساءً") : "صباحاً";
          return [ctrl.title, value];
        });
      });
    }
  }

  async noticesReport() {
    const excelData = await this.fetchExelData();
    if (!excelData) return;
    function userForm() { }

    function report() { }
  }

  async pendingCasesReport() {
    const excelData = await this.fetchExelData();
    if (!excelData) return;
    const casesTypes: { [index: string]: { name: string; tableName: string } } =
    {
      //!the types must be in this order !!!
      CivilAppealed: {
        name: "مدني مستأنف",
        tableName: "CivilLawCases",
      },
      Administrative: {
        name: "قضاء إداري",
        tableName: "CivilLawCases",
      },
      Civil: {
        name: "مدني",
        tableName: "CivilLawCases",
      },
      LabourAppealed: {
        name: "عمال مستأنف",
        tableName: "LabourLawCases",
      },
      Labour: {
        name: "عمال",
        tableName: "LabourLawCases",
      },
      Tax: {
        name: "ضرائب",
        tableName: "TaxLawCases",
      },
      Persons: {
        name: "أحوال شخصية",
        tableName: "PersonsLawCases",
      },
      Criminal: {
        name: "جنح ونيابات",
        tableName: "CriminalLawCases",
      },
      PublicCasess: {
        name: "أموال عامة",
        tableName: "PublicFundsLawCases",
      },
    };
    const cols = this.cols,
      status = this.status,
      getFirst = this.findCaseFirstRow,
      rowID = this.rowID,
      editReport = this.editReport,
      today = this.today,
      template = this.templates.monthly; //!Needs to be checked

    function userForm() { }

    async function report() {
      //!The Word template has several tables types that we will use depending on the case type:
      const contentControls: [string, string][] = [
        ["RTReportingPeriod", m.getDateString(today)],
      ];

      const ongoing = excelData.filter(
        (row) => row[cols.CurrrentStatus] !== status.issued,
      );

      const identifiers = Object.values(casesTypes).map((type) => {
        const idsSet = new Set<string>(); //!We must recreate it for each type.
        const sameType = ongoing.filter(
          (row) => row[cols.CaseType] === type.name,
        );
        sameType.forEach((row) => idsSet.add(rowID(row)));
        return { idsSet, tableName: type.tableName };
      });

      const setsArray = Array.from(identifiers);

      for (const { idsSet, tableName } of setsArray) {
        //!We use for-of loop because forEach() does not await
        const cases = Array.from(idsSet).map((id) =>
          excelData.find((row) => rowID(row) === id),
        ); //! we return the first row for each case. This is the row that we update each time, the other rows contain only the parties names and adresses.

        const wdRows = cases
          .filter((row) => row?.length)
          .map((row, i) => {
            return [
              (i + 1).toString(), //Serial Number
              row![cols.RegisterNumber] || row![cols.ReceiptNumber],
              `${row![cols.CaseNumber]} ${row![cols.CaseCourt]}`,
              `${row![cols.ClaimantName]} x ${row![cols.DefendantName]}`,
              row![cols.TransferDate],
              row![cols.ReceiptionDate],
              row![cols.FirstMeetingDate],
              row![cols.CurrrentStatus] === status.notStarted ? "x" : "",
              row![cols.CurrrentStatus] === status.ongoing ? "x" : "",
              row![cols.CurrrentStatus] === status.issued ? "x" : "", //!needs to be checked. If issued, it should not be in the filered pending cases array
              row![cols.AchievementType],
              row![cols.ReturnedPreviousExpertName],
              row![cols.ReturnedRegisterNbr],
              row![cols.Observations],
            ];
          });

        await editReport(
          wdRows,
          { fileName: "PendingCasesReport", tableName },
          contentControls,
        );
      }
    }
  }

  private async editReport(
    wdRows: any[][],
    { fileName, tableName }: { fileName: string; tableName: string },
    contentControls: [string, string][],
  ) {
    //fetch the word template, get the xml files
    //pass the cases array and the tableName to the function that will find it in the document.xml
    //for each pending element add a row to this table
    //we might need a logic for adding totals in the last row of the table
  }

  private async fetchExelData(): Promise<any[][]> {
    return [];
  }

  private showUserForm() { }
  private rowID(row: any[]) {
    return this.idCols.map((col) => row[col]).join("&");
  }

  private findCaseFirstRow(row: any[], tblRows: any[][]) {
    return tblRows.find((r) => this.rowID(r) === this.rowID(row));
  }
  /**
   * Replaces: Public Sub AddCasesToReport(myUserForm As Object)
   * !! I don't understand what this function was supposed to do
   */
  AddCasesToReport(newRows: any[][], myUserForm: any): void {
    let visibleCells: any;
    const reportCasesArray = [];
    const cols = this.cols;

    function rewritten(
      meetingDate: HTMLInputElement,
      meetingHour: HTMLInputElement,
      meetingAmPm: HTMLInputElement,
    ) {
      const tableColumns = [
        cols.ClaimantName,
        cols.DefendantName,
        cols.CaseNumber,
        cols.CaseType,
        cols.CaseCourt,
        cols.ReceiptNumber,
        cols.ReceiptDate,
        cols.CaseYear,
        cols.JudiciaryCaseYear,
      ];

      tableColumns.forEach((tableColumn, i) => {
        reportCasesArray.push([
          "",
          ...(visibleCells.Cells(tableColumn).Value as any[]),
        ]);
      });

      const contentControls = [
        meetingDate!.value,
        meetingHour!.value,
        meetingAmPm!.value,
      ].map((value) => ["Ctrl title", value]);
    }

    //@ts-ignore
    visibleCells = myTable.filter((rows) => rows)[0]; // This should be a single row passed to the function and the function will add it to the reportCasesArray. In VBA, the SpecialCells(xlCellTypeVisible) returns a range of visible cells after applying a filter. The Areas(1) part is used to get the first contiguous range of visible cells, and then Rows(1) is used to get the first row of that range. In this TypeScript translation, we will assume that the function receives the relevant row of data directly as an argument, so we can skip the filtering step here and directly use that row for our data mapping.

    (function vbaLogic_Unused() {
      return;
      if (
        myUserForm.ComboBox_ClaimantName.text !== "" ||
        myUserForm.ComboBox_ClaimantName.Visible === false
      )
        myUserForm.Label_ConsoleBig.caption = "";
    })();

    (function vbaLogic() {
      //This was aimed at redimensioning the array to add a new case (column) while preserving the existing data. In VBA, ReDim Preserve allows you to resize an array while keeping its contents intact. However, in JavaScript/TypeScript, we can simply push new elements to the existing arrays without needing to manually resize them. The logic below is a direct translation of the VBA approach, but in practice, we can just push new data into our 2D array structure without worrying about dimensions.
      const reportCasesArray = [];
      return;
      // VBA Logic: Expand the 2nd dimension (Columns) by 1 for the new case
      // In this TS representation, reportCasesArray[1 to 12] are the Fields (Rows)
      // reportCasesArray[r][c] where c is the Case
      if (reportCasesArray.length === 0) {
        for (let r = 0; r <= 12; r++) {
          reportCasesArray[r] = [];
        }
      }

      const currentFieldsCount = 12;
      for (let r = 1; r <= currentFieldsCount; r++) {
        // VBA: ReDim Preserve - adding a new "cell" to the end of each field row
        reportCasesArray[r].push(undefined);
      }

      // VBA: For r = 1 To UBound(reportCasesArray, 1): reportCasesArray(r, 1) = "": Next
      // Ensures the first data column (buffer) remains empty if it's the start
      for (let r = 1; r <= 12; r++) {
        reportCasesArray[r][1] = "";
      }
    })();

    // --- Data Mapping (Vertical Slice) ---
    // VBA: reportCasesArray(1, D2Length) = .Cells(ColumnClaimantName.index).Value
    const tableColumns = [
      cols.ClaimantName,
      cols.DefendantName,
      cols.CaseNumber,
      cols.CaseType,
      cols.CaseCourt,
      cols.ReceiptNumber,
      cols.ReceiptDate,
      cols.CaseYear,
      cols.JudiciaryCaseYear,
    ];

    tableColumns.forEach((tableColumn, i) => {
      reportCasesArray.push([
        "",
        ...(visibleCells.Cells(tableColumn).Value as any[]),
      ]);
    });

    // --- UI Field Mapping ---
    [
      myUserForm.TextBox_MeetingDate,
      myUserForm.TextBox_MeetingHour.text,
      myUserForm.ComboBox_AmPm,
    ].forEach((control, i) => {
      reportCasesArray.push(["", control.text]);
    });

    (function vbaLogic() {
      return;
      let rowsLength: number = reportCasesArray[1].length - 1; // Correcting for 0-based array index
      reportCasesArray[1][rowsLength] = visibleCells.Cells(
        cols.ClaimantName,
      ).Value;
      reportCasesArray[2][rowsLength] = visibleCells.Cells(
        cols.DefendantName,
      ).Value;
      reportCasesArray[3][rowsLength] = visibleCells.Cells(
        cols.CaseNumber,
      ).Value;
      reportCasesArray[4][rowsLength] = visibleCells.Cells(cols.CaseType).Value;
      reportCasesArray[5][rowsLength] = visibleCells.Cells(
        cols.CaseCourt,
      ).Value;
      reportCasesArray[6][rowsLength] = visibleCells.Cells(
        cols.ReceiptNumber,
      ).Value;
      reportCasesArray[7][rowsLength] = visibleCells.Cells(
        cols.ReceiptDate,
      ).Value;
      reportCasesArray[8][rowsLength] = visibleCells.Cells(cols.CaseYear).Value;
      reportCasesArray[9][rowsLength] = visibleCells.Cells(
        cols.JudiciaryCaseYear,
      ).Value;
      // --- UI Field Mapping ---
      reportCasesArray[10][rowsLength] = myUserForm.TextBox_MeetingDate.text;
      reportCasesArray[11][rowsLength] = myUserForm.TextBox_MeetingHour.text;
      reportCasesArray[12][rowsLength] = myUserForm.ComboBox_AmPm.text;
    })();
  }

  createInput(
    id: string,
    label: string,
    type: string = "text",
  ): HTMLInputElement {
    const div = document.createAttribute("div");
    this.form?.appendChild(div);
    const lbl = document.createElement("label");
    lbl.innerText = label;
    div.appendChild(lbl);
    const input = document.createElement("input");
    input.type = type;
    div.appendChild(input);
    return input;
  }
  createBtn() {
    return document.createElement("button");
  }
}
/**
 * Convert the date in an Excel row into a javascript date (in milliseconds)
 * @param {number} excelDate - The date retrieved from an Excel cell
 * @returns {Date} - a javascript format of the date
 */
function dateFromExcel(excelDate: number): Date {
  const day = 86400000; //this is the milliseconds in a day
  const dateMS = Math.round((excelDate - 25569) * day); //This gives the days converted from milliseconds.
  //!We have to do this in order to avoid the timezone conversion issues
  const date = new Date(dateMS);
  return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

/**
 * Filters the table according to the values of the inputs. The value of each input is compared to the value of the cell in the corresponding column in the table. If the value of the input is included in the cell value, it means that this row matches the criteria of this input. For a row to be included in the resulting filtered table, it must match the criteria of all the inputs.
 * @param {HTMLInputElement[]} inputs - the html inputs containing the values based on which the table will be filtered
 * @param {any[][]} table - The table that will be filtered
 * @returns {any[][]} - The resulting filtered table
 */
function filterTableByInputsValues(
  inputs: InputCol[],
  table: any[][],
): any[][] {
  const values = inputs.map(
    ([input, col]) => [col, input.value.split(splitter)] as const,
  ); //!some inputs may contain multiple comma separated values if the user has selected more than one option in the data list. So we split the input value by ", " and we check if the cell value is included in the resulting array.
  return table.filter((row) =>
    values.every(([col, value]) => value.includes(row[col])),
  );
}

export function saveSettings(
  ui?: IUserInterface,
  values?: [string, string][],
  get: boolean = false,
) {
  const settings: settings = {
    issueInvoice: {
      workBook: {
        label: "Invoices workbook path :",
        name: settingsNames.invoices.workBook,
        value: "",
      },
      wordTemplate: {
        label: "Invoices'Word template path: ",
        name: settingsNames.invoices.wordTemplate,
        value: "",
      },
      saveTo: {
        label: "Invoices' save to path: ",
        name: settingsNames.invoices.saveTo,
        value: "",
      },
      tableName: {
        label: "Invoices' Excel Table name: ",
        name: settingsNames.invoices.tableName,
        value: "",
      },
    },
    Letter: {
      wordTemplate: {
        label: "Letter Word template path: ",
        name: settingsNames.letter.wordTemplate,
        value: "",
      },
      saveTo: {
        label: "Letter save to path: ",
        name: settingsNames.letter.saveTo,
        value: "",
      },
    },
    leases: {
      workBook: {
        label: "Leases Excel workbook path :",
        name: settingsNames.leases.workBook,
        value: "",
      },
      tableName: {
        label: "Leas's Excel Table name: ",
        name: settingsNames.leases.tableName,
        value: "",
      },
      wordTemplate: {
        label: "Leases Word Template path :",
        name: settingsNames.leases.wordTemplate,
        value: "",
      },
      saveTo: {
        label: "Leases' save to path: ",
        name: settingsNames.leases.saveTo,
        value: "",
      },
    },
  };

  const groups = Object.values(settings);
  const inputs = groups.map((group) => Object.values(group)).flat();

  let stored: settingInput[];
  localStorage.InvoicingPWA
    ? (stored = JSON.parse(localStorage.InvoicingPWA) as settingInput[])
    : (stored = inputs);
  if (get) return stored;

  const findSetting = (name: string, settings: settingInput[]) =>
    settings?.find((setting) => setting.name === name);

  if (values?.length) return save(values); //If the values of some settings have been passed as argument, we save the changes to the localStorage directly withouth showing inputs;

  const form = byID();
  if (!form) return;
  form.innerHTML = "";
  groups.forEach((group) => showInputs(group));

  (function homeBtn() {
    if (ui) showUI(ui, true);
  })();

  function showInputs(group: setting) {
    const groupDiv = document.createElement("div");
    form!.appendChild(groupDiv);
    Object.values(group).forEach((input) =>
      groupDiv.appendChild(createInput(input)),
    );
  }

  function createInput({ label, name, value }: settingInput): HTMLDivElement {
    const container = document.createElement("div");
    const labelHtml = document.createElement("label");
    labelHtml.innerText = label;
    const input = document.createElement("input");
    input.classList.add("field");
    input.value = findSetting(name, stored)!.value || "";
    input.onchange = () => confirmSaving(input.value, label, name);
    container.appendChild(labelHtml);
    container.appendChild(input);
    return container;
  }

  function confirmSaving(value: string, label: string, name: string) {
    if (
      !confirm(
        `Are you sure you want to change the ${label} localStorage value to ${value}?`,
      )
    )
      return;
    save([[name, value]]);
  }
  function save(values: [string, string][]) {
    values.forEach(
      ([name, value]) =>
        (findSetting(name, stored)!.value = value.replaceAll("\\", "/") || ""),
    );
    localStorage.InvoicingPWA = JSON.stringify(stored);
  }
}
