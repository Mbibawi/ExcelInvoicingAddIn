const settingsNames = {
  invoices: {
    workBook: 'invoicesWorkbook',
    tableName: 'invoicesTable',
    wordTemplate: 'invoicesTemplate',
    saveTo: 'invoicesSaveTo',
  },
  letter: {
    workBook: 'letterWorkbook',
    wordTemplate: 'letterTemplate',
    saveTo: 'letterSaveTo',
    tableName:'',
  },
  leases: {
    workBook: 'leasesWorkbook',
    tableName: 'leasesTable',
    wordTemplate: 'leasesTemplate',
    saveTo: 'leasesSaveTo',
  },
};

(function savedSettings() {
  if (localStorage.InvoicingPWA) return;
  const root = 'Legal/Mon Cabinet d\'Avocat/';
  const settings: settingInput[] = [
    {
      name: settingsNames.invoices.workBook,
      value: `${root}Comptabilité/Comptabilité de Mon Cabinet_04 07 2024.xlsm`,
      label: 'Please provide the OneDrive full path (including the file name and extension) for the Word template'
    },
    {
      name: settingsNames.invoices.tableName,
      value: "LivreJournal",
      label: 'Provide the name of the accounts table'
    },
    {
      name: settingsNames.invoices.wordTemplate,
      value: `${root}Comptabilité/Factures/FactureTEMPLATE [NE PAS MODIFIDER].docx`,
      label: 'Please provide the OneDrive full path (including the file name and extension) for the Word template for the invoices'
    },
    {
      name: settingsNames.invoices.saveTo,
      value: `${root}Comptabilité/Factures/Clients}`,
      label: 'Please provide the OneDrive defalut folder path where the generated invoices will be saved'
    },
    {
      name: settingsNames.letter.wordTemplate,
      value: `${root}Administratif/Modèles Actes/Template_Lettre With Letter Head [DO NOT MODIFY].docx`,
      label: 'Please provide the OneDrive full path (including the file name and extension) for the Word template for the letter heads'
    },
    {
      name: settingsNames.letter.saveTo,
      value: `${root}Clients`,
      label: 'Please provide the path of the OneDrive folder where the created letter will be saved'
    },
    {
      name: settingsNames.leases.workBook,
      value: `${root}Clients/LeasesDataBase.xlsm`,
      label: 'Please provide the OneDrive full path (including the file name and extension) for the Leases Excel Workbook'
    },
    {
      name: settingsNames.leases.tableName,
      value: "LEASES",
      label: 'Please provide the Leases Excel Table'
    },
    {
      name: settingsNames.leases.wordTemplate,
      value: `${root}Administratif/Modèles Actes/Template_Révision de loyer [DO NOT MODIFY].docx`,
      label: 'Please provide the OneDrive full path (including the file name and extension) for the Leases Word template'
    },
    {
      name: settingsNames.leases.saveTo,
      value: `${root}Clients`,
      label: 'Please provide the path of the OneDrive folder where the created letter will be saved'
    },
  ];

  const values = settings.map(({ name, value, label }) => {
    const setting = prompt(label, value) || '';
    return [name, setting] as [string, string];
  });
  saveSettings(values);
})();

const byID = (id: string = "form") => document.getElementById(id);

const splitter = "; OR ";//This is the splitter that will be used to separate multiple values in the input fields. We need to use a splitter that is not likely to be included in the values themselves.


(function RegisterServiceWorker() {
  // Check if the browser supports service workers
  if ("serviceWorker" in navigator) {
    window.addEventListener("load", async () => {
      try {
        const registration = await navigator.serviceWorker.register("/ExcelInvoicingAddIn/dist/sw.js");
        console.log("Service Worker registered successfully:", registration);
      } catch (error) {
        console.error("Service Worker registration failed:", error);
      }
    });
  }

  // Handle updates to the service worker
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    console.log("New service worker activated. Reloading page...");
    window.location.reload();
  });

  //@ts-ignore Handling the "beforeinstallprompt" event for PWA installability
  let installPromptEvent: BeforeInstallPromptEvent | null = null;

  //@ts-ignore
  window.addEventListener("beforeinstallprompt", (event: BeforeInstallPromptEvent) => {
    event.preventDefault();  // Prevent the default mini-infobar
    installPromptEvent = event;

    const installButton = document.getElementById("install-button");
    if (installButton) {
      installButton.style.display = "block";  // Show the install button
      installButton.addEventListener("click", async () => {
        if (installPromptEvent) {
          await installPromptEvent.prompt();  // Show the install prompt
          const choiceResult = await installPromptEvent.userChoice;
          console.log("User install choice:", choiceResult.outcome);

          installPromptEvent = null;  // Clear the event after the prompt
          installButton.style.display = "none";  // Hide the button
        }
      });
    }
  });
})();


function loadMsalScript() {
  var token;
  const script = document.createElement("script");
  script.src = "https://alcdn.msauth.net/browser/2.17.0/js/msal-browser.min.js";
  //script.onload = async () => (token = await getTokenWithMSAL());
  script.onerror = () => console.error("Failed to load MSAL.js");
  document.head.appendChild(script);
  token ? console.log('Got token', token) : console.log('could not retrieve the token');
};

async function showForm(id?: string) {

  const form = document.getElementById("form") as HTMLDivElement;
  form.innerHTML = '';
  if (!form) return;

  let table: Excel.Table;

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const stored = getSavedSettings();
    const tableName = stored?.find(setting => setting.name === settingsNames.invoices.tableName)?.value;
    if (!tableName) return alert('The table name was not found')
    table = sheet.tables.getItem(tableName);
    const header = table.getHeaderRowRange();
    header.load('text');
    await context.sync();
    const body = table.getDataBodyRange();
    body.load('text');
    await context.sync();
    const headers = header.text[0];
    const clientUniqueValues: string[] = getUniqueValues(0, body.text);

    if (id === 'entry') await addingEntry(tableName, headers, clientUniqueValues);
    else if (id === 'invoice') invoice(tableName, headers, clientUniqueValues);
  });


  function invoice(tableName: string, title: string[], clientUniqueValues: string[]) {
    const inputs = insertInputsAndLables([0, 1, 2, 3]);//Inserting the fields inputs (Client, Matter, Nature, Date)

    inputs.forEach(input => input?.addEventListener('focusout', async () => await _inputOnChange(input), { passive: true }));

    insertInputsAndLables(['Français', 'English'], true); //Inserting langauges checkboxes
    form.innerHTML += `<button onclick="generateInvoice()"> Generate Invoice</button>`; //Inserting the button that generates the invoice

    function insertInputsAndLables(indexes: (number | string)[], checkBox: boolean = false): HTMLInputElement[] {
      const id = 'input';
      return indexes.map(index => {
        const input = document.createElement('input');
        if (checkBox) input.type = 'checkbox';
        else if (Number(index) < 3) input.type = 'text';
        else input.type = 'date';
        checkBox ? input.id = id : input.id = id + index.toString();
        if (!checkBox) {
          input.name = input.id;
          input.dataset.index = index.toString();
          input.setAttribute('list', input.id + 's');
          input.autocomplete = "on";
        }

        const label = document.createElement('label');
        checkBox ? label.innerText = index.toString() : label.innerText = title[Number(index)];
        label.htmlFor = input.id;

        form.appendChild(label);
        form.appendChild(input);
        if (Number(index) < 1) createDataList(input?.id, clientUniqueValues);//We create a unique values dataList for the 'Client' input
        return input
      });
    };

    async function _inputOnChange(input: HTMLInputElement, unfilter: boolean = false) {
      const index = getIndex(input);

      if (index < 1) unfilter = true;//If this is the 'Client' column, we remove any filter from the table;

      //We filter the table accordin to the input's value and return the visible cells
      const visibleCells = await filterTable(tableName, [{ column: index, value: getArray(input.value) }], unfilter);

      if (visibleCells.length < 1) return alert('There are no visible cells in the filtered table');

      //We create (or update) the unique values dataList for the next input 
      const nextInput = getNextInput(input);
      if (!nextInput) return;
      const list = getUniqueValues(getIndex(nextInput), visibleCells);
      if (list?.length < 2) return nextInput.value = list[0].toString() || '';//If there is only one value in the list, we set it as the value of the input and we don't create a data list for it because there is no need
      populateSelectElement(nextInput, list);

      function getNextInput(input: HTMLInputElement) {
        let nextInput: Element | null = input.nextElementSibling;
        while (nextInput?.tagName !== 'INPUT' && nextInput?.nextElementSibling) {
          nextInput = nextInput.nextElementSibling
        };

        return nextInput as HTMLInputElement
      }

      if (index === 1) {
        //!Need to figuer out how to create a multiple choice input for nature
        const nature = new Set((await filterTable(tableName, undefined, false)).map(row => row[index]));
        nature.forEach(el => form.appendChild(createCheckBox(undefined, el)));
      }

    };

  }

  async function addingEntry(tableName: string, title: string[], uniqueValues: string[]) {
    await filterTable(tableName, undefined, true);

    for (const t of title) {//!We could not use for(let i=0; i<title.length; i++) because the await does not work properly inside this loop
      const i = title.indexOf(t);
      if (![4, 7].includes(i)) form.appendChild(createLable(i));//We exclued the labels for "Total Time" and for "Year"
      form.appendChild(await createInput(i));
    };

    const inputs = Array.from(document.getElementsByTagName('input'));
    inputs
      .filter(input => Number(input?.dataset.index) < 2)
      .forEach(input => input?.addEventListener('change', async () => await onFoucusOut(input), { passive: true }));

    inputs
      .filter(input => [4, 7].includes(Number(input?.dataset.index)))
      .forEach(input => input.style.display = 'none');//We hide the inputs of some columns like the "Total Hours" or the "Link" column


    async function onFoucusOut(input: HTMLInputElement) {

      const i = Number(input.dataset.index);
      const criteria = [{ column: i, value: getArray(input.value) }];
      let unfilter = false;
      if (i === 0) unfilter = true;
      await filterTable(tableName, criteria, unfilter);
      //if (i < 1) createDataList('input' + String(i + 1), getUniqueValues(i + 1, await filterTable(undefined, undefined)));
    }

    form.innerHTML += `<button onclick="addEntry()"> Ajouter </button>`;


    function createLable(i: number) {
      const label = document.createElement('label');
      label.htmlFor = 'input' + i.toString();
      label.innerHTML = title[i] + ':';
      return label
    }

    async function createInput(i: number) {
      const input = document.createElement('input');
      const id = 'input';
      input.name = id + i.toString();
      input.id = input.name;
      input.autocomplete = "on";
      input.dataset.index = i.toString();
      input.type = 'text';
      if ([8, 9, 10].includes(i))
        input.type = 'number';
      else if (i === 3)
        input.type = 'date';
      else if ([5, 6].includes(i))
        input.type = 'time';
      else if ([0, 1, 2, 11, 12, 13, 16].includes(i)) {
        //We add a dataList for those fields
        input.setAttribute('list', input.id + 's');
        createDataList(input?.id, uniqueValues);
      }

      return input

    }

  }

  function createCheckBox(input: HTMLInputElement | undefined, id: string = '') {
    if (!input) input = document.createElement('input');
    input.type = 'checkbox';
    input.id += id;


    return input


  };



}

/**
 *
 * @param select 
 * @param uniqueValues 
 * @param  {boolean} combine - determines whether we will add to the list an element containing all the options. Its defalult value is "false"
 */
function populateSelectElement(select: HTMLInputElement, uniqueValues: string[], combine: boolean = false) {
  const list = createDataList(select.id, uniqueValues, combine);
  if (!list) return;
  select.setAttribute('list', list.id);
  select.autocomplete = "on";
  return list
}

/**
 * 
 * @param id 
 * @param uniqueValues 
 * @param combine 
 * @returns 
 */
function createDataList(id: string, uniqueValues: string[], combine: boolean = false) {
  //const uniqueValues = Array.from(new Set(visible.map(row => row[i])));
  if (!id || uniqueValues?.length < 2) return;
  id += 's';

  // Create a new datalist element
  let dataList = Array.from(document.getElementsByTagName('datalist')).find(list => list.id === id);
  if (dataList) dataList.remove();
  dataList = document.createElement('datalist');
  dataList.id = id;
  // Append options to the datalist
  uniqueValues.forEach(option => addOption(option));

  if (combine)
    addOption(uniqueValues.join(splitter));

  // Attach the datalist to the body or a specific element
  document.body.appendChild(dataList);
  function addOption(option: string) {
    const optionElement = document.createElement('option');
    optionElement.value = option;
    dataList?.appendChild(optionElement);
  }
  return dataList
};

/**
 * Filters the Excel table based on a criteria
 * @param {[[number, string[]]]} criteria - the first element is the column index, the second element is the values[] based on which the column will be filtered
 */
async function filterTable(tableName: string, criteria?: { column: number, value: string[] }[], clearFilter: boolean = false) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItem(tableName);

    if (clearFilter) table.autoFilter.clearCriteria();

    if (!criteria) return await getVisible();

    criteria.forEach(column => filterColumn(column.column, column.value));

    return await getVisible();

    function filterColumn(index: number, filter: string[]) {
      if (!index || !filter) return;
      table.columns.getItemAt(index).filter.applyValuesFilter(filter)
    }

    async function getVisible() {
      const visible = table.getDataBodyRange().getVisibleView();
      visible.load('values');
      await context.sync();
      return visible.values
    }
  });
}

/**
 * Converts the ',' separated text in the input into an array
 * @param value 
 * @returns {string[]}
 */
function getArray(value: string | undefined): string[] {
  if (!value) return [];
  const array =
    value.replaceAll(', ', ',')
      .replaceAll(' ,', ',')
      .split(',');
  return array.filter((el) => el);
}

async function _generateInvoice(tableName: string, templatePath: string, saveTo: string) {
  const inputs = Array.from(document.getElementsByName('input')) as HTMLInputElement[];
  if (!inputs) return;
  const lang = inputs.find(input => input.type === 'checkbox' && input.checked === true)?.id.slice(0, 3).toUpperCase() || 'FR';

  const discount = parseInt(inputs.find(input => input.id = 'discount')?.value || '0%');

  const visible = await filterTable(tableName, undefined, false);
  const date = new Date()
  const invoiceDetails = {
    number: getInvoiceNumber(new Date()),
    clientName: visible.map(row => String(row[0]))[0] || 'CLIENT',
    matters: (getUniqueValues(1, visible)).map(el => String(el)),
    adress: (getUniqueValues(15, visible)).map(el => String(el)),
    lang: lang
  };

  const savePath = `${saveTo}/${getInvoiceFileName(invoiceDetails.clientName, invoiceDetails.matters, invoiceDetails.number)}`
  const lf = new LawFirm();
  const {wordRows, totalsLabels} = lf.getRowsData(visible, discount, lang, getInvoiceNumber(date));
  const graph = new GraphAPI('');
    await graph.createAndUploadDocumentFromTemplate(templatePath, savePath, lang, [['Invoice', wordRows, 1]], lf.getContentControlsValues(invoiceDetails, new Date()));

}

function getUniqueValues(index: number, array: any[][]): any[] {
  if (!array) array = [];
  return Array.from(new Set(array.map(row => row[index])))
    .map(el => el)//we remove empty strings/values
};


class GraphAPI {
  private GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0/me/drive/root:/";
  private accessToken: string;
  private sessionId: string;
  private filePath: string;
  private methods = {
    post: 'POST',
    put: 'PUT',
    get: 'GET',
    patch: 'PATCH',
    delete: 'DELETE'
  }

  constructor(accessToken: string = '', filePath?: string, sessionId?: string, presist: Boolean = false) {
    this.accessToken = accessToken;
    this.filePath = filePath || '';
    this.sessionId = sessionId || '';
  }

  async getAccessToken() {
    const clientId = "157dd297-447d-4592-b2d3-76b643b97132";
    const redirectUri = "https://mbibawi.github.io/ExcelInvoicingAddIn"; //!must be the same domain as the app
    const msalConfig: Object = {
        auth: {
            clientId: clientId,
            authority: "https://login.microsoftonline.com/common",
            redirectUri: redirectUri,
        },
        cache: {
            cacheLocation: "localStorage",
            storeAuthStateInCookie: true
        }
    };
    return await new MSAL(clientId, redirectUri, msalConfig).getTokenWithMSAL()
}

  /**
 * Creates a new Graph API File session and returns its id
 * @returns 
 */
  async createFileSession(persist: Boolean = false) {
    const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/createSession`;
    const body = { persistChanges: persist };

    const response = await this.sendRequest(endPoint, this.methods.post, body, undefined, undefined, "Erro: Failed to create workbook session");

    const session = await response?.json();
    return session.id as string;
  }

  /**
   * Closes the current Excel file session
   */
  async closeFileSession(sessionId: string) {
    if (!this.filePath) return;
    const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/closeSession`;
    const resp = await this.sendRequest(endPoint, this.methods.post, undefined, sessionId, undefined, 'Error closing the session');
    if (resp) console.log(`The session was closed successfully! ${await resp.text()}`)
  }

  /**
   * Returns all the rows of an Excel table in a workbook stored on OneDrive, using the Graph API
   * @param {string} tableName - Name of the table to be fetched
   * @param {boolean} headers - Its default value is true. If true, it calls the "/range" endpoint and returns the whole table including the headers row, otherwise, it calls the "/rows" endpoint and returns only the body (the rows) of the table. The structure of the date returned is different for each endpoint 
   * @param {boolean} columns - If true it will return the columns
   * @returns {any[][] | number | void} - All the rows (including the title) of the Excel table
   */
  async fetchExcelTable(tableName: string, headers: boolean = true, columns?: boolean): Promise<any[][] | void> {

    let endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/`;

    if (headers) endPoint += 'range';//The "range" endpoint returns all the table including the headers row
    if (!headers) endPoint += 'rows';//The "rows" endpoint returns only  the body of the table without the headers
    else if (columns) endPoint += 'columns';
    const response = await this.sendRequest(endPoint, this.methods.get, undefined, undefined, undefined, `Error fetching row count`);

    const data = await response?.json();
    if (headers)
      return data.values as any[][];
    else return data.value.flatMap((row: any) => row.values) as any[][]//! the graph api returns an object with a "value" property which is an array of rows, each row is also an object with a "values" property which is an array of the cells values of the row. So we need to flatMap() the data to return an array of rows, each row being an array of cells values;
  };

  /**
   * Filters an Excel table column based on the values
   * @param {string} filePath - the full path and file name of the Excel workbook
   * @param {string} tableName - the name of the table that will be filtered
   * @param {string} columnName - the name of the column that will be filtered
   * @param {string[]|number[]|boolean[]} values - the values based on which the column will be filtered
   * @param {string} sessionId - the id of the current Excel file session
   * @returns {string} 
   */
  async filterExcelTable(tableName: string, columnName: string, values: string[] | number[] | boolean[], sessionId: string = this.sessionId, onValues: boolean = true) {
    if (!columnName || !values?.length || !this.filePath) return;

    // Step 3: Apply filter using the column name
    const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/columns/${columnName}/filter/apply`;

    let body;
    if (onValues)
      body = {
        "criteria": {
          "filterOn": "values",
          "values": values,
        }
      };
    else body = {
      "criteria": {
        "filterOn": "custom",
        "criterion1": values[0],
        "criterion2": values[1] || null,
        "operator": "And",
      }
    }
    const response = await this.sendRequest(endPoint, this.methods.post, body, sessionId, undefined, 'Error while applying filter to the Excel table');

    if (response) console.log(`Filter successfully applied to column ${columnName}!`);

  };

  /**
   * Returns the visible cells of a filtered Excel table using Graph API
   * @param {string} tableName - the name of the table that will be filtered
   * @param {string} sessionId - the id of the current Excel file session
   * @returns {any[][]} - the visible cells of the filtered table
   */
  async getVisibleCells(tableName: string, sessionId?: string) {
    if (!tableName || !this.filePath) return alert('Either the tableName or the filePath are mission or not valid');

    // Step 3: Apply filter using the column name
    const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/range/visibleView`;
    const response = await this.sendRequest(endPoint, this.methods.get, undefined, sessionId, undefined, "Error applying filter");
    const data = await response?.json();
    return data.values as any[][];

  };

  /**
   * Clears the filters on an Excel table using the Graph API
   * @param {string} tableName - the name of the table that will be filtered
   * @param {string} sessionId - the id of the current Excel file session
   */
  async clearFilterExcelTable(tableName: string, sessionId?: string) {
    const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/clearFilters`
    await this.sendRequest(endPoint, this.methods.post, undefined, sessionId, undefined, "Erro: Failed to clear the Excel table filter");
  };

  /**
 * Adds a new row to the Excel table using the Grap API
 * @param {string} row - The row that will be added to the Excel table
 * @param {number} index - The index at which the row will be added
 * @param {string} tableName - The name of the Excel table
 * @param {string[]} tableTitles - The titles row of the Excel table. If provided, the table will be filtered after adding the new row is added
 * @returns 
 */
  async addRowToExcelTable(row: any[], index: number | null, tableName: string, tableTitles?:string[]) {
    if (!this.filePath || !tableName || !row?.length) return alert('The filePath or the tableName argument is missing or not valid');
    const sessionId = this.sessionId || await this.createFileSession(true);
    if (!sessionId) return alert('The sessionId is missing Check the console.log for more details');
    const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/rows`;//The url to add a row to the table

    const body = {
      index: index,
      values: [row],
    };

    await this.clearFilterExcelTable(tableName, sessionId);//We clear the filtering of the table

    const resp = await this.sendRequest(endPoint, this.methods.post, body, sessionId, undefined, "Error adding row to the Excel Table");

    if (resp) console.log("Row added successfully!");

    for (const index of [0, 1]) {
      if (!tableTitles?.length) break;
      await this.filterExcelTable(tableName, tableTitles[index], [row[index].toString()], sessionId);  //!We use "for of" loop because forEach doesn't await
    }

    await this.sortExcelTable(tableName, [[3, true]], false, sessionId);//We sort the table by the first column (the date column)
    const visible = await this.getVisibleCells(tableName, sessionId);
    return visible;
  };

  /**
   * Updates a specific row in an Excel table using the file path.
   * @param {string} workbookPath - 
   * @param {string} tableName - The name of the table.
   * @param {number} rowIndex - 0-based index of the row.
   * @param {Array} values - 1D array of values for the row.
   */
  async updateExcelTableRow(tableName: string, rowIndex: number, values: any[]) {
    if (!this.filePath || !tableName || !rowIndex || !values?.length) return alert('One of the arguments is missing or not valid');
    const sessionId = await this.createFileSession();
    if (!sessionId) return alert('Failed to create a new Session');
    const url = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/rows/itemAt(index=${rowIndex})`;

    const body = {
      values: [values] // API requires a 2D array
    };

    try {
      const response = await this.sendRequest(url, this.methods.patch, body, sessionId, undefined, "Error while updating the Excel Table row's values");
      if (response?.ok) alert('Successfully updated the Excel table');
      const data = await response?.json();
      console.log('Update Successful:', data);
      return data;
    } catch (err) {
      console.error('Update Failed:', err);
    }
  };

  /**
 * Creates an invoice Word document from the invoice Word template, then uploads it to the destination folder
 * @param {string} templatePath - The full path of the Word invoice template
 * @param {string} savePath - The full path of the destination folder where the new invoice will be saved
 * @param {string} lang - The language in which the invoice will be issued
 * @param {[string, string[][], number][]} tables - An array containing for each table: the title of the Word table, the data for the new Word table rows that will be added to the table, and the index of the row after which we will insert the new rows.
 * @param {[string, string][]} contentControls - The titles and text of each of the content controls that will be updated in the Word document. Each element of the array contains the title of the contentControl, and the text with which it will be filled.
 * @param {string[]} totalsLabels - The labels of the rows that will be formatted as totals
 * @returns 
 */
  async createAndUploadDocumentFromTemplate(templatePath: string, savePath: string = this.filePath, lang: string, tables?: [string, string[][], number][], contentControls?: [string, string][], totalsLabels?: string[]) {
    if (!templatePath || !this.filePath) return;

    const blob = await this.fetchFileFromOneDrive(templatePath);//!We must provide the Word templatePath not the Excel workbook path stored in the this.filePath variable

    if (!blob) return;

    const [xmlDocs, zip] = await this.convertBlobIntoXML(blob);

    if (!xmlDocs.length) return;

    xmlDocs.forEach(doc => editXML(doc));

    const newblob = await this.convertXMLIntoBlob(xmlDocs, zip);
    if (!newblob) return;
    await this.uploadFileToOneDrive(newblob, savePath);

    function editXML([doc, fileName]: [XMLDocument, string]) {
      const xml = new XML(doc, lang), schema = xml.schema();
      editTables(xml, doc, schema);
      editContentControls(xml, doc, schema);
    };

    function editTables(xml: XML, doc: XMLDocument, schema: string) {
      if (!tables) return;
      const allTables = xml.getTables(doc);
      tables.forEach(table => editTable(table));

      function editTable([tableTitle, rows, index]: [string, string[][], number]) {
        const table = xml.findTableByTitle(allTables, tableTitle);
        if (!table) return;
        const afterRow = xml.getTableRow(table, index);//We retrieve the table row below which we will insert the new rows
        rows.forEach((row, index) => {
          const newXmlRow = xml.insertRowAfter(table, afterRow, NaN, true) || table.appendChild(xml.createTableRow());
          if (!newXmlRow) return;
          const isTotal = totalsLabels?.includes(row[0]);
          const isLast = index === rows.length - 1;
          return editCells(newXmlRow, row, isLast, isTotal);
        });
        afterRow.remove();//We remove the first row when we finish
      }

      function editCells(tableRow: Element, values: string[], isLast: boolean = false, isTotal: boolean = false) {
        const cells = xml.getRowCells(tableRow) || values.map(v => tableRow.appendChild(xml.createTableCell()));//getting all the cells in the row element

        cells.forEach((cell, index) => {
          const textElement = xml.getTextElement(cell, 0) || xml.appendParagraph(cell);
          if (!textElement) return console.log('No text element was found !');
          const pPr = xml.setTextLanguage(cell);//We call this here in order to set the language for all the cells. It returns the pPr element if any.
          textElement.textContent = values[index];

          (function totalsRowsFormatting() {
            if (!isLast && !isTotal) return;
            (function cellBackgroundColor() {
              const tcPr = xml.getPropElement(cell, 0) || cell.prepend(xml.createPropElement(cell));
              const shadow = xml.getShadowElement(tcPr, 0) || tcPr.appendChild(xml.createShadowElement());//Adding background color to cell
              shadow.setAttributeNS(schema, 'val', "clear");
              shadow.setAttributeNS(schema, 'fill', 'D9D9D9');
            })();

            (function paragraphStyle() {
              if (!pPr) return console.log('No "w:pPr" or "w:rPr" property element was found !');
              const style = xml.getParagraphStyle(pPr, 0) || pPr.appendChild(xml.createParagraphStyle());
              style.setAttributeNS(schema, 'val', xml.getStyle(index, isTotal && !isLast));
            })();
          })();
        })
      }
    };

    function editContentControls(xml: XML, doc: XMLDocument, schema: string) {
      if (!contentControls?.length) return;
      const ctrls = xml.getContentControls(doc);
      contentControls
        .forEach(([title, value]) => {
          const sameTitle = xml.findContentControlsByTitle(ctrls, title) as Element[];//!we  retrieve all then XML ContentControls having the same title
          sameTitle.forEach(control => xml.editContentControlText(control, value))
        });
    };
  };

  /**
 * Converts the blob of a Word document into XML files: the XML for the document and the XMLs for the header and footer
 * @param blob - the blob of the file to be converted
 * @returns {[XMLDocument, JSZip]} - The xml document, and the zip containing all the xml files
 */
  private async convertBlobIntoXML(blob: Blob): Promise<[[XMLDocument, string][], JSZip]> {

    const arrayBuffer = await blob.arrayBuffer();

    const zip = new JSZip();
    await zip.loadAsync(arrayBuffer);
    const zipFiles = Object.keys(zip.files);

    const parser = new DOMParser();
    const xmlFiles: [XMLDocument, string][] = [];

    const Patterns = [
      /^word\/document\.xml$/,
      /^word\/header\d+\.xml$/,
      /^word\/footer\d+\.xml$/
    ];

    const fileNames = zipFiles.filter(file => Patterns.find(pattern => pattern.test(file)));

    for (const fileName of fileNames) {
      const file = await getXmlFromZip(fileName);
      if (file) xmlFiles.push([file, fileName.replace(/^word\//, '')])
    }

    return [xmlFiles, zip];

    async function getXmlFromZip(fileName: string): Promise<XMLDocument | null> {
      const content = await zip.file(fileName)?.async("string");
      if (!content) return null
      return parser.parseFromString(content, "application/xml");
    }
  }

  /**
 * Filters an Excel table column based on the values
 * @param {string} tableName - the name of the table that will be filtered
 * @param {[string, boolean][]} columns - each element contains the name of the column and whether it will be sorted ascending or descending
 * @param {string} sessionId - the id of the current Excel file session
 * @returns {string} 
 */
  private async sortExcelTable(tableName: string, columns: [number, boolean][], matchCase: boolean, sessionId?: string) {
    if (!this.filePath) return;

    // Step 3: Apply filter using the column name
    const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/sort/apply`;

    const fields = columns.map(([index, ascending]) => {
      return {
        key: index,
        ascending: ascending,
        sortOn: "value"
      }
    });

    const body = {
      fields: fields,
      "matchCase": matchCase
    }

    const resp = await this.sendRequest(endPoint, this.methods.post, body, sessionId, undefined, "Error sorting table");

    if (resp)
      console.log(`Table successfully sorted according to columns criteria: ${columns.map(([col, asc]) => col).join(' & ')}!`);

  };

  /**
 * Returns a blob from a file stored on OneDrive, using the Graph API and the file path
 * @param {string} filePath 
 * @returns {Blob} - A blob of the fetched file, if successful
 */
  private async fetchFileFromOneDrive(filePath = this.filePath): Promise<Blob | undefined> {
    const endPoint = `${this.GRAPH_API_BASE_URL}${filePath}:/content`;
    const response = await this.sendRequest(endPoint, this.methods.get, undefined, undefined, undefined, "Failed to fetch Word template");
    return await response?.blob(); // Returns the Word template as a Blob
  };

  /**
 * Uploads a file blob to OneDrive using the Graph API
 * @param {Blob } blob 
 * @param {string} filePath 
 */
  private async uploadFileToOneDrive(blob: Blob, filePath: string) {
    if (!filePath) return;
    const endpoint = `${this.GRAPH_API_BASE_URL}${filePath}:/content`
    const contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

    const response = await this.sendRequest(endpoint, this.methods.put, blob, undefined, contentType, 'Failed to upload the Word file to OneDrive')

    if (response?.ok)
      alert('succefully uploaded the new file');
    else console.log('failed to upload the file to onedrive error = ', await response?.json())
  };

  /**
   * Adds files (or replaces the existing files if the fileName is the same) in the zip folder
   * @param {XMLDocument, string} docs - the XMLDocument we want to add to the zip folder, and its fileName
   * @param {JSZip} zip -the zip folder into which we want to add the XMLDocuments 
   * @returns {blob} a blob of the zip folder (which is in our case a Word document)
   */
  private async convertXMLIntoBlob(docs: [XMLDocument, string][], zip: JSZip,) {
    const serializer = new XMLSerializer();
    docs.forEach(doc => serialize(doc))
    return await zip.generateAsync({ type: "blob" });

    function serialize([doc, fileName]: [XMLDocument, string]) {
      const serialized = serializer.serializeToString(doc);
      zip.file(`word/${fileName}`, serialized);
    }
  }

  /**
   * 
   * @param {string} endPoint 
   * @param {string} method 
   * @param {object} body 
   * @param {string} sessionId 
   * @param {string} contentType 
   * @param {string} message 
   * @returns {Promise<Response | undefined>} 
   */
  async sendRequest(endPoint: string, method: string, body?: object|Blob|ArrayBuffer, sessionId?: string, contentType?: string, message = "") {
    if (!this.accessToken) this.accessToken = await this.getAccessToken() || '';
    if (!this.accessToken) return alert('Could not get an accessToken');
    const request: RequestInit = {
      method: method,
      headers: this.graphHeaders(sessionId, contentType)
    };

    if (body) {
      if (body instanceof Blob || body instanceof ArrayBuffer) {
        request.body = body; // Send raw binary
      } else {
        request.body = JSON.stringify(body); // Send JSON
      }
    }

    const response = await fetch(endPoint, request);

    if (response?.ok) return response;

    message = `${message || `Error while sending ${method} request`}:\n ${await response?.text()}`;
    if (sessionId) await this.closeFileSession(sessionId)
    throwAndAlert(message)
  };

  /**
 * Returns the headers of the Microsoft Graph API calls
 */
  private graphHeaders(sessionId?: string, contentType?: string) {
    const headers: header = {
      'Authorization': `Bearer ${this.accessToken}`,
      'Content-Type': contentType || 'application/json',
    };
    if (sessionId) headers["workbook-session-id"] = sessionId;
    return headers
  }

  private async getExcelTableRowsCount(filePath: string, tableName: string) {
    const endPoint = `${this.GRAPH_API_BASE_URL}${filePath}:/workbook/tables/${tableName}/rows/$count`;
    const response = await this.sendRequest(endPoint, this.methods.get)
  
    if (response?.ok) {
      const rowCount = await response?.text(); // The API returns a number as plain text
      console.log(`Row count: ${rowCount}`);
      return parseInt(rowCount, 10); // Convert to number
    } else {
      console.error("Error fetching row count:", await response?.text());
      return null;
    }
  }
};

class XML {
  private doc;
  private lang;
  private tags = {
    ctrl: "sdt",
    ctrlContent: 'sdtContent',
    table: "tbl",
    row: "tr",
    cell: "tc",
    text: 't',
    shadow: 'shd',
    style: 'pStyle',
    paragraph: 'p',
    run: 'r',
    alias: 'alias',
    lang: 'lang',
    tableCaption: 'tblCaption',
  };

  constructor(doc: XMLDocument, lang: string) {
    this.doc = doc;
    this.lang = lang;
  }

  schema = () => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  Pr = (tag: string) => `${tag.replace('w:', '')}Pr`;//!we need to remove the "w:" prefix from the tag

  /**
   * Returns all the XML ContentConrol ("sdt") elements of the parent XML Element passed as argument
   * @param {XMLDocument | Element} parent - the parent XML Document or Element that we want to retrieve all its nested XML ContentControl elements
   * @returns {Element[]}
   */
  getContentControls(parent: XMLDocument | Element = this.doc) {
    return this.getXMLElements(parent, this.tags.ctrl) as Element[]
  }

  /**
   * Returns a XML ContentControl element ("sdt") nested in the parent XML Element passed as argument
   * @param {Element} parent - the parent XML element of the XML ContentControl element we want to retrieve
   * @param {number} index - the index of the XML ContentControl element we want to retrieve
   * @returns {Element}
   */
  getControlContent(parent: Element, index: number) {
    return this.getXMLElements(parent, this.tags.ctrlContent, index) as Element
  }

  /**
   * Returns all the XML table ("tbl") elements nested in the XML Document or the XML Element passed as argument
   * @param {XMLDocument | Element} parent - the parent XML document or Element for which we want to retrieve the XML tables
   * @returns {Elemnt[]}
   */
  getTables(parent: XMLDocument | Element = this.doc) {
    return this.getXMLElements(parent, this.tags.table) as Element[]
  }

  /**
   * Returns an XML table row ("tr") element nested in the XML table element passed as argument
   * @param {Element} table - the table element for which we want to retrieve a specifc XML row element by its index
   * @param {number} index - the index of the row
   * @returns {Element}
   */
  getTableRow(table: Element, index: number) {
    return this.getXMLElements(table, this.tags.row, index) as Element
  }

  /**
 * Returns the XML '[tag]Pr' element child of any element
 * @param {Element} parent - the parent XML Element for which we are trying to retrieve the "[tag]Pr" element
 * @param {number} index - the index of the "[tag]Pr" element
 * @returns {Element}
 */
  getPropElement(parent: Element, index: number = 0) {
    const tag = this.Pr(parent.tagName.toLowerCase())
    return this.getXMLElements(parent, tag, index) as Element;
  }

  /**
   * Creates and returns a table row ("tr") XML element
   * @returns {Element}
   */
  createTableRow() {
    return this.createXMLElement(this.tags.row)
  }

  /**
   * Creates and returns a table XML cell ("tc") element
   * @returns {Element}
   */
  createTableCell() {
    return this.createXMLElement(this.tags.cell)
  }

  /**
   * Creates and returns a "[tag]Pr" XML element which is a child element holding the properties of the parent XML Element
   * @param {Element} parent - the parent XML Element from which we will retrieve the tag of XML "[tag]Pr" element that we will create
   * @returns {Element}
   */
  createPropElement(parent: Element) {
    const tag = this.Pr(parent.tagName.toLowerCase())
    return this.createXMLElement(tag)
  }

  /**
   * Returns the XML cell elements of row in the table
   * @param {Element} tableRow - the XML row element that we want to retrieve its XML cell children elements
   * @returns {Element[]}
   */
  getRowCells(tableRow: Element) {
    return this.getXMLElements(tableRow, this.tags.cell) as Element[];
  }

  /**
   * Returns a text XML element of the parent according to its index
   * @param {Element} parent - the XML element that we want to retrieve one of its text ("t") XML children
   * @param {number} index - the index of the text ("t") XML element we want to retrieve
   * @returns {Element}
   */
  getTextElement(parent: Element, index: number) {
    return this.getXMLElements(parent, this.tags.text, index) as Element
  }

  /**
   * Creates and returns a pargraph style ("pStyle") XML element
   * @returns {Element}
   */
  createParagraphStyle() {
    return this.createXMLElement(this.tags.style)
  }

  /**
   * Returns a XML shadow element ("shd") of the parent according to its index passed as argument
   * @param {Element} parent - 
   * @param {number} index - 
   * @returns {Element}
   */
  getShadowElement(parent: Element, index: number) {
    return this.getXMLElements(parent, this.tags.shadow, index) as Element
  }

  /**
   * Creates and returns a XML shadow element ("shd")
   * @returns {Element} - an XML shadow ("shd") element
   */
  createShadowElement() {
    return this.createXMLElement(this.tags.shadow) as Element
  }

  /**
    * Looks for a child "w:p" (paragraph) element, if it doesn't find any, it looks for a "w:r" (run) element.
    * @param {Element} parent - the parent XML of the paragraph or run element we want to retrieve. 
    * @returns {Element | undefined} - an XML element representing a "w:p" (paragraph) or, if not found, a "w:r" (run), or undefined
    */
  getParagraphOrRun(parent: Element) {
    return this.getXMLElements(parent, this.tags.paragraph, 0) as Element || this.getXMLElements(parent, this.tags.run, 0) as Element;
  }

  /**
   * Returns the cells of row in the table
   * @param {Element} tableRow
   */
  getParagraphStyle(parent: Element, index: number) {
    return this.getXMLElements(parent, this.tags.style, index) as Element;
  }

  insertRowAfter(table: Element, rowTemplate: Element, after: number = -1, clone: boolean = false) {
    const this$ = this;

    if (clone) return cloneAndAppend();
    else return create();

    function create() {
      if (!table) return;
      const row = this$.createTableRow();
      after >= 0 ? (this$.getXMLElements(table, this$.tags.row, after) as Element)?.insertAdjacentElement('afterend', row) :
        table.appendChild(row);
      return row;
    }

    function cloneAndAppend() {
      if (!rowTemplate) return;
      const row = rowTemplate.cloneNode(true) as Element;
      table?.appendChild(row);
      return row
    };
  }

  getStyle(cell: number, isTotal: boolean = false) {
    let style = 'Invoice';
    if (cell === 0 && isTotal) style += 'BoldItalicLeft';
    else if (cell === 0) style += 'BoldLeft';
    else if (cell === 1) style += 'NotBoldItalicLeft';
    else if (cell === 2 && isTotal) style += 'BoldItalicCentered';
    else if (cell === 2) style += 'BoldCentered';
    else if (cell === 3) style += 'BoldItalicCentered';
    else style = '';
    return style
  }

  /**
   * 
   * @param {Element[]} ctrls - the XML ContentControls array from which we will retrieve an XML ContentControl by its title
   * @param {string} title - the title of the XML ContentControl we want to retrieve
   * @param {number} index - if omitted, the function will return a collection of all the XML ContentControl elements having the same title. Otherwise it will return a ContentControl by its index
   * @returns {Element | Element[] | undefined}
   */
  findContentControlsByTitle(ctrls: Element[], title: string): Element[] {
    return this.findElementsByPropertyValue(ctrls, this.tags.alias, title)
  }

  /**
   * Finds and returns a XML Table by its title ("tblCaption")
   * @param {Element[]} tables - the XML tables array in which we will search for a table having the specified title
   * @param {string} title - the title of the table
   * @returns {Element | undefined}
   */
  findTableByTitle(tables: Element[], title: string) {
    return this.findElementsByPropertyValue(tables, this.tags.tableCaption, title)?.[0]
  }
  /**
   * 
   * @param {Element[]} elements - the XML Elements collection in which we will search for specific XML elemnt(s) by the value of a given property
   * @param {string} tag - the name of property in which the title of the XML Element title is stored
   * @param {string} value - the value of the property we are looking for
   * @returns {Element []}
   */
  private findElementsByPropertyValue(elements: Element[], tag: string, value: string): Element[] {
    if (!tag || !value) return [];
    const children = (parent: Element) => this.getXMLElements(parent, tag) as Element[];//This returns the child elements of the parent (if any) having the specified tag. The children hold a property of the element
    return elements.filter(element => children(element)?.find(child => child.getAttributeNS(this.schema(), 'val') === value));
  }

  /**
* Adds a new paragraph XML element or appends a cloned paragraph, and in both cases, it returns the textElement of the paragraph
* @param {Element} element - The element to which the new paragraph will be appended if the parent argument is not provided. If the parent argument is provided, the element will be cloned assuming that this is a pargraph element
* @param {Elemenet} parent - If provided, element will be cloned and appended to parent.
* @returns {Element} the textElemenet attached to the paragraph
*/
  appendParagraph(element: Element, parent?: Element) {
    const this$ = this;
    if (parent) return clone();
    else return create();
    function clone() {
      const parag = element?.cloneNode(true) as Element;
      parent?.appendChild(parag);
      return this$.getXMLElements(parag, 't', 0) as Element
    }
    function create() {
      const parag = element.appendChild(this$.createXMLElement(this$.tags.paragraph));
      parag.appendChild(this$.createXMLElement(this$.Pr(this$.tags.paragraph)));
      const run = parag.appendChild(this$.createXMLElement(this$.tags.run));
      return run.appendChild(this$.createXMLElement(this$.tags.text));
    }
  }

  private createXMLElement(tag: string) {
    return this.doc.createElementNS(this.schema(), tag);
  }
  /**
   * Returns the XML element(s) nested under the XML parent element by its/their tag; 
   * @param {XMLDocument | Element} parent - the parent XML Document or XML Element nesting the XML Element(s) we want to retrieve
   * @param {string} tag - the tag of the XML Element(s) we want to retrieve
   * @param {number} index - if provided, the function will only return the element having the specified index
   * @returns {Element[] | Element | undefined}
   */
  private getXMLElements(parent: XMLDocument | Element, tag: string, index: number = NaN): Element[] | Element | void {
    const elements = parent?.getElementsByTagNameNS(this.schema(), tag);
    if (!isNaN(index)) return elements[index];
    return Array.from(elements)
  }

  editContentControlText(control: Element, text: string | null) {
    if (text === "DELETECONTENTECONTROL") return control.remove();
    if (!text) text = 'NO VALUE WAS PROVIDED';

    const sdtContent = this.getControlContent(control, 0);
    if (!sdtContent) return;
    const paragTemplate = this.getParagraphOrRun(sdtContent);//This will set the language for the paragraph or the run
    if (!paragTemplate) return console.log('No template paragraph or run were found !');
    this.setTextLanguage(paragTemplate);//We amend the language element to the "w:pPr" or "r:pPr" child elements of paragTemplate
    const this$ = this;
    text?.split('\n')
      .forEach((parag, index) => editParagraph(parag, index));

    function editParagraph(parag: string, index: number) {
      let textElement: Element;
      if (index < 1)
        textElement = this$.getXMLElements(paragTemplate, this$.tags.text, index) as Element;
      else textElement = this$.appendParagraph(paragTemplate, sdtContent);//We pass sdtContent as parent argumself
      if (!textElement) return console.log('No textElement was found !');

      textElement.textContent = parag;

    }
  }

  /**
 * Finds a "w:pPr" XML element (property element) which is a child of the XML parent element passed as argument. If does not find it, it looks for a "w:rPr" XML element. When it finds either a "w:pPr" or a "w:rPr" element, it appends a "w:lang" element to it, and sets its "w:val" attribute to the language passed as "lang"
 * @param {Element} parent - the XML element containing the paragraph or the run for which we want to set the language.
 * @returns {Element | undefined} - the "w:pPr" or "w:rPr" property XML element child of the parent element passed as argument
 */
  setTextLanguage(parent: Element) {
    const pPr = this.getXMLElements(parent, this.Pr(this.tags.paragraph), 0) as Element ||
      this.getXMLElements(parent, this.Pr(this.tags.run), 0) as Element;
    if (!pPr) return;
    pPr
      .appendChild(this.createXMLElement(this.tags.lang))//appending a "w:lang" element
      .setAttributeNS(this.schema(), 'val', `${this.lang.toLowerCase()}-${this.lang.toUpperCase()}`);//setting the "w:val" attribute of "w:lang" to the appropriate language like "fr-FR"
    return pPr as Element
  }

}

class officeJS {
  private GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0/me/drive/root:/";

  async editDocumentWordJSAPI(id: string, accessToken: string= '', data: string[][], controlsData: string[][]) {
    if (!id || !data) return;
    const graph = new GraphAPI(accessToken);

    await Word.run(async (context) => {
      // Open the document by downloading its content
      const graph = new GraphAPI('');
      const endPoint = `${this.GRAPH_API_BASE_URL}/items/${id}/content`;
      //@ts-ignore
      const fileResponse = await graph.fetchExcelTable('', true) as Response;
    

      if (!fileResponse?.ok)
        throw new Error("Failed to retrieve document");

      const blob = await fileResponse?.blob();
      const base64File = await convertBlobToBase64(blob);

      const doc = context.application.createDocument(base64File);
      console.log("Word document opened for editing:", document);

      const tables = doc.body.tables;
      const contentControls = doc.body.contentControls;
      context.load(tables);
      context.load(contentControls);
      await context.sync();

      const table = tables.items[0];
      if (!table) return;

      data.forEach(dataRow => table.addRows("End", 1, [dataRow]));

      await editRichTextContentControls();

      async function editRichTextContentControls() {
        if (!controlsData || contentControls) return;

        controlsData.forEach(control => edit(control));

        async function edit(control: string[]) {
          const [title, text] = control;
          const field = contentControls.getByTitle(title).getFirst();
          if (!field) return;
          context.load(field);
          await context.sync();
          if (!text) field.delete(false);
          else field.insertText(text, 'Replace');
          await context.sync();
          return field
        }
      }

    });
  }
  async addEntry(tableName: string, rows?: any[][]) {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const table = sheet.tables.getItem(tableName);
      const columns = table.columns.getCount();
      await context.sync();
      table.rows.add(-1, getNewRow(columns.value), true);
      table.getDataBodyRange().load('rowCount');
      await context.sync();

      [5, 6].forEach(i => {
        const cell = table.getRange()
          .getCell(table.getDataBodyRange().rowCount - 1, i);
        console.log('cell = ', cell);
        cell.numberFormatLocal = [["hh:mm:ss"]];
      });

      await context.sync();
    });

    function getNewRow(columns: number) {
      const newRow = Array(columns).map(el => '') as any[];
      const inputs = Array.from(document.getElementsByTagName('input')).filter(input => input.dataset.index);
      console.log('inputs = ', inputs)
      if (inputs.length < 1) return;

      inputs.forEach(input => {
        const index = Number(input.dataset.index);
        let value: string | number | Date = input.value;
        if (input.type === 'number')
          value = parseFloat(value);
        else if (input.type === 'date' && input.valueAsDate)
          value = getDateString(input.valueAsDate);
        else if (input.type === 'time' && input.valueAsDate) value = [input.valueAsDate?.getHours().toString().padStart(2, '0'), input.valueAsDate?.getMinutes().toString().padStart(2, '0'), '00'].join(':');

        newRow[index] = value;
      });

      console.log('newRow = ', newRow);
      return [newRow];

      function convertTo24HourFormat(time12h: string): string {
        const [time, modifier] = time12h.split(' ');
        let [hours, minutes] = time.split(':');

        if (hours === '12') hours = '00';

        if (modifier === 'PM') hours = String(parseInt(hours, 10) + 12);

        return `${hours}:${minutes}:00`;
      }
    }
  }
};

class blob {
  // Utility function: Convert Blob to Base64
  async blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result!.toString().split(",")[1]);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  }

  // Utility function: Convert Base64 to Blob
  base64ToBlob(base64: string): Blob {
    const byteCharacters = atob(base64);
    const byteNumbers = new Array(byteCharacters.length).fill(0).map((_, i) => byteCharacters.charCodeAt(i));
    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }
};

/**
 * Returns the Word file name by which the newly issued invoice will be saved on OneDrive
 * @param {string} clientName - The name of the client for which the invoice will be issued
 * @param {string} matters - The matters included in the invoice
 * @param {string} invoiceNumber - The invoice serial number
 * @returns {string} - The name of the Word file to be saved
 */
function getInvoiceFileName(clientName: string, matters: string[], invoiceNumber: string): string {
  // return 'test file name for now.docx'
  return `${clientName}_Facture_${Array.from(matters).join('&')}_No.${invoiceNumber.replace('/', '@')}.docx`
    .replaceAll('/', '_')
    .replaceAll('"', '')
    .replaceAll("\\", '');
};

function getInvoiceNumber(date: Date): string {
  const padStart = (n: number) => n.toString().padStart(2, '0');

  return `${date.getFullYear() - 2000}${padStart(date.getMonth() + 1)}${padStart(date.getDate())}/${padStart(date.getHours())}${padStart(date.getMinutes())}`;
};

/**
 * Returns any date in the ISO format (YYY-MM-DD) accepted by Excel
 * @param {Date} date - the Date that we need to convert to ISO format
 * @returns {string} - The date in ISO format
 */
function getISODate(date: Date | null) {
  if (!date) return '';
  return [date.getFullYear(), date.getMonth() + 1, date.getDate()].map(el => el.toString().padStart(2, '0')).join('-');
};

/**
 * Returns the date in a string formated like: "DD/MM/YYYY"
 */
function getDateString(date: Date | null) {
  if (!date) return ''
  return [date.getDate(), date.getMonth() + 1, date.getFullYear()]
    .map(el => el.toString().padStart(2, '0'))
    .join('/');
};

/**
 * Returns the value from a time input as a number matching the Excel time format (which is a fraction of the day)
 * @param {HTMLInputElement[]} inputs - If a single input is passed, it will return the Excel formatted time value from this input or 0. If 2 inputs are passed, it will return the total time by calculting the difference between the second input and the first input in the array
 * @returns {number} - The time as a number matching the Excel time format
 */
function getTime(inputs: (HTMLInputElement | undefined)[]) {
  const day = (1000 * 60 * 60 * 24);

  if (inputs.length < 2 && inputs[0])
    return inputs[0].valueAsNumber / day || 0;

  const from = inputs[0]?.valueAsNumber;//this gives the time in milliseconds
  const to = inputs[1]?.valueAsNumber;

  if (!from || !to) return 0;

  const quarter = 15 * 60 * 1000; //quarter of an hour
  let time = to - from;
  time = Math.round(time / quarter) * quarter;//We are rounding the time by 1/4 hours
  time = time / day;
  if (time < 0) time = (to + day - from) / day//It means we started on one day and finished the next day 
  return time;
};


/**
 * Helper function to convert Blob to Base64
 *  */
function convertBlobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result!.toString().split(",")[1]); // Extract base64 part
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
};

class MSAL {
  private msalInstance;
  private clientId: string = '';
  private redirectUri: string = '';
  private loginRequest = { scopes: [''] };

  constructor(clientId: string, redirectUri: string, msalConfig: Object, scopes: string[] = ["Files.ReadWrite"]) {
    this.clientId = clientId;
    this.redirectUri = redirectUri;
    this.loginRequest.scopes = scopes;
    //@ts-expect-error
    this.msalInstance = new msal.PublicClientApplication(msalConfig);
  }

  async getTokenWithMSAL() {
    if (!this.clientId || !this.redirectUri || !this.msalInstance) return;
    return await this.acquireToken();
  };
  // Function to check existing authentication context
  private async acquireToken(): Promise<string | undefined | void> {
    try {
      const account = this.msalInstance.getAllAccounts()[0];
      if (account) {
        return await this.acquireTokenSilently(account);
      } else {
        return await this.loginWithPopup();
      }
    } catch (error) {
      console.error("Failed to acquire token from acquireToken(): ", error);
    }
  }
  // Function to get access token silently
  private async acquireTokenSilently(account: any): Promise<string | undefined> {
    try {
      const tokenRequest = {
        account: account,
        scopes: this.loginRequest.scopes, // OneDrive scopes
      };

      const tokenResponse = await this.msalInstance.acquireTokenSilent(tokenRequest);
      if (tokenResponse && tokenResponse.accessToken) {
        console.log("Token acquired silently :", tokenResponse.accessToken);

        return tokenResponse.accessToken;
      }
    } catch (error) {
      console.error("Token silent acquisition error:", error);
    }
  };

  private async loginWithPopup() {
    try {
      const loginResponse = await this.msalInstance.loginPopup(this.loginRequest);
      console.log('loginResponse = ', loginResponse);

      this.msalInstance.setActiveAccount(loginResponse.account);

      const tokenResponse = await this.msalInstance.acquireTokenSilent({
        account: loginResponse.account,
        scopes: ["Files.ReadWrite"]
      });

      console.log("Token acquired from loginWithPopup: ", tokenResponse.accessToken);
      return tokenResponse.accessToken
    } catch (error) {
      console.error("Error acquiring token from loginWithPopup(): ", error);
      //@ts-ignore
      if (error instanceof InteractionRequiredAuthError) {
        // Fallback to popup if silent token acquisition fails
        const response = await this.msalInstance.acquireTokenPopup({
          scopes: ["Files.ReadWrite"]
        });
        console.log("Token acquired via popup:", response.accessToken);
        return response.accessToken
      }
    }
  }

  private async credentitalsToken(tenantId: string) {
    const msalConfig = {
      auth: {
        clientId: this.clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        //clientSecret: clientSecret,
      }
    }
    //@ts-ignore
    const cca = new msal.application.ConfidentialClientApplication(msalConfig);

    const tokenRequest = {
      scopes: ["Files.ReadWrite"],
    }

    try {
      const response = await cca.acquireTokenByClientCredential(tokenRequest);
      return response.accessToken;
    } catch (error) {
      console.log('Error acquiring Token: ', error)
      return null

    }

  }

  private async getOfficeToken() {
    try {
      //@ts-ignore
      return await OfficeRuntime.auth.getAccessToken()

    } catch (error) {
      console.log("Error : ", error)

    }

  }

  private async getTokenWithSSO(email: string, tenantId: string) {
    const msalConfig = {
      auth: {
        clientId: this.clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        redirectUri: this.redirectUri,
        navigateToLoginRequestUrl: true,
      },
      cache: {
        cacheLocation: "ExcelAddIn",
        storeAuthStateInCookie: true
      }
    };
    try {
      //@ts-ignore
      const response = await this.msalInstance.ssoSilent({
        scopes: ["Files.ReadWrite"],
        //scopes: ["https://graph.microsoft.com/.default"],
        loginHint: email // Forces MSAL to recognize the signed-in user
      });

      console.log("Token acquired via SSO:", response.accessToken);
      return response.accessToken;
    } catch (error) {
      console.error("SSO silent authentication failed:", error);
      return null;
    }
  }

  private openLoginWindow() {
    const loginUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${this.clientId}&response_type=token&redirect_uri=${this.redirectUri}&scope=https://graph.microsoft.com/.default`;

    // Open in a new window (only works if triggered by user action)
    const authWindow = window.open(loginUrl, "_blank", "width=500,height=600");

    if (!authWindow) {
      console.error("Popup blocked! Please allow popups.");
    }
  }

  // Function to handle login and acquire token
  private async loginAndGetToken(): Promise<string | undefined> {
    const msalConfig = {
      auth: {
        clientId: this.clientId,
        authority: "https://login.microsoftonline.com/common",
        redirectUri: this.redirectUri
      },

      cache: {
        cacheLocation: "ExcelInvoicing", // Specify cache location
        storeAuthStateInCookie: true  // Set this to true for IE 11
      }
    };

    return await acquire(this);
    async function acquire(this$: MSAL) {
      try {
        const response = await this$.msalInstance.handleRedirectPromise();
        if (response !== null) {
          console.log("Login successful:", response);
          return response.accessToken;
        }
        const accounts = this$.msalInstance.getAllAccounts();
        if (accounts.length > 0) {
          const tokenResponse = await this$.msalInstance.acquireTokenSilent({
            account: accounts[0],
            scopes: ["https://graph.microsoft.com/.default"]
          });
          console.log("Token acquired silently:", tokenResponse.accessToken);
          return tokenResponse.accessToken;
        }
      } catch (error) {
        console.error("Error acquiring token:", error);
        //@ts-ignore
        if (error instanceof msal.InteractionRequiredAuthError) {
          this$.msalInstance.acquireTokenRedirect({
            scopes: ["https://graph.microsoft.com/.default"]
          });
        }
      }
    }


    // Function to handle redirect response
    async function handleRedirectResponse(this$: MSAL): Promise<string | undefined> {
      try {
        const authResult = await this$.msalInstance.handleRedirectPromise();
        if (authResult && authResult.accessToken) {
          console.log("Access token:", authResult.accessToken);
          return authResult.accessToken;
        }
      } catch (error) {
        console.error("Redirect handling error:", error);
      }
      return undefined;
    }
  }


}


function sortByColumn(data: any[][], columnIndex: number): any[][] {
  return data.slice().sort((a, b) => {
    const valA = a[columnIndex];
    const valB = b[columnIndex];

    if (typeof valA === "number" && typeof valB === "number") {
      return valA - valB; // Numeric sorting
    }

    return String(valA).localeCompare(String(valB)); // String sorting
  });
};

function getInputByIndex(inputs: HTMLInputElement[], index: number) {
  return inputs.find(input => Number(input.dataset.index) === index)
}
/**
 * Returns the dataset.index value of the input as a number
 * @param {HTMLInputElement} input - the input with a dataset.index attribute
 * @returns {number} - the dataset.index value of the input as a number
 */
function getIndex(element: HTMLElement) {
  return Number(element?.dataset.index)
}

function getSavedSettings() {
  return saveSettings(undefined, true);
}

function saveSettings(values?: [string, string][], get: boolean = false) {

  const settings: settings = {
    issueInvoice: {
      workBook: {
        label: 'Invoices workbook path :',
        name: settingsNames.invoices.workBook,
        value: ''
      },
      wordTemplate: {
        label: 'Invoices\'Word template path: ',
        name: settingsNames.invoices.wordTemplate,
        value: ''
      },
      saveTo: {
        label: 'Invoices\' save to path: ',
        name: settingsNames.invoices.saveTo,
        value: ''
      },
      tableName: {
        label: 'Invoices\' Excel Table name: ',
        name: settingsNames.invoices.tableName,
        value: ''
      },
    },
    Letter: {
      wordTemplate: {
        label: 'Letter Word template path: ',
        name: settingsNames.letter.wordTemplate,
        value: ''
      },
      saveTo: {
        label: 'Letter save to path: ',
        name: settingsNames.letter.saveTo,
        value: ''
      },
    },
    leases: {
      workBook: {
        label: 'Leases Excel workbook path :',
        name: settingsNames.leases.workBook,
        value: ''
      },
      tableName: {
        label: 'Leas\'s Excel Table name: ',
        name: settingsNames.leases.tableName,
        value: ''
      },
      wordTemplate: {
        label: 'Leases Word Template path :',
        name: settingsNames.leases.wordTemplate,
        value: ''
      },
      saveTo: {
        label: 'Leases\' save to path: ',
        name: settingsNames.leases.saveTo,
        value: ''
      },
    },
  };

  const groups = Object.values(settings);
  const inputs = groups.map(group => Object.values(group)).flat();

  let stored: settingInput[];
  localStorage.InvoicingPWA ? stored = JSON.parse(localStorage.InvoicingPWA) as settingInput[] : stored = inputs;
  if (get) return stored;

  const findSetting = (name: string, settings: settingInput[]) => settings?.find(setting => setting.name === name);

  if (values?.length) return save(values);//If the values of some settings have been passed as argument, we save the changes to the localStorage directly withouth showing inputs;

  const form = byID();
  if (!form) return;
  form.innerHTML = '';
  groups.forEach(group => showInputs(group));

  (function homeBtn() {
    showMainUI(true);
  })();

  function showInputs(group: setting) {
    const groupDiv = document.createElement('div');
    form!.appendChild(groupDiv);
    Object.values(group)
      .forEach(input => groupDiv.appendChild(createInput(input)));
  };

  function createInput({ label, name, value }: settingInput): HTMLDivElement {
    const container = document.createElement('div');
    const labelHtml = document.createElement('label');
    labelHtml.innerText = label;
    const input = document.createElement('input');
    input.classList.add('field');
    input.value = findSetting(name, stored)!.value || '';
    input.onchange = () => confirmSaving(input.value, label, name);
    container.appendChild(labelHtml);
    container.appendChild(input);
    return container
  };

  function confirmSaving(value: string, label: string, name: string) {
    if (!confirm(`Are you sure you want to change the ${label} localStorage value to ${value}?`)) return;
    save([[name, value]])
  }
  function save(values: [string, string][]) {
    values.forEach(([name, value]) => findSetting(name, stored)!.value = value.replaceAll('\\', '/') || '')
    localStorage.InvoicingPWA = JSON.stringify(stored);
  }
};

function throwAndAlert(message:string) {
  message = `Error: ${message}`
  alert(message);
  throw new Error(message) ;
}
function spinner(show: boolean) {
  if (!show) return document.querySelector('.spinner')?.remove();
  const form = byID('form');
  if (!form) return;
  const spinner = document.createElement('div');
  spinner.classList.add('spinner');
  form.prepend(spinner)
}