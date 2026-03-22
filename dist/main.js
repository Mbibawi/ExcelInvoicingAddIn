"use strict";
if (!localStorage.templatePath)
    localStorage.templatePath = prompt('Please provide the OneDrive full path (including the file name and extension) for the Word template', "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/FactureTEMPLATE [NE PAS MODIFIDER].docx");
const templatePath = localStorage.templatePath || alert('The template path is not valid or is missing');
if (!localStorage.tableName)
    localStorage.tableName = prompt('Please provide the name of the Excel Table where the invoicing data is stored', 'LivreJournal');
const tableName = localStorage.tableName || alert('The table name is not valid or is issing');
if (!localStorage.destinationFolder)
    localStorage.destinationFolder = prompt('Please provide the OneDrive path where the issued invoices will be stored', "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/Clients");
const destinationFolder = localStorage.destinationFolder || alert('the destination folder path is missing or not valid');
var TableRows, accessToken, tableTitles = JSON.parse(localStorage.tableTitles);
const tenantId = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
const byID = (id = "form") => document.getElementById(id);
const splitter = "; OR "; //This is the splitter that will be used to separate multiple values in the input fields. We need to use a splitter that is not likely to be included in the values themselves.
function getAccountsWorkBookPath() {
    if (!localStorage.accountsPath)
        localStorage.accountsPath = prompt('Please provide the OneDrive full path (including the file name and extension) for the Excel Workbook', "Legal/Mon Cabinet d'Avocat/Comptabilité/Comptabilité de Mon Cabinet_15 10 2023.xlsm");
    return localStorage.accountsPath || alert('The excel Workbook path is not valid');
}
(function RegisterServiceWorker() {
    // Check if the browser supports service workers
    if ("serviceWorker" in navigator) {
        window.addEventListener("load", async () => {
            try {
                const registration = await navigator.serviceWorker.register("/ExcelInvoicingAddIn/dist/sw.js");
                console.log("Service Worker registered successfully:", registration);
            }
            catch (error) {
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
    let installPromptEvent = null;
    //@ts-ignore
    window.addEventListener("beforeinstallprompt", (event) => {
        event.preventDefault(); // Prevent the default mini-infobar
        installPromptEvent = event;
        const installButton = document.getElementById("install-button");
        if (installButton) {
            installButton.style.display = "block"; // Show the install button
            installButton.addEventListener("click", async () => {
                if (installPromptEvent) {
                    await installPromptEvent.prompt(); // Show the install prompt
                    const choiceResult = await installPromptEvent.userChoice;
                    console.log("User install choice:", choiceResult.outcome);
                    installPromptEvent = null; // Clear the event after the prompt
                    installButton.style.display = "none"; // Hide the button
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
}
;
function selectForm(id) {
    showForm(id);
}
async function showForm(id) {
    const form = document.getElementById("form");
    form.innerHTML = '';
    if (!form)
        return;
    let table;
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        table = sheet.tables.getItem(tableName);
        const header = table.getHeaderRowRange();
        header.load('text');
        await context.sync();
        const body = table.getDataBodyRange();
        body.load('text');
        await context.sync();
        const headers = header.text[0];
        const clientUniqueValues = getUniqueValues(0, body.text);
        if (id === 'entry')
            await addingEntry(headers, clientUniqueValues);
        else if (id === 'invoice')
            invoice(headers, clientUniqueValues);
    });
    function invoice(title, clientUniqueValues) {
        const inputs = insertInputsAndLables([0, 1, 2, 3]); //Inserting the fields inputs (Client, Matter, Nature, Date)
        inputs.forEach(input => input?.addEventListener('focusout', async () => await _inputOnChange(input), { passive: true }));
        insertInputsAndLables(['Français', 'English'], true); //Inserting langauges checkboxes
        form.innerHTML += `<button onclick="generateInvoice()"> Generate Invoice</button>`; //Inserting the button that generates the invoice
        function insertInputsAndLables(indexes, checkBox = false) {
            const id = 'input';
            return indexes.map(index => {
                const input = document.createElement('input');
                if (checkBox)
                    input.type = 'checkbox';
                else if (Number(index) < 3)
                    input.type = 'text';
                else
                    input.type = 'date';
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
                if (Number(index) < 1)
                    createDataList(input?.id, clientUniqueValues); //We create a unique values dataList for the 'Client' input
                return input;
            });
        }
        ;
        async function _inputOnChange(input, unfilter = false) {
            const index = getIndex(input);
            if (index < 1)
                unfilter = true; //If this is the 'Client' column, we remove any filter from the table;
            //We filter the table accordin to the input's value and return the visible cells
            const visibleCells = await filterTable(tableName, [{ column: index, value: getArray(input.value) }], unfilter);
            if (visibleCells.length < 1)
                return alert('There are no visible cells in the filtered table');
            //We create (or update) the unique values dataList for the next input 
            const nextInput = getNextInput(input);
            if (!nextInput)
                return;
            const list = getUniqueValues(getIndex(nextInput), visibleCells);
            if (list?.length < 2)
                return nextInput.value = list[0].toString() || ''; //If there is only one value in the list, we set it as the value of the input and we don't create a data list for it because there is no need
            populateSelectElement(nextInput, list);
            function getNextInput(input) {
                let nextInput = input.nextElementSibling;
                while (nextInput?.tagName !== 'INPUT' && nextInput?.nextElementSibling) {
                    nextInput = nextInput.nextElementSibling;
                }
                ;
                return nextInput;
            }
            if (index === 1) {
                //!Need to figuer out how to create a multiple choice input for nature
                const nature = new Set((await filterTable(tableName, undefined, false)).map(row => row[index]));
                nature.forEach(el => form.appendChild(createCheckBox(undefined, el)));
            }
        }
        ;
    }
    async function addingEntry(title, uniqueValues) {
        await filterTable(tableName, undefined, true);
        for (const t of title) { //!We could not use for(let i=0; i<title.length; i++) because the await does not work properly inside this loop
            const i = title.indexOf(t);
            if (![4, 7].includes(i))
                form.appendChild(createLable(i)); //We exclued the labels for "Total Time" and for "Year"
            form.appendChild(await createInput(i));
        }
        ;
        const inputs = Array.from(document.getElementsByTagName('input'));
        inputs
            .filter(input => Number(input?.dataset.index) < 2)
            .forEach(input => input?.addEventListener('change', async () => await onFoucusOut(input), { passive: true }));
        inputs
            .filter(input => [4, 7].includes(Number(input?.dataset.index)))
            .forEach(input => input.style.display = 'none'); //We hide the inputs of some columns like the "Total Hours" or the "Link" column
        async function onFoucusOut(input) {
            const i = Number(input.dataset.index);
            const criteria = [{ column: i, value: getArray(input.value) }];
            let unfilter = false;
            if (i === 0)
                unfilter = true;
            await filterTable(tableName, criteria, unfilter);
            //if (i < 1) createDataList('input' + String(i + 1), getUniqueValues(i + 1, await filterTable(undefined, undefined)));
        }
        form.innerHTML += `<button onclick="addEntry()"> Ajouter </button>`;
        function createLable(i) {
            const label = document.createElement('label');
            label.htmlFor = 'input' + i.toString();
            label.innerHTML = title[i] + ':';
            return label;
        }
        async function createInput(i) {
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
            return input;
        }
    }
    function createCheckBox(input, id = '') {
        if (!input)
            input = document.createElement('input');
        input.type = 'checkbox';
        input.id += id;
        return input;
    }
    ;
}
/**
 * Creates a dataList with the provided id from the unique values of the column which index is passed as parameter
 * @param {string} id - the id of the dataList that will be created
 * @param {number} index - the index of the column from which the unique values of the datalist will be retrieved
 *
*/
/**
 *
 * @param select
 * @param uniqueValues
 * @param  {boolean} combine - determines whether we will add to the list an element containing all the options. Its defalult value is "false"
 */
function populateSelectElement(select, uniqueValues, combine = false) {
    const list = createDataList(select.id, uniqueValues, combine);
    if (!list)
        return;
    select.setAttribute('list', list.id);
    select.autocomplete = "on";
    return list;
}
/**
 *
 * @param id
 * @param uniqueValues
 * @param combine
 * @returns
 */
function createDataList(id, uniqueValues, combine = false) {
    //const uniqueValues = Array.from(new Set(visible.map(row => row[i])));
    if (!id || uniqueValues?.length < 2)
        return;
    id += 's';
    // Create a new datalist element
    let dataList = Array.from(document.getElementsByTagName('datalist')).find(list => list.id === id);
    if (dataList)
        dataList.remove();
    dataList = document.createElement('datalist');
    dataList.id = id;
    // Append options to the datalist
    uniqueValues.forEach(option => addOption(option));
    if (combine)
        addOption(uniqueValues.join(splitter));
    // Attach the datalist to the body or a specific element
    document.body.appendChild(dataList);
    function addOption(option) {
        const optionElement = document.createElement('option');
        optionElement.value = option;
        dataList?.appendChild(optionElement);
    }
    return dataList;
}
;
/**
 * Filters the Excel table based on a criteria
 * @param {[[number, string[]]]} criteria - the first element is the column index, the second element is the values[] based on which the column will be filtered
 */
async function filterTable(tableName, criteria, clearFilter = false) {
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const table = sheet.tables.getItem(tableName);
        if (clearFilter)
            table.autoFilter.clearCriteria();
        if (!criteria)
            return await getVisible();
        criteria.forEach(column => filterColumn(column.column, column.value));
        return await getVisible();
        function filterColumn(index, filter) {
            if (!index || !filter)
                return;
            table.columns.getItemAt(index).filter.applyValuesFilter(filter);
        }
        async function getVisible() {
            const visible = table.getDataBodyRange().getVisibleView();
            visible.load('values');
            await context.sync();
            return visible.values;
        }
    });
}
/**
 * Converts the ',' separated text in the input into an array
 * @param value
 * @returns {string[]}
 */
function getArray(value) {
    if (!value)
        return [];
    const array = value.replaceAll(', ', ',')
        .replaceAll(' ,', ',')
        .split(',');
    return array.filter((el) => el);
}
async function _generateInvoice() {
    const inputs = Array.from(document.getElementsByName('input'));
    if (!inputs)
        return;
    const lang = inputs.find(input => input.type === 'checkbox' && input.checked === true)?.id.slice(0, 3).toUpperCase() || 'FR';
    const discount = parseInt(inputs.find(input => input.id = 'discount')?.value || '0%');
    const visible = await filterTable(tableName, undefined, false);
    const date = new Date();
    const invoiceDetails = {
        number: getInvoiceNumber(new Date()),
        clientName: visible.map(row => String(row[0]))[0] || 'CLIENT',
        matters: (getUniqueValues(1, visible)).map(el => String(el)),
        adress: (getUniqueValues(15, visible)).map(el => String(el)),
        lang: lang
    };
    const savePath = `${destinationFolder}/${getInvoiceFileName(invoiceDetails.clientName, invoiceDetails.matters, invoiceDetails.number)}`;
    const [rows, totalsLabels] = getRowsData(visible, discount, lang, getInvoiceNumber(date));
    const accessToken = await getAccessToken() || '';
    await new GraphAPI(accessToken, '').createAndUploadWordDocument(templatePath, savePath, lang, 'Invoice', rows, getContentControlsValues(invoiceDetails, new Date()));
}
/**
 * Returns a string[][] representing the rows to be inserted in the Word table containing the invoice details
 * @param {string[][]} tableRows - The filtered Excel rows from which the data will be extracted and put in the required format
 * @param {string} lang - The language in which the invoice will issued
 * @returns {string[][]} - the rows to be added to the table. Each row has 4 elements
 */
function getRowsData(tableRows, discount = 0, lang, invoiceNumber) {
    const labels = {
        totalFees: {
            nature: ['Honoraire'],
            FR: 'Total honoraires',
            EN: 'Total Fees'
        },
        totalExpenses: {
            nature: ['Débours/Dépens', 'Rétrocession d\'honoraires', 'Débours/Dépens - Ackad Law Office', 'Charges déductibles'],
            FR: 'Total débours et frais',
            EN: 'Total Expenses'
        },
        totalPayments: {
            nature: ['Provision/Règlement'],
            FR: 'Total provisions reçues',
            EN: 'Total Downpayments'
        },
        totalTimeSpent: {
            nature: [],
            FR: 'Total des heures facturables (hors prestations facturées au forfait) ',
            EN: 'Total billable hours (other than lump-sum billed services)'
        },
        totalDue: {
            nature: [],
            FR: 'Montant dû',
            EN: 'Total Due'
        },
        totalReinbursement: {
            nature: [],
            FR: 'A rembourser',
            EN: 'Reimbursement'
        },
        totalDeduction: {
            nature: ['Remise'],
            FR: 'Total des remises sur honoraires',
            EN: 'Total fees\' discounts'
        },
        netFees: {
            nature: [],
            FR: 'Total honoraires après réduction',
            EN: 'Total fee after discount'
        },
        discountDescription: {
            //This value is not used
            nature: [],
            FR: `XXX% de remise sur les honoraires`,
            EN: `XXX% discount on accrued fees`
        },
        hourlyBilled: {
            nature: [],
            FR: 'facturation au temps passé\u00A0:',
            EN: 'hourly billed:',
        },
        hourlyRate: {
            nature: [],
            FR: 'au taux horaire de\u00A0:',
            EN: 'at an hourly rate of:',
        },
        decimal: {
            nature: [],
            FR: ',',
            EN: '.'
        },
        bankHoler: {
            nature: [],
            FR: 'Titulaire du compte',
            EN: 'Account holder'
        },
        bankName: {
            nature: [],
            FR: 'Banque',
            EN: 'Bank'
        },
        bankAdress: {
            nature: [],
            FR: 'Adresse',
            EN: 'Adress'
        },
    };
    const totalsLabels = [];
    const colDate = 3, colAmount = 9, colVAT = 10, colHours = 7, colRate = 8, colNature = 2, colDescr = 14; //Indexes of the Excel table columns from which we extract the date 
    const wordRows = tableRows.map(row => {
        const date = dateFromExcel(Number(row[colDate]));
        const time = getTimeSpent(Number(row[colHours]));
        let description = `${String(row[colNature])} : ${String(row[colDescr])}`; //Column Nature + Column Description;
        //If the billable hours are > 0, we add to the description: time spent and hourly rate
        if (time)
            description += ` (${labels.hourlyBilled[lang]} ${time}, ${labels.hourlyRate[lang]} ${Math.abs(row[colRate]).toString()}\u00A0€).`;
        const rowValues = [
            getDateString(date), //Column Date
            description,
            getAmountString(row[colAmount] * -1), //Column "Amount": we inverse the +/- sign for all the values 
            getAmountString(Math.abs(row[colVAT])), //Column VAT: always a positive value
        ];
        return rowValues;
    });
    pushTotalsRows();
    return [wordRows, totalsLabels];
    function pushTotalsRows() {
        //Adding rows for the totals of the different categories and amounts
        const total = (lable) => [colAmount, colVAT].map(col => sumColumn(col, lable.nature)); //!It always returns the absolute values of the total amount and the total VAT
        const amount = (v) => v[0];
        const totalFees = total(labels.totalFees);
        const feesDiscount = totalFees.map(amount => amount * (discount / 100)); //This is an additional discount applied when the invoice is issued. The Excel table may already include other discounts registered as "Remise"
        const feesDeductions = total(labels.totalDeduction).map((amount, index) => amount += feesDiscount[index]); //This is the total of the deductions from the fees: the "Remise" deductions, and the additional discount added at the time the invoice is issued
        const netFees = totalFees.map((amount, index) => amount - feesDeductions[index]);
        const totalPayments = total(labels.totalPayments);
        const totalExpenses = total(labels.totalExpenses);
        const totalTimeSpent = [sumColumn(colHours), NaN]; //by omitting to pass the "natures" argument to sumColumn, we do not filter the "Total Time" column by any crieteria. We will get the sum of all the column. since the VAT = NaN, the VAT cell will end up empty.
        const totalDue = netFees.map((amount, index) => amount + totalExpenses[index] - totalPayments[index]);
        const percentage = (amount(feesDeductions) / amount(totalFees)) * 100;
        ['EN', 'FR'].forEach((lang) => labels.totalDeduction[lang] += ` (${percentage}%)`);
        (function pushTotalsRows() {
            pushRow(labels.totalFees, totalFees);
            pushRow(labels.totalDeduction, feesDeductions, !amount(feesDeductions));
            pushRow(labels.netFees, netFees, !(amount(netFees) < amount(totalFees))); //We don't push this row if the there is no deduction applied on the fees or if the deduction is = 0
            pushRow(labels.totalTimeSpent, totalTimeSpent, !amount(totalTimeSpent));
            pushRow(labels.totalExpenses, totalExpenses, !amount(totalExpenses));
            pushRow(labels.totalPayments, totalPayments, !amount(totalPayments));
            amount(totalDue) < 0 ? pushRow(labels.totalReinbursement, totalDue) : pushRow(labels.totalDue, totalDue);
        })();
        (function addDiscountRowToExcel() {
            if (!discount)
                return;
            const newRow = tableRows
                .find(row => labels.totalFees.nature.includes(row[colNature]));
            if (!newRow)
                return;
            const [amount, vat] = feesDiscount; //!The discount must be added as a positive number. This is like a payment made by the client
            const descr = prompt('Provide a description for the discount', `Remise sur les honoraires de la facture n° ${invoiceNumber}`) || '';
            const date = getISODate(new Date());
            const cells = [
                [colNature, 'Remise'],
                [colAmount, amount],
                [colVAT, vat],
                [colDate, date],
                [colDate + 1, date],
                [colDescr, descr]
            ];
            cells.forEach(([col, value]) => newRow[col] = value);
            addNewEntry(true, newRow);
        })();
        function pushRow(rowLable, [amount, vat], ignore = false) {
            if (ignore || !amount || isNaN(amount))
                return;
            const lable = rowLable?.[lang] || '';
            if (lable)
                totalsLabels.push(lable);
            const value = rowLable === labels.totalTimeSpent ? getTimeSpent(amount) : getAmountString(amount);
            wordRows.push([
                lable,
                '',
                value,
                getAmountString(vat) //VAT is always a positive value
            ]);
        }
        /**
         *
         * @param {number} col - the index of the column to be summed
         * @param {string[] | null} natures - the natures of the rows to be included in the sum. If null, we include all the rows regardless of their nature
         * @returns
         */
        function sumColumn(col, natures = []) {
            let rows = tableRows;
            if (natures.length)
                rows = tableRows.filter(row => natures.includes(row[colNature])); //If natures is specified, we filter the rows to include only the ones whose nature is included in the natures array 
            return Math.abs(sumArray(rows.map(row => Number(row[col])))); //!We return the absolute value of the total
        }
    }
    function sumArray(values) {
        let sum = 0;
        values.forEach(value => sum += value);
        return sum;
    }
    function getAmountString(value) {
        if (isNaN(value))
            return '';
        const amount = value.toLocaleString(`${lang.toLowerCase()}-${lang.toUpperCase()}`, {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });
        const versions = {
            FR: `${amount}\u00A0€`,
            EN: `€\u00A0${amount}`,
        };
        return versions[lang];
    }
    /**
     * Convert the time as retrieved from an Excel cell into 'hh:mm' format
     * @param {number} time - The time as stored in an Excel cell
     * @returns {string} - The time as 'hh:mm' format
     */
    function getTimeSpent(time) {
        if (!time || time <= 0)
            return '';
        time = time * (60 * 60 * 24); //84600 is the number in seconds per day. Excel stores the time as fraction number of days like "1.5" which is = 36 hours 0 minutes 0 seconds;
        const minutes = Math.floor(time / 60);
        const hours = Math.floor(minutes / 60);
        return [hours, minutes % 60, 0]
            .map(el => el.toString().padStart(2, '0'))
            .join(':');
    }
}
function getContentControlsValues(invoice, date) {
    const fields = {
        dateLabel: {
            title: 'LabelParisLe',
            value: { FR: 'Paris le ', EN: 'Paris on ' }[invoice.lang] || '',
        },
        date: {
            title: 'RTInvoiceDate',
            value: getDateString(date),
        },
        numberLabel: {
            title: 'LabelInvoiceNumber',
            value: { FR: 'Facture n°\u00A0:', EN: 'Invoice No.:' }[invoice.lang] || '',
        },
        number: {
            title: 'RTInvoiceNumber',
            value: invoice.number,
        },
        subjectLable: {
            title: 'LabelSubject',
            value: { FR: 'Affaires\u00A0: ', EN: 'Matters: ' }[invoice.lang] || '',
        },
        subject: {
            title: 'RTMatter',
            value: invoice.matters.join(' & '),
        },
        fee: {
            title: 'LabelTableHeadingHonoraire',
            value: { FR: 'Honoraire/Débours', EN: 'Fees/Expenses' }[invoice.lang] || '',
        },
        amount: {
            title: 'LabelTableHeadingMontantTTC',
            value: { FR: 'Montant TTC', EN: 'Amount VAT Included' }[invoice.lang] || '',
        },
        vat: {
            title: 'LabelTableHeadingTVA',
            value: { FR: 'TVA', EN: 'VAT' }[invoice.lang] || '',
        },
        disclaimer: {
            title: 'LabelDisclamer' + ['French', 'English'].find(el => !el.toUpperCase().startsWith(invoice.lang)) || 'English',
            value: 'DELETECONTENTECONTROL', //!by setting text = "DELETECONTENTECONTROL", the contentControl will be deleted
        },
        clientName: {
            title: 'RTClient',
            value: invoice.clientName,
        },
        adress: {
            title: 'RTClientAdresse',
            value: invoice.adress.join(' & '),
        },
    };
    return Object.values(fields).map(RT => [RT.title, RT.value]);
}
function getUniqueValues(index, array) {
    if (!array)
        array = [];
    return Array.from(new Set(array.map(row => row[index])))
        .map(el => el); //we remove empty strings/values
}
;
class GraphAPI {
    constructor(accessToken, filePath, sessionId, presist = false) {
        this.GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0/me/drive/root:/";
        this.methods = {
            post: "POST",
            get: "GET",
            patch: "PATCH",
        };
        this.accessToken = accessToken;
        this.filePath = filePath || '';
        this.sessionId = sessionId || '';
    }
    /**
   * Creates a new Graph API File session and returns its id
   * @returns
   */
    async createFileSession(persist = false) {
        const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/createSession`;
        const body = { persistChanges: persist };
        const response = await this.sendRequest(endPoint, this.methods.post, body, undefined, undefined, "Erro: Failed to create workbook session");
        const session = await response?.json();
        return session.id;
    }
    /**
     * Closes the current Excel file session
     */
    async closeFileSession(sessionId) {
        if (!this.filePath)
            return;
        const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/closeSession`;
        const resp = await this.sendRequest(endPoint, this.methods.post, undefined, sessionId, undefined, 'Error closing the session');
        if (resp)
            console.log(`The session was closed successfully! ${await resp.text()}`);
    }
    /**
     * Returns all the rows of an Excel table in a workbook stored on OneDrive, using the Graph API
     * @param {string} tableName - Name of the table to be fetched
     * @param {boolean} headers - Its default value is true. If true, it calls the "/range" endpoint and returns the whole table including the headers row, otherwise, it calls the "/rows" endpoint and returns only the body (the rows) of the table. The structure of the date returned is different for each endpoint
     * @param {boolean} columns - If true it will return the columns
     * @returns {any[][] | number | void} - All the rows (including the title) of the Excel table
     */
    async fetchExcelTable(tableName, headers = true, columns) {
        if (!this.accessToken)
            this.accessToken = await getAccessToken() || '';
        let endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/`;
        if (headers)
            endPoint += 'range'; //The "range" endpoint returns all the table including the headers row
        if (!headers)
            endPoint += 'rows'; //The "rows" endpoint returns only  the body of the table without the headers
        else if (columns)
            endPoint += 'columns';
        const response = await this.sendRequest(endPoint, this.methods.get, undefined, undefined, undefined, `Error fetching row count`);
        const data = await response?.json();
        if (headers)
            return data.values;
        else
            return data.value.flatMap((row) => row.values); //! the graph api returns an object with a "value" property which is an array of rows, each row is also an object with a "values" property which is an array of the cells values of the row. So we need to flatMap() the data to return an array of rows, each row being an array of cells values;
    }
    ;
    /**
     * Filters an Excel table column based on the values
     * @param {string} filePath - the full path and file name of the Excel workbook
     * @param {string} tableName - the name of the table that will be filtered
     * @param {string} columnName - the name of the column that will be filtered
     * @param {string[]|number[]|boolean[]} values - the values based on which the column will be filtered
     * @param {string} sessionId - the id of the current Excel file session
     * @param {string} accessToken - the access token
     * @returns {string}
     */
    async filterExcelTable(tableName, columnName, values, sessionId = this.sessionId, onValues = true) {
        if (!columnName || !values?.length || !this.filePath)
            return;
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
        else
            body = {
                "criteria": {
                    "filterOn": "custom",
                    "criterion1": values[0],
                    "criterion2": values[1] || null,
                    "operator": "And",
                }
            };
        const response = await this.sendRequest(endPoint, this.methods.post, body, sessionId, undefined, 'Error while applying filter to the Excel table');
        if (response)
            console.log(`Filter successfully applied to column ${columnName}!`);
    }
    ;
    /**
     * Returns the visible cells of a filtered Excel table using Graph API
     * @param {string} filePath - the full path and file name of the Excel workbook
     * @param {string} tableName - the name of the table that will be filtered
     * @param {string} sessionId - the id of the current Excel file session
     * @param {string} accessToken - the access token
     * @returns {any[][]} - the visible cells of the filtered table
     */
    async getVisibleCells(tableName, sessionId) {
        if (!tableName || !this.filePath)
            return alert('Either the tableName or the filePath are mission or not valid');
        // Step 3: Apply filter using the column name
        const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/range/visibleView`;
        const response = await this.sendRequest(endPoint, this.methods.get, undefined, sessionId, undefined, "Error applying filter");
        const data = await response?.json();
        return data.values;
    }
    ;
    /**
     * Clears the filters on an Excel table using the Graph API
     * @param {string} filePath - the full path and file name of the Excel workbook
     * @param {string} tableName - the name of the table that will be filtered
     * @param {string} sessionId - the id of the current Excel file session
     * @param {string} accessToken - the access token
     */
    async clearFilterExcelTable(tableName, sessionId) {
        const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/clearFilters`;
        await this.sendRequest(endPoint, this.methods.post, undefined, sessionId, undefined, "Erro: Failed to clear the Excel table filter");
    }
    ;
    /**
   * Adds a new row to the Excel table using the Grap API
   * @param {string} row - The row that will be added to the Excel table
   * @param {number} index - The index at which the row will be added
   * @param {string} workbookPath - The full path of the Excel file
   * @param {string} tableName - The name of the Excel table
   * @param {string} accessToken - The Graph API access token
   * @param {boolean} filter - If true, the table will be filtered after the row is added
   * @returns
   */
    async addRowToExcelTable(row, index, tableName, filter = false) {
        if (!this.filePath || !tableName || !row?.length)
            return alert('The filePath or the tableName argument is missing or not valid');
        const sessionId = this.sessionId || await this.createFileSession(true);
        if (!sessionId)
            return alert('The sessionId is missing Check the console.log for more details');
        const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/rows`; //The url to add a row to the table
        const body = {
            index: index,
            values: [row],
        };
        await this.clearFilterExcelTable(tableName, sessionId); //We clear the filtering of the table
        const resp = await this.sendRequest(endPoint, this.methods.post, body, sessionId, undefined, "Error adding row to the Excel Table");
        if (resp)
            console.log("Row added successfully!");
        for (const index of [0, 1]) {
            if (!filter)
                break;
            await this.filterExcelTable(tableName, tableTitles[index], [row[index].toString()], sessionId); //!We use "for of" loop because forEach doesn't await
        }
        await this.sortExcelTable(tableName, [[3, true]], false, sessionId); //We sort the table by the first column (the date column)
        const visible = await this.getVisibleCells(tableName, sessionId);
        await this.closeFileSession(sessionId);
        return visible;
    }
    ;
    /**
     * Updates a specific row in an Excel table using the file path.
     * @param {string} accessToken - OAuth2 token.
     * @param {string} workbookPath -
     * @param {string} tableName - The name of the table.
     * @param {number} rowIndex - 0-based index of the row.
     * @param {Array} values - 1D array of values for the row.
     */
    async updateExcelTableRow(tableName, rowIndex, values) {
        if (!this.filePath || !tableName || !rowIndex || !values?.length)
            return alert('One of the arguments is missing or not valid');
        const sessionId = await this.createFileSession();
        if (!sessionId)
            return alert('Failed to create a new Session');
        const url = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/rows/itemAt(index=${rowIndex})`;
        const body = {
            values: [values] // API requires a 2D array
        };
        try {
            const response = await this.sendRequest(url, this.methods.patch, body, sessionId, undefined, "Error while updating the Excel Table row's values");
            const data = await response?.json();
            console.log('Update Successful:', data);
            return data;
        }
        catch (err) {
            console.error('Update Failed:', err);
            this.closeFileSession(sessionId);
        }
    }
    ;
    /**
   * Creates an invoice Word document from the invoice Word template, then uploads it to the destination folder
   * @param {string} accessToken - The access token that will be used to authenticate the user
   * @param {string} templatePath - The full path of the Word invoice template
   * @param {string} savePath - The full path of the destination folder where the new invoice will be saved
   * @param {string} lang - The language in which the invoice will be issued
   * @param {string} tableTitle - The title of the table in the Word document that will be updated
   * @param {string[][]} rows - The rows that will be added to the table in the Word document
   * @param {string[][]} contentControls - The titles and text of each of the content controls that will be updated in the Word document
   * @param {string[]} totalsLabels - The labels of the rows that will be formatted as totals
   * @returns
   */
    async createAndUploadWordDocument(templatePath, savePath = this.filePath, lang, tableTitle, rows, contentControls, totalsLabels) {
        if (!this.accessToken || !templatePath || !this.filePath)
            return;
        const blob = await this.fetchFileFromOneDrive(templatePath); //!We must provide the Word templatePath not the Excel workbook path stored in the this.filePath variable
        if (!blob)
            return;
        const [doc, zip] = await this.convertBlobIntoXML(blob);
        if (!doc)
            return;
        const xml = new XML(doc, lang);
        const schema = xml.schema();
        (function editTable() {
            if (!rows || !tableTitle)
                return;
            const tables = xml.getTables(doc);
            const table = xml.findTableByTitle(tables, tableTitle);
            if (!table)
                return;
            const firstRow = xml.getTableRow(table, 1);
            rows.forEach((row, index) => {
                const newXmlRow = xml.insertRowAfterFirst(table, firstRow, NaN, true) || table.appendChild(xml.createTableRow());
                if (!newXmlRow)
                    return;
                const isTotal = totalsLabels?.includes(row[0]);
                const isLast = index === rows.length - 1;
                return editCells(newXmlRow, row, isLast, isTotal);
            });
            firstRow.remove(); //We remove the first row when we finish
            function editCells(tableRow, values, isLast = false, isTotal = false) {
                const cells = xml.getRowCells(tableRow) || values.map(v => tableRow.appendChild(xml.createTableCell())); //getting all the cells in the row element
                cells.forEach((cell, index) => {
                    const textElement = xml.getTextElement(cell, 0) || xml.appendParagraph(cell);
                    if (!textElement)
                        return console.log('No text element was found !');
                    const pPr = xml.setTextLanguage(cell); //We call this here in order to set the language for all the cells. It returns the pPr element if any.
                    textElement.textContent = values[index];
                    (function totalsRowsFormatting() {
                        if (!isLast && !isTotal)
                            return;
                        (function cellBackgroundColor() {
                            const tcPr = xml.getPropElement(cell, 0) || cell.prepend(xml.createPropElement(cell));
                            const shadow = xml.getShadowElement(tcPr, 0) || tcPr.appendChild(xml.createShadowElement()); //Adding background color to cell
                            shadow.setAttributeNS(schema, 'val', "clear");
                            shadow.setAttributeNS(schema, 'fill', 'D9D9D9');
                        })();
                        (function paragraphStyle() {
                            if (!pPr)
                                return console.log('No "w:pPr" or "w:rPr" property element was found !');
                            const style = xml.getParagraphStyle(pPr, 0) || pPr.appendChild(xml.createParagraphStyle());
                            style.setAttributeNS(schema, 'val', xml.getStyle(index, isTotal && !isLast));
                        })();
                    })();
                });
            }
        })();
        (function editContentControls() {
            if (!contentControls?.length)
                return;
            const ctrls = xml.getContentControls(doc);
            contentControls
                .forEach(([title, text]) => {
                const sameTitle = xml.findContentControlsByTitle(ctrls, title); //!we  retrieve all then XML ContentControls having the same title
                sameTitle.forEach(control => xml.editContentControlText(control, text));
            });
        })();
        await this.convertXMLToBlobAndUpload(doc, zip, savePath);
    }
    ;
    /**
   * Converts the blob of a Word document into an XML
   * @param blob - the blob of the file to be converted
   * @returns {[XMLDocument, JSZip]} - The xml document, and the zip containing all the xml files
   */
    //@ts-expect-error
    async convertBlobIntoXML(blob) {
        //@ts-ignore
        const zip = new JSZip();
        const arrayBuffer = await blob.arrayBuffer();
        await zip.loadAsync(arrayBuffer);
        const documentXml = await zip.file("word/document.xml").async("string");
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(documentXml, "application/xml");
        return [xmlDoc, zip];
    }
    /**
   * Filters an Excel table column based on the values
   * @param {string} filePath - the full path and file name of the Excel workbook
   * @param {string} tableName - the name of the table that will be filtered
   * @param {[string, boolean][]} columns - each element contains the name of the column and whether it will be sorted ascending or descending
   * @param {string} sessionId - the id of the current Excel file session
   * @param {string} accessToken - the access token
   * @returns {string}
   */
    async sortExcelTable(tableName, columns, matchCase, sessionId) {
        if (!this.filePath)
            return;
        // Step 3: Apply filter using the column name
        const endPoint = `${this.GRAPH_API_BASE_URL}${this.filePath}:/workbook/tables/${tableName}/sort/apply`;
        const fields = columns.map(([index, ascending]) => {
            return {
                key: index,
                ascending: ascending,
                sortOn: "value"
            };
        });
        const body = {
            fields: fields,
            "matchCase": matchCase
        };
        const resp = await this.sendRequest(endPoint, this.methods.post, body, sessionId, undefined, "Error sorting table");
        if (resp)
            console.log(`Table successfully sorted according to columns criteria: ${columns.map(([col, asc]) => col).join(' & ')}!`);
    }
    ;
    /**
   * Returns a blob from a file stored on OneDrive, using the Graph API and the file path
   * @param {string} accessToken
   * @param {string} filePath
   * @returns {Blob} - A blob of the fetched file, if successful
   */
    async fetchFileFromOneDrive(filePath = this.filePath) {
        const endPoint = `${this.GRAPH_API_BASE_URL}${filePath}:/content`;
        const response = await this.sendRequest(endPoint, this.methods.get, undefined, undefined, undefined, "Failed to fetch Word template");
        return await response?.blob(); // Returns the Word template as a Blob
    }
    ;
    /**
   * Uploads a file blob to OneDrive using the Graph API
   * @param {Blob } blob
   * @param {string} filePath
   * @param {string} accessToken
   */
    async uploadFileToOneDrive(blob, filePath) {
        if (!filePath || !this.accessToken)
            return;
        const endpoint = `${this.GRAPH_API_BASE_URL}${filePath}:/content`;
        const response = await fetch(endpoint, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${this.accessToken}`,
                'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // Correct MIME type for Word docs
            },
            body: blob, // Use the template's content as the new document's content
        });
        response.ok ? alert('succefully uploaded the new file') : console.log('failed to upload the file to onedrive error = ', await response.json());
    }
    ;
    /**
   * Converts an XML Word document into a Blob, and uploads it to OneDrive using the Graph API
   * @param {XMLDocument} doc
   * @param {JSZip} zip
   * @param {string} filePath - the full OneDrive file path (including file name and extension) of the file that will be uploaded
   * @param {string} accessToken - the Graph API accessToken
   */
    //@ts-expect-error
    async convertXMLToBlobAndUpload(doc, zip, filePath = this.filePath) {
        const blob = await convertXMLIntoBlob();
        if (!blob)
            return;
        await this.uploadFileToOneDrive(blob, filePath);
        async function convertXMLIntoBlob() {
            const serializer = new XMLSerializer();
            let modifiedDocumentXml = serializer.serializeToString(doc);
            zip.file("word/document.xml", modifiedDocumentXml);
            return await zip.generateAsync({ type: "blob" });
        }
    }
    ;
    async sendRequest(endPoint, method, body, sessionId, contentType, message = "") {
        if (!this.accessToken)
            return;
        const request = {
            method: method,
            headers: this.graphHeaders(sessionId, contentType)
        };
        if (body)
            request.body = JSON.stringify(body);
        const response = await fetch(endPoint, request);
        if (response?.ok)
            return response;
        message = `${message || `Error while sending ${method} request`}:\n ${await response?.text()}`;
        alert(message);
        if (sessionId)
            await this.closeFileSession(sessionId);
        throw new Error(message);
    }
    ;
    /**
   * Returns the headers of the Microsoft Graph API calls
   */
    graphHeaders(sessionId, contentType) {
        const headers = {
            'Authorization': `Bearer ${this.accessToken}`,
            'Content-Type': contentType || 'application/json',
        };
        if (sessionId)
            headers["workbook-session-id"] = sessionId;
        return headers;
    }
    async getExcelTableRowsCount(filePath, tableName, accessToken) {
        const url = `${this.GRAPH_API_BASE_URL}${filePath}:/workbook/tables/${tableName}/rows/$count`;
        const response = await fetch(url, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${accessToken}`
            }
        });
        if (response.ok) {
            const rowCount = await response.text(); // The API returns a number as plain text
            console.log(`Row count: ${rowCount}`);
            return parseInt(rowCount, 10); // Convert to number
        }
        else {
            console.error("Error fetching row count:", await response.text());
            return null;
        }
    }
}
;
class XML {
    constructor(doc, lang) {
        this.tags = {
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
        this.schema = () => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        this.Pr = (tag) => `${tag.replace('w:', '')}Pr`; //!we need to remove the "w:" prefix from the tag
        this.doc = doc;
        this.lang = lang;
    }
    /**
     * Returns all the XML ContentConrol ("sdt") elements of the parent XML Element passed as argument
     * @param {XMLDocument | Element} parent - the parent XML Document or Element that we want to retrieve all its nested XML ContentControl elements
     * @returns {Element[]}
     */
    getContentControls(parent = this.doc) {
        return this.getXMLElements(parent, this.tags.ctrl);
    }
    /**
     * Returns a XML ContentControl element ("sdt") nested in the parent XML Element passed as argument
     * @param {Element} parent - the parent XML element of the XML ContentControl element we want to retrieve
     * @param {number} index - the index of the XML ContentControl element we want to retrieve
     * @returns {Element}
     */
    getControlContent(parent, index) {
        return this.getXMLElements(parent, this.tags.ctrlContent, index);
    }
    /**
     * Returns all the XML table ("tbl") elements nested in the XML Document or the XML Element passed as argument
     * @param {XMLDocument | Element} parent - the parent XML document or Element for which we want to retrieve the XML tables
     * @returns {Elemnt[]}
     */
    getTables(parent = this.doc) {
        return this.getXMLElements(parent, this.tags.table);
    }
    /**
     * Returns an XML table row ("tr") element nested in the XML table element passed as argument
     * @param {Element} table - the table element for which we want to retrieve a specifc XML row element by its index
     * @param {number} index - the index of the row
     * @returns {Element}
     */
    getTableRow(table, index) {
        return this.getXMLElements(table, this.tags.row, index);
    }
    /**
   * Returns the XML '[tag]Pr' element child of any element
   * @param {Element} parent - the parent XML Element for which we are trying to retrieve the "[tag]Pr" element
   * @param {number} index - the index of the "[tag]Pr" element
   * @returns {Element}
   */
    getPropElement(parent, index = 0) {
        const tag = this.Pr(parent.tagName.toLowerCase());
        return this.getXMLElements(parent, tag, index);
    }
    /**
     * Creates and returns a table row ("tr") XML element
     * @returns {Element}
     */
    createTableRow() {
        return this.createXMLElement(this.tags.row);
    }
    /**
     * Creates and returns a table XML cell ("tc") element
     * @returns {Element}
     */
    createTableCell() {
        return this.createXMLElement(this.tags.cell);
    }
    /**
     * Creates and returns a "[tag]Pr" XML element which is a child element holding the properties of the parent XML Element
     * @param {Element} parent - the parent XML Element from which we will retrieve the tag of XML "[tag]Pr" element that we will create
     * @returns {Element}
     */
    createPropElement(parent) {
        const tag = this.Pr(parent.tagName.toLowerCase());
        return this.createXMLElement(tag);
    }
    /**
     * Returns the XML cell elements of row in the table
     * @param {Element} tableRow - the XML row element that we want to retrieve its XML cell children elements
     * @returns {Element[]}
     */
    getRowCells(tableRow) {
        return this.getXMLElements(tableRow, this.tags.cell);
    }
    /**
     * Returns a text XML element of the parent according to its index
     * @param {Element} parent - the XML element that we want to retrieve one of its text ("t") XML children
     * @param {number} index - the index of the text ("t") XML element we want to retrieve
     * @returns {Element}
     */
    getTextElement(parent, index) {
        return this.getXMLElements(parent, this.tags.text, index);
    }
    /**
     * Creates and returns a pargraph style ("pStyle") XML element
     * @returns {Element}
     */
    createParagraphStyle() {
        return this.createXMLElement(this.tags.style);
    }
    /**
     * Returns a XML shadow element ("shd") of the parent according to its index passed as argument
     * @param {Element} parent -
     * @param {number} index -
     * @returns {Element}
     */
    getShadowElement(parent, index) {
        return this.getXMLElements(parent, this.tags.shadow, index);
    }
    /**
     * Creates and returns a XML shadow element ("shd")
     * @returns {Element} - an XML shadow ("shd") element
     */
    createShadowElement() {
        return this.createXMLElement(this.tags.shadow);
    }
    /**
      * Looks for a child "w:p" (paragraph) element, if it doesn't find any, it looks for a "w:r" (run) element.
      * @param {Element} parent - the parent XML of the paragraph or run element we want to retrieve.
      * @returns {Element | undefined} - an XML element representing a "w:p" (paragraph) or, if not found, a "w:r" (run), or undefined
      */
    getParagraphOrRun(parent) {
        return this.getXMLElements(parent, this.tags.paragraph, 0) || this.getXMLElements(parent, this.tags.run, 0);
    }
    /**
     * Returns the cells of row in the table
     * @param {Element} tableRow
     */
    getParagraphStyle(parent, index) {
        return this.getXMLElements(parent, this.tags.style, index);
    }
    insertRowAfterFirst(table, firstRow, after = -1, clone = false) {
        const self = this;
        if (clone)
            return cloneFirstRow();
        else
            return create();
        function create() {
            if (!table)
                return;
            const row = self.createTableRow();
            after >= 0 ? self.getXMLElements(table, self.tags.row, after)?.insertAdjacentElement('afterend', row) :
                table.appendChild(row);
            return row;
        }
        function cloneFirstRow() {
            if (!firstRow)
                return;
            const row = firstRow.cloneNode(true);
            table?.appendChild(row);
            return row;
        }
        ;
    }
    getStyle(cell, isTotal = false) {
        let style = 'Invoice';
        if (cell === 0 && isTotal)
            style += 'BoldItalicLeft';
        else if (cell === 0)
            style += 'BoldLeft';
        else if (cell === 1)
            style += 'NotBoldItalicLeft';
        else if (cell === 2 && isTotal)
            style += 'BoldItalicCentered';
        else if (cell === 2)
            style += 'BoldCentered';
        else if (cell === 3)
            style += 'BoldItalicCentered';
        else
            style = '';
        return style;
    }
    /**
     *
     * @param {Element[]} ctrls - the XML ContentControls array from which we will retrieve an XML ContentControl by its title
     * @param {string} title - the title of the XML ContentControl we want to retrieve
     * @param {number} index - if omitted, the function will return a collection of all the XML ContentControl elements having the same title. Otherwise it will return a ContentControl by its index
     * @returns {Element | Element[] | undefined}
     */
    findContentControlsByTitle(ctrls, title) {
        return this.findElementsByPropertyValue(ctrls, this.tags.alias, title);
    }
    /**
     * Finds and returns a XML Table by its title ("tblCaption")
     * @param {Element[]} tables - the XML tables array in which we will search for a table having the specified title
     * @param {string} title - the title of the table
     * @returns {Element | undefined}
     */
    findTableByTitle(tables, title) {
        return this.findElementsByPropertyValue(tables, this.tags.tableCaption, title)?.[0];
    }
    /**
     *
     * @param {Element[]} elements - the XML Elements collection in which we will search for specific XML elemnt(s) by the value of a given property
     * @param {string} tag - the name of property in which the title of the XML Element title is stored
     * @param {string} value - the value of the property we are looking for
     * @returns {Element []}
     */
    findElementsByPropertyValue(elements, tag, value) {
        if (!tag || !value)
            return [];
        const children = (parent) => this.getXMLElements(parent, tag); //This returns the child elements of the parent (if any) having the specified tag. The children hold a property of the element
        return elements.filter(element => children(element)?.find(child => child.getAttributeNS(this.schema(), 'val') === value));
    }
    /**
  * Adds a new paragraph XML element or appends a cloned paragraph, and in both cases, it returns the textElement of the paragraph
  * @param {Element} element - The element to which the new paragraph will be appended if the parent argument is not provided. If the parent argument is provided, the element will be cloned assuming that this is a pargraph element
  * @param {Elemenet} parent - If provided, element will be cloned and appended to parent.
  * @returns {Element} the textElemenet attached to the paragraph
  */
    appendParagraph(element, parent) {
        const self = this;
        if (parent)
            return clone();
        else
            return create();
        function clone() {
            const parag = element?.cloneNode(true);
            parent?.appendChild(parag);
            return self.getXMLElements(parag, 't', 0);
        }
        function create() {
            const parag = element.appendChild(self.createXMLElement(self.tags.paragraph));
            parag.appendChild(self.createXMLElement(self.Pr(self.tags.paragraph)));
            const run = parag.appendChild(self.createXMLElement(self.tags.run));
            return run.appendChild(self.createXMLElement(self.tags.text));
        }
    }
    createXMLElement(tag) {
        return this.doc.createElementNS(this.schema(), tag);
    }
    /**
     * Returns the XML element(s) nested under the XML parent element by its/their tag;
     * @param {XMLDocument | Element} parent - the parent XML Document or XML Element nesting the XML Element(s) we want to retrieve
     * @param {string} tag - the tag of the XML Element(s) we want to retrieve
     * @param {number} index - if provided, the function will only return the element having the specified index
     * @returns {Element[] | Element | undefined}
     */
    getXMLElements(parent, tag, index = NaN) {
        const elements = parent?.getElementsByTagNameNS(this.schema(), tag);
        if (!elements.length)
            return;
        if (!isNaN(index))
            return elements[index];
        return Array.from(elements);
    }
    editContentControlText(control, text) {
        if (text === "DELETECONTENTECONTROL")
            return control.remove();
        if (!text)
            text = 'NO VALUE WAS PROVIDED';
        const sdtContent = this.getControlContent(control, 0);
        if (!sdtContent)
            return;
        const paragTemplate = this.getParagraphOrRun(sdtContent); //This will set the language for the paragraph or the run
        if (!paragTemplate)
            return console.log('No template paragraph or run were found !');
        this.setTextLanguage(paragTemplate); //We amend the language element to the "w:pPr" or "r:pPr" child elements of paragTemplate
        const self = this;
        text?.split('\n')
            .forEach((parag, index) => editParagraph(parag, index));
        function editParagraph(parag, index) {
            let textElement;
            if (index < 1)
                textElement = self.getXMLElements(paragTemplate, self.tags.text, index);
            else
                textElement = self.appendParagraph(paragTemplate, sdtContent); //We pass sdtContent as parent argumself
            if (!textElement)
                return console.log('No textElement was found !');
            textElement.textContent = parag;
        }
    }
    /**
   * Finds a "w:pPr" XML element (property element) which is a child of the XML parent element passed as argument. If does not find it, it looks for a "w:rPr" XML element. When it finds either a "w:pPr" or a "w:rPr" element, it appends a "w:lang" element to it, and sets its "w:val" attribute to the language passed as "lang"
   * @param {Element} parent - the XML element containing the paragraph or the run for which we want to set the language.
   * @returns {Element | undefined} - the "w:pPr" or "w:rPr" property XML element child of the parent element passed as argument
   */
    setTextLanguage(parent) {
        const pPr = this.getXMLElements(parent, this.Pr(this.tags.paragraph), 0) ||
            this.getXMLElements(parent, this.Pr(this.tags.run), 0);
        if (!pPr)
            return;
        pPr
            .appendChild(this.createXMLElement(this.tags.lang)) //appending a "w:lang" element
            .setAttributeNS(this.schema(), 'val', `${this.lang.toLowerCase()}-${this.lang.toUpperCase()}`); //setting the "w:val" attribute of "w:lang" to the appropriate language like "fr-FR"
        return pPr;
    }
}
class officeJS {
    constructor() {
        this.GRAPH_API_BASE_URL = "https://graph.microsoft.com/v1.0/me/drive/root:/";
    }
    async editDocumentWordJSAPI(id, accessToken, data, controlsData) {
        if (!id || !accessToken || !data)
            return;
        const graph = new GraphAPI(accessToken);
        await Word.run(async (context) => {
            // Open the document by downloading its content
            const fileResponse = await fetch(`${this.GRAPH_API_BASE_URL}/items/${id}/content`, {
                method: "GET",
                headers: {
                    "Authorization": `Bearer ${accessToken}`
                }
            });
            if (!fileResponse.ok)
                throw new Error("Failed to retrieve document");
            const blob = await fileResponse.blob();
            const base64File = await convertBlobToBase64(blob);
            const doc = context.application.createDocument(base64File);
            console.log("Word document opened for editing:", document);
            const tables = doc.body.tables;
            const contentControls = doc.body.contentControls;
            context.load(tables);
            context.load(contentControls);
            await context.sync();
            const table = tables.items[0];
            if (!table)
                return;
            data.forEach(dataRow => table.addRows("End", 1, [dataRow]));
            await editRichTextContentControls();
            async function editRichTextContentControls() {
                if (!controlsData || contentControls)
                    return;
                controlsData.forEach(control => edit(control));
                async function edit(control) {
                    const [title, text] = control;
                    const field = contentControls.getByTitle(title).getFirst();
                    if (!field)
                        return;
                    context.load(field);
                    await context.sync();
                    if (!text)
                        field.delete(false);
                    else
                        field.insertText(text, 'Replace');
                    await context.sync();
                    return field;
                }
            }
        });
    }
    async addEntry(tableName, rows) {
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
        function getNewRow(columns) {
            const newRow = Array(columns).map(el => '');
            const inputs = Array.from(document.getElementsByTagName('input')).filter(input => input.dataset.index);
            console.log('inputs = ', inputs);
            if (inputs.length < 1)
                return;
            inputs.forEach(input => {
                const index = Number(input.dataset.index);
                let value = input.value;
                if (input.type === 'number')
                    value = parseFloat(value);
                else if (input.type === 'date' && input.valueAsDate)
                    //@ts-ignore
                    value = getDateString(input.valueAsDate);
                else if (input.type === 'time' && input.valueAsDate)
                    value = [input.valueAsDate?.getHours().toString().padStart(2, '0'), input.valueAsDate?.getMinutes().toString().padStart(2, '0'), '00'].join(':');
                newRow[index] = value;
            });
            console.log('newRow = ', newRow);
            return [newRow];
            function convertTo24HourFormat(time12h) {
                const [time, modifier] = time12h.split(' ');
                let [hours, minutes] = time.split(':');
                if (hours === '12')
                    hours = '00';
                if (modifier === 'PM')
                    hours = String(parseInt(hours, 10) + 12);
                return `${hours}:${minutes}:00`;
            }
        }
    }
}
;
class blob {
    // Utility function: Convert Blob to Base64
    async blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result.toString().split(",")[1]);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    }
    // Utility function: Convert Base64 to Blob
    base64ToBlob(base64) {
        const byteCharacters = atob(base64);
        const byteNumbers = new Array(byteCharacters.length).fill(0).map((_, i) => byteCharacters.charCodeAt(i));
        const byteArray = new Uint8Array(byteNumbers);
        return new Blob([byteArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    }
}
;
/**
 * Returns the Word file name by which the newly issued invoice will be saved on OneDrive
 * @param {string} clientName - The name of the client for which the invoice will be issued
 * @param {string} matters - The matters included in the invoice
 * @param {string} invoiceNumber - The invoice serial number
 * @returns {string} - The name of the Word file to be saved
 */
function getInvoiceFileName(clientName, matters, invoiceNumber) {
    // return 'test file name for now.docx'
    return `${clientName}_Facture_${Array.from(matters).join('&')}_No.${invoiceNumber.replace('/', '@')}.docx`
        .replaceAll('/', '_')
        .replaceAll('"', '')
        .replaceAll("\\", '');
}
;
function getInvoiceNumber(date) {
    const padStart = (n) => n.toString().padStart(2, '0');
    return `${date.getFullYear() - 2000}${padStart(date.getMonth() + 1)}${padStart(date.getDate())}/${padStart(date.getHours())}${padStart(date.getMinutes())}`;
}
;
/**
 * Returns any date in the ISO format (YYY-MM-DD) accepted by Excel
 * @param {Date} date - the Date that we need to convert to ISO format
 * @returns {string} - The date in ISO format
 */
function getISODate(date) {
    //@ts-ignore
    return [date?.getFullYear(), date?.getMonth() + 1, date?.getDate()].map(el => el.toString().padStart(2, '0')).join('-');
}
;
/**
 * Returns the date in a string formated like: "DD/MM/YYYY"
 */
function getDateString(date) {
    return [date.getDate(), date.getMonth() + 1, date.getFullYear()]
        .map(el => el.toString().padStart(2, '0'))
        .join('/');
}
;
/**
 * Returns the value from a time input as a number matching the Excel time format (which is a fraction of the day)
 * @param {HTMLInputElement[]} inputs - If a single input is passed, it will return the Excel formatted time value from this input or 0. If 2 inputs are passed, it will return the total time by calculting the difference between the second input and the first input in the array
 * @returns {number} - The time as a number matching the Excel time format
 */
function getTime(inputs) {
    const day = (1000 * 60 * 60 * 24);
    if (inputs.length < 2 && inputs[0])
        return inputs[0].valueAsNumber / day || 0;
    const from = inputs[0]?.valueAsNumber; //this gives the time in milliseconds
    const to = inputs[1]?.valueAsNumber;
    if (!from || !to)
        return 0;
    const quarter = 15 * 60 * 1000; //quarter of an hour
    let time = to - from;
    time = Math.round(time / quarter) * quarter; //We are rounding the time by 1/4 hours
    time = time / day;
    if (time < 0)
        time = (to + day - from) / day; //It means we started on one day and finished the next day 
    return time;
}
;
/**
 * Helper function to convert Blob to Base64
 *  */
function convertBlobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result.toString().split(",")[1]); // Extract base64 part
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}
;
class MSAL {
    constructor(clientId, redirectUri, msalConfig, scopes = ["Files.ReadWrite"]) {
        this.clientId = '';
        this.redirectUri = '';
        this.loginRequest = { scopes: [''] };
        this.clientId = clientId;
        this.redirectUri = redirectUri;
        this.loginRequest.scopes = scopes;
        //@ts-expect-error
        this.msalInstance = new msal.PublicClientApplication(msalConfig);
    }
    async getTokenWithMSAL() {
        if (!this.clientId || !this.redirectUri || !this.msalInstance)
            return;
        return await this.acquireToken();
    }
    ;
    // Function to check existing authentication context
    async acquireToken() {
        try {
            const account = this.msalInstance.getAllAccounts()[0];
            if (account) {
                return await this.acquireTokenSilently(account);
            }
            else {
                return await this.loginWithPopup();
            }
        }
        catch (error) {
            console.error("Failed to acquire token from acquireToken(): ", error);
        }
    }
    // Function to get access token silently
    async acquireTokenSilently(account) {
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
        }
        catch (error) {
            console.error("Token silent acquisition error:", error);
        }
    }
    ;
    async loginWithPopup() {
        try {
            const loginResponse = await this.msalInstance.loginPopup(this.loginRequest);
            console.log('loginResponse = ', loginResponse);
            this.msalInstance.setActiveAccount(loginResponse.account);
            const tokenResponse = await this.msalInstance.acquireTokenSilent({
                account: loginResponse.account,
                scopes: ["Files.ReadWrite"]
            });
            console.log("Token acquired from loginWithPopup: ", tokenResponse.accessToken);
            return tokenResponse.accessToken;
        }
        catch (error) {
            console.error("Error acquiring token from loginWithPopup(): ", error);
            //@ts-ignore
            if (error instanceof InteractionRequiredAuthError) {
                // Fallback to popup if silent token acquisition fails
                const response = await this.msalInstance.acquireTokenPopup({
                    scopes: ["Files.ReadWrite"]
                });
                console.log("Token acquired via popup:", response.accessToken);
                return response.accessToken;
            }
        }
    }
    async credentitalsToken(tenantId) {
        const msalConfig = {
            auth: {
                clientId: this.clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
                //clientSecret: clientSecret,
            }
        };
        //@ts-ignore
        const cca = new msal.application.ConfidentialClientApplication(msalConfig);
        const tokenRequest = {
            scopes: ["Files.ReadWrite"],
        };
        try {
            const response = await cca.acquireTokenByClientCredential(tokenRequest);
            return response.accessToken;
        }
        catch (error) {
            console.log('Error acquiring Token: ', error);
            return null;
        }
    }
    async getOfficeToken() {
        try {
            //@ts-ignore
            return await OfficeRuntime.auth.getAccessToken();
        }
        catch (error) {
            console.log("Error : ", error);
        }
    }
    async getTokenWithSSO(email, tenantId) {
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
        }
        catch (error) {
            console.error("SSO silent authentication failed:", error);
            return null;
        }
    }
    openLoginWindow() {
        const loginUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${this.clientId}&response_type=token&redirect_uri=${this.redirectUri}&scope=https://graph.microsoft.com/.default`;
        // Open in a new window (only works if triggered by user action)
        const authWindow = window.open(loginUrl, "_blank", "width=500,height=600");
        if (!authWindow) {
            console.error("Popup blocked! Please allow popups.");
        }
    }
    // Function to handle login and acquire token
    async loginAndGetToken() {
        const msalConfig = {
            auth: {
                clientId: this.clientId,
                authority: "https://login.microsoftonline.com/common",
                redirectUri: this.redirectUri
            },
            cache: {
                cacheLocation: "ExcelInvoicing", // Specify cache location
                storeAuthStateInCookie: true // Set this to true for IE 11
            }
        };
        return await acquire(this);
        async function acquire(self) {
            try {
                const response = await self.msalInstance.handleRedirectPromise();
                if (response !== null) {
                    console.log("Login successful:", response);
                    return response.accessToken;
                }
                const accounts = self.msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    const tokenResponse = await self.msalInstance.acquireTokenSilent({
                        account: accounts[0],
                        scopes: ["https://graph.microsoft.com/.default"]
                    });
                    console.log("Token acquired silently:", tokenResponse.accessToken);
                    return tokenResponse.accessToken;
                }
            }
            catch (error) {
                console.error("Error acquiring token:", error);
                //@ts-ignore
                if (error instanceof msal.InteractionRequiredAuthError) {
                    self.msalInstance.acquireTokenRedirect({
                        scopes: ["https://graph.microsoft.com/.default"]
                    });
                }
            }
        }
        // Function to handle redirect response
        async function handleRedirectResponse(self) {
            try {
                const authResult = await self.msalInstance.handleRedirectPromise();
                if (authResult && authResult.accessToken) {
                    console.log("Access token:", authResult.accessToken);
                    return authResult.accessToken;
                }
            }
            catch (error) {
                console.error("Redirect handling error:", error);
            }
            return undefined;
        }
    }
}
function sortByColumn(data, columnIndex) {
    return data.slice().sort((a, b) => {
        const valA = a[columnIndex];
        const valB = b[columnIndex];
        if (typeof valA === "number" && typeof valB === "number") {
            return valA - valB; // Numeric sorting
        }
        return String(valA).localeCompare(String(valB)); // String sorting
    });
}
;
function getInputByIndex(inputs, index) {
    return inputs.find(input => Number(input.dataset.index) === index);
}
/**
 * Returns the dataset.index value of the input as a number
 * @param {HTMLInputElement} input - the input with a dataset.index attribute
 * @returns {number} - the dataset.index value of the input as a number
 */
function getIndex(element) {
    return Number(element?.dataset.index);
}
function settings() {
    const form = byID();
    if (!form)
        return;
    form.innerHTML = '';
    const inputs = [
        {
            label: 'Accounts Workbook Path: ',
            name: 'excelPath'
        },
        {
            label: 'Leases Workbook Path: ',
            name: 'leasesPath'
        },
        {
            label: 'Word Template Path: ',
            name: 'templatePath'
        },
        {
            label: 'Destination Folder: ',
            name: 'destinationFolder'
        },
        {
            label: 'Table Name: ',
            name: 'tableName'
        },
    ];
    inputs.forEach((el, index) => {
        const label = document.createElement('label');
        label.innerText = el.label;
        form.appendChild(label);
        const input = document.createElement('input');
        input.classList.add('field');
        input.value = localStorage.getItem(el.name) || 'not found';
        input.dataset.index = index.toString();
        input.onchange = () => set(input, el.label, el.name);
        form.appendChild(input);
        function set(input, label, name) {
            if (!confirm(`Are you sure you want to change the ${label} localStorage value to + ${input.value}?`))
                return;
            localStorage.setItem(name, input.value);
            alert(`${label} has been updated`);
        }
    });
    (function homeBtn() {
        showMainUI(true);
    })();
}
function spinner(show) {
    if (!show)
        return document.querySelector('.spinner')?.remove();
    const form = document.getElementById('form');
    if (!form)
        return;
    const spinner = document.createElement('div');
    spinner.classList.add('spinner');
    form.appendChild(spinner);
}
//# sourceMappingURL=main.js.map