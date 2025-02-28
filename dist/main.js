"use strict";
if (!localStorage.excelPath)
    localStorage.excelPath = prompt('Please provide the OneDrive full path (including the file name and extension) for the Excel Workbook', "Legal/Mon Cabinet d'Avocat/Comptabilité/Comptabilité de Mon Cabinet_15 10 2023.xlsm");
const workbookPath = localStorage.excelPath || alert('The excel Workbook path is not valid');
if (!localStorage.templatePath)
    localStorage.templatePath = prompt('Please provide the OneDrive full path (including the file name and extension) for the Word template', "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/FactureTEMPLATE [NE PAS MODIFIDER].docx");
const templatePath = localStorage.templatePath || alert('The template path is not valid or is missing');
if (!localStorage.tableName)
    localStorage.tableName = prompt('Please provide the name of the Excel Table where the invoicing data is stored', 'LivreJournal');
const tableName = localStorage.tableName || alert('The table name is not valid or is issing');
if (!localStorage.destinationFolder)
    localStorage.destinationFolder = prompt('Please provide the OneDrive path where the issued invoices will be stored', "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/Clients");
const destinationFolder = localStorage.destinationFolder || alert('the destination folder path is missing or not valid');
const tenantId = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
var TableRows, accessToken;
/* Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Excel-specific initialization code goes here
    console.log("Excel is ready!");

    loadMsalScript();
  }
}); */
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
        inputs.forEach(input => input?.addEventListener('focusout', async () => await inputOnChange(input), { passive: true }));
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
                if (Number(index) === 0)
                    createDataList(input?.id, clientUniqueValues); //We create a unique values dataList for the 'Client' input
                return input;
            });
        }
        ;
        async function inputOnChange(input, unfilter = false) {
            const index = Number(input.dataset.index);
            if (index === 0)
                unfilter = true; //If this is the 'Client' column, we remove any filter from the table;
            //We filter the table accordin to the input's value and return the visible cells
            const visibleCells = await filterTable(tableName, [{ column: index, value: getArray(input.value) }], unfilter);
            if (visibleCells.length < 1)
                return alert('There are no visible cells in the filtered table');
            //We create (or update) the unique values dataList for the next input 
            const nextInput = getNextInput(input);
            if (!nextInput)
                return;
            createDataList(nextInput?.id || '', getUniqueValues(Number(nextInput.dataset.index), visibleCells));
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
function createDataList(id, uniqueValues, multiple = false) {
    //const uniqueValues = Array.from(new Set(visible.map(row => row[i])));
    if (!id || !uniqueValues || uniqueValues.length < 1)
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
    if (multiple && uniqueValues.length > 1)
        addOption(uniqueValues.join(', '));
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
    const array = value.replaceAll(', ', ',')
        .replaceAll(' ,', ',')
        .split(',');
    return array.filter((el) => el);
}
async function generateInvoice() {
    const inputs = Array.from(document.getElementsByName('input'));
    if (!inputs)
        return;
    const lang = inputs.find(input => input.type === 'checkbox' && input.checked === true)?.id.slice(0, 3).toUpperCase() || 'FR';
    const discount = parseInt(inputs.find(input => input.id = 'discount')?.value || '0%');
    const visible = await filterTable(tableName, undefined, false);
    const invoiceDetails = {
        number: getInvoiceNumber(new Date()),
        clientName: visible.map(row => String(row[0]))[0] || 'CLIENT',
        matters: (getUniqueValues(1, visible)).map(el => String(el)),
        adress: (getUniqueValues(15, visible)).map(el => String(el)),
        lang: lang
    };
    const filePath = `${destinationFolder}/${getInvoiceFileName(invoiceDetails.clientName, invoiceDetails.matters, invoiceDetails.number)}`;
    const [rows, totals] = getRowsData(visible, discount, lang);
    await createAndUploadXmlDocument(await getAccessToken() || '', templatePath, filePath, 'Invoice', lang, rows, getContentControlsValues(invoiceDetails, new Date()));
}
/**
 * Returns a string[][] representing the rows to be inserted in the Word table containing the invoice details
 * @param {string[][]} tableData - The filtered Excel rows from which the data will be extracted and put in the required format
 * @param {string} lang - The language in which the invoice will issued
 * @returns {string[][]} - the rows to be added to the table. Each row has 4 elements
 */
function getRowsData(tableData, discount, lang) {
    const labels = {
        totalFees: {
            nature: 'Honoraire',
            FR: 'Total honoraires',
            EN: 'Total Fees'
        },
        totalExpenses: {
            nature: 'Débours/Dépens, Rétrocession d\'honoraires',
            FR: 'Total débours et frais',
            EN: 'Total Expenses'
        },
        totalPayments: {
            nature: 'Provision/Règlement, Réduction',
            FR: 'Total provisions reçues',
            EN: 'Total Payments'
        },
        totalTimeSpent: {
            nature: '',
            FR: 'Total des heures facturables (hors prestations facturées au forfait) ',
            EN: 'Total billable hours (other than lump-sum billed services)'
        },
        totalDue: {
            nature: '',
            FR: 'Montant dû',
            EN: 'Total Due'
        },
        totalReinbursement: {
            nature: '',
            FR: 'A rembourser',
            EN: 'To reimburse'
        },
        FeesDeduction: {
            nature: '',
            FR: 'Remise',
            EN: 'Discount'
        },
        totalFeesAfterDeduction: {
            nature: '',
            FR: 'Total honoraires après remise',
            EN: 'Total fee after discount'
        },
        discountDescription: {
            nature: '',
            FR: `${discount.toString()}% de remise sur les honoraires`,
            EN: `${discount.toString()}% discount on accrued fees`
        },
        hourlyBilled: {
            nature: '',
            FR: 'facturation au temps passé\u00A0:',
            EN: 'hourly billed:',
        },
        hourlyRate: {
            nature: '',
            FR: 'au taux horaire de\u00A0:',
            EN: 'at an hourly rate of:',
        },
        decimal: {
            nature: '',
            FR: ',',
            EN: '.'
        },
    };
    const amount = 9, vat = 10, hours = 7, rate = 8, nature = 2, descr = 14; //Indexes of the Excel table columns from which we extract the date 
    const data = tableData.map(row => {
        const date = dateFromExcel(Number(row[3]));
        const time = getTimeSpent(Number(row[hours]));
        let description = `${String(row[nature])} : ${String(row[descr])}`; //Column Nature + Column Description;
        //If the billable hours are > 0, we add to the description: time spent and hourly rate
        if (time)
            description += ` (${labels.hourlyBilled[lang]} ${time}, ${labels.hourlyRate[lang]} ${Math.abs(row[rate]).toString()}\u00A0€).`;
        const rowValues = [
            [date.getDate(), date.getMonth() + 1, date.getFullYear()].map(el => el.toString().padStart(2, '0')).join('/'), //Column Date
            description,
            getAmountString(row[amount] * -1), //Column "Amount": we inverse the +/- sign for all the values 
            getAmountString(Math.abs(row[vat])), //Column VAT: always a positive value
        ];
        return rowValues;
    });
    pushTotalsRows();
    return [data, Object.values(labels).map(value => value[lang])];
    function pushTotalsRows() {
        //Adding rows for the totals of the different categories and amounts
        const totalFee = getTotals(amount, labels.totalFees.nature);
        const totalFeeVAT = getTotals(vat, labels.totalFees.nature);
        const totalPayments = getTotals(amount, labels.totalPayments.nature);
        const totalPaymentsVAT = getTotals(vat, labels.totalPayments.nature);
        const totalExpenses = getTotals(amount, labels.totalExpenses.nature);
        const totalExpensesVAT = getTotals(vat, labels.totalExpenses.nature);
        const totalTimeSpent = getTotals(hours, null); //by passing the nature = null, we do not filter the "Total Time" column by any crieteria. We will get the sum of all the column.
        const totalDue = totalFee + totalExpenses + totalPayments;
        const totalDueVAT = totalFeeVAT + totalExpensesVAT;
        const insert = (sum) => Math.abs(sum) > 0;
        (function subTotalsRows() {
            push(insert(totalFee), labels.totalFees, totalFee, totalFeeVAT);
            push(insert(totalExpenses), labels.totalExpenses, totalExpenses, totalExpensesVAT);
            push(insert(totalPayments), labels.totalPayments, Math.abs(totalPayments), totalPaymentsVAT);
            push(insert(totalTimeSpent), labels.totalTimeSpent, Math.abs(totalTimeSpent), NaN); //!We don't pass the vat argument in order to get the corresponding cell of the Word table empty
        })();
        (function dueRow() {
            if (!discount)
                return totalDueRow(totalDue, totalDueVAT); //If there is no discount to be applied on the fees, we return the "Total Due" row;
            const feeDiscount = (fee) => fee * (discount / 100); //returns the amount to be deducted from the fees or from the VAT on the fees as a negative number
            const deduction = feeDiscount(totalFee);
            const deductionVAT = feeDiscount(totalFeeVAT);
            if (!insert(deduction))
                return totalDueRow(totalDue, totalDueVAT); //If the total fee is 0 for whatever reason, it means that deduction will be = 0. In such case we return the "Total Due" row as if there were no deduction applied.
            push(true, labels.FeesDeduction, deduction * -1, deductionVAT * -1, labels.discountDescription); //We add a description for this row. This is the amount that will be deducted from the fee and from the fee pplied on the fee
            push(true, labels.totalFeesAfterDeduction, (totalFee - deduction), (totalFeeVAT - deductionVAT)); //This is the amount that will be deducted from the fee and from the fee pplied on the fee. 
            totalDueRow(totalDue - deduction, totalDueVAT - deductionVAT);
            addDiscountRowToExcel(deduction, deductionVAT);
        })();
        function addDiscountRowToExcel(amount, vat) {
            const newRow = tableData
                .find(row => row[2] === 'Honoraire')
                ?.map((cell, index) => {
                if ([0, 1, 11, 15].includes(index))
                    return cell;
                else if (index === 2)
                    return 'Réduction';
                else if ([3, 4].includes(index))
                    return getISODate(new Date());
                else if (index === 9)
                    return amount; //When the amount represents a fee or expense billed to the client,  it is a negative value. That's why in this case we will add it as a positive value in order to deduct the it from the already billed fees 
                else if (index === 10)
                    return vat * -1; //VAT is usually added as a positive value, but since we want to deduct this amount from the total VAT, we will add it as a negative value
                else
                    undefined;
            });
            addNewEntry(true, newRow);
        }
        function totalDueRow(total, vat) {
            total >= 0 ? push(true, labels.totalDue, total, vat)
                : push(true, labels.totalReinbursement, total, vat);
        }
        function push(insert, label, amount, vat, description) {
            if (!insert)
                return;
            data.push([
                label?.[lang] || '',
                description?.[lang] || '',
                label === labels.totalTimeSpent ? getTimeSpent(amount) : getAmountString(amount),
                getAmountString(Math.abs(vat)), //Column VAT: always a positive value
            ]);
        }
        function getTotals(index, nature) {
            const total = tableData.filter(row => nature ? nature.split(', ').includes(row[2]) : row === row)
                .map(row => Number(row[index]));
            let sum = 0;
            for (let i = 0; i < total.length; i++) {
                sum += total[i];
            }
            return sum * -1; //We reverse the sign of any other amount
        }
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
            text: { FR: 'Paris le ', EN: 'Paris on ' }[invoice.lang] || '',
        },
        date: {
            title: 'RTInvoiceDate',
            text: [date.getDate(), date.getMonth() + 1, date.getFullYear()].map(el => el.toString().padStart(2, '0')).join('/'),
        },
        numberLabel: {
            title: 'LabelInvoiceNumber',
            text: { FR: 'Facture n°\u00A0:', EN: 'Invoice No.:' }[invoice.lang] || '',
        },
        number: {
            title: 'RTInvoiceNumber',
            text: invoice.number,
        },
        subjectLable: {
            title: 'LabelSubject',
            text: { FR: 'Affaires\u00A0: ', EN: 'Matters: ' }[invoice.lang] || '',
        },
        subject: {
            title: 'RTMatter',
            text: invoice.matters.join(' & '),
        },
        fee: {
            title: 'LabelTableHeadingHonoraire',
            text: { FR: 'Honoraire/Débours', EN: 'Fees/Expenses' }[invoice.lang] || '',
        },
        amount: {
            title: 'LabelTableHeadingMontantTTC',
            text: { FR: 'Montant TTC', EN: 'Amount VAT Included' }[invoice.lang] || '',
        },
        vat: {
            title: 'LabelTableHeadingTVA',
            text: { FR: 'TVA', EN: 'VAT' }[invoice.lang] || '',
        },
        disclaimer: {
            title: 'LabelDisclamer' + ['French', 'English'].find(el => !el.toUpperCase().startsWith(invoice.lang)) || 'French',
            text: '',
        },
        clientName: {
            title: 'RTClient',
            text: invoice.clientName,
        },
        adress: {
            title: 'RTClientAdresse',
            text: invoice.adress.join(' & '),
        },
    };
    return Object.keys(fields).map(key => [fields[key].title, fields[key].text]);
}
function getUniqueValues(index, array) {
    if (!array)
        array = [];
    return Array.from(new Set(array.map(row => row[index])))
        .map(el => el); //we remove empty strings/values
}
;
/**
 * Returns all the rows of an Excel table in a workbook stored on OneDrive, using the Graph API
 * @param {string} accessToken - the access token
 * @param {string} filePath - file path (folder + file nam) of the file to be fetched
 * @param {string} tableName - Name of the table to be fetched
 * @param {boolean} range - Its default value is true. If true, it returns all the rows of the table including the title, otherwise, it returns only the body of the table
 * @param {boolean} columns - If true it will return the columns
 * @returns {any[][] | number | void} - All the rows (including the title) of the Excel table
 */
async function fetchExcelTableWithGraphAPI(accessToken, filePath, tableName, range = true, columns) {
    if (!accessToken)
        accessToken = await getAccessToken() || '';
    let endPoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/`;
    if (range)
        endPoint = endPoint += 'range';
    if (!range)
        endPoint = endPoint += 'rows';
    else if (columns)
        endPoint += 'columns';
    const response = await fetch(endPoint, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!response.ok) {
        alert(`Error fetching row count: ${await response.text()}`);
        throw new Error(`Error fetching row count: ${await response.text()}`);
    }
    ;
    const data = await response.json();
    if (range)
        return data.values;
    else
        return data.value.map((v) => v.values);
}
async function clearFilterExcelTableGraphAPI(filePath, tableName, accessToken) {
    // First, clear filters on the table (optional step)
    await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/clearFilters`, {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
        }
    });
}
/**
 * Returns a blob from a file stored on OneDrive, using the Graph API and the file path
 * @param {string} accessToken
 * @param {string} filePath
 * @returns {Blob} - A blob of the fetched file, if successful
 */
async function fetchFileFromOneDriveWithGraphAPI(accessToken, filePath) {
    const fileUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/content`;
    const response = await fetch(fileUrl, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!response.ok)
        throw new Error("Failed to fetch Word template");
    return await response.blob(); // Returns the Word template as a Blob
}
/**
 * Uploads a file blob to OneDrive using the Graph API
 * @param {Blob } blob
 * @param {string} filePath
 * @param {string} accessToken
 */
async function uploadFileToOneDriveWithGraphAPI(blob, filePath, accessToken) {
    const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/content`;
    const response = await fetch(endpoint, {
        method: 'PUT',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // Correct MIME type for Word docs
        },
        body: blob, // Use the template's content as the new document's content
    });
    response.ok ? alert('succefully uploaded the new file') : console.log('failed to upload the file to onedrive error = ', await response.json());
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
async function getExcelTableRowsCountViaGraphAPI(filePath, tableName, accessToken) {
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/rows/$count`;
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
function getInvoiceNumber(date) {
    const padStart = (n) => n.toString().padStart(2, '0');
    return (date.getFullYear() - 2000).toString() + padStart(date.getMonth() + 1) + padStart(date.getDate()) + '/' + padStart(date.getHours()) + padStart(date.getMinutes());
}
/**
 * Returns any date in the ISO format (YYY-MM-DD) accepted by Excel
 * @param {Date} date - the Date that we need to convert to ISO format
 * @returns {string} - The date in ISO format
 */
function getISODate(date) {
    //@ts-ignore
    return [date?.getFullYear(), date?.getMonth() + 1, date?.getDate()].map(el => el.toString().padStart(2, '0')).join('-');
}
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
async function editDocumentWordJSAPI(id, accessToken, data, controlsData) {
    if (!id || !accessToken || !data)
        return;
    await Word.run(async (context) => {
        // Open the document by downloading its content
        const fileResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${id}/content`, {
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
async function addEntry(tableName, rows) {
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
                value = [String(input.valueAsDate?.getDay()).padStart(2, '0'), String(input.valueAsDate.getMonth() + 1).padStart(2, '0'), String(input.valueAsDate?.getFullYear())].join('/');
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
/*
// Create a new Word document based on a template and populate it with filtered data
async function createWordDocument(filtered: any[][]) {
  return console.log("filtered = ", filtered);
 
  await Word.run(async (context) => {
    const templateUrl = "https://your-onedrive-path/template.docx";
    const newDoc = context.application.createDocument(templateUrl);
    await context.sync();
 
    const table = newDoc.body.tables.getFirst();
    //const filteredData = await getFilteredData();
 
    //filtered.forEach(el) => {
    // table.(index + 1, row);
    //});
 
    await context.sync();
 
    const saveUrl = "https://your-onedrive-path/newDocument.docx";
    await newDoc.saveAs(saveUrl);
  });
}*/
function getTokenWithMSAL(clientId, redirectUri, msalConfig) {
    if (!clientId || !redirectUri || !msalConfig)
        return;
    //@ts-expect-error
    const msalInstance = new msal.PublicClientApplication(msalConfig);
    const loginRequest = { scopes: ["Files.ReadWrite"] };
    return acquireToken();
    // Function to check existing authentication context
    async function acquireToken() {
        try {
            const account = msalInstance.getAllAccounts()[0];
            if (account) {
                return acquireTokenSilently(account);
            }
            else {
                return loginWithPopup();
                //return loginAndGetToken();
                //openLoginWindow()
                //return getOfficeToken()
                //return getTokenWithSSO('minabibawi@gmail.com')
                //return credentitalsToken()
            }
        }
        catch (error) {
            console.error("Failed to acquire token from acquireToken(): ", error);
        }
    }
    async function loginWithPopup() {
        try {
            const loginResponse = await msalInstance.loginPopup(loginRequest);
            console.log('loginResponse = ', loginResponse);
            msalInstance.setActiveAccount(loginResponse.account);
            const tokenResponse = await msalInstance.acquireTokenSilent({
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
                const response = await msalInstance.acquireTokenPopup({
                    scopes: ["Files.ReadWrite"]
                });
                console.log("Token acquired via popup:", response.accessToken);
                return response.accessToken;
            }
        }
    }
    async function credentitalsToken(tenantId) {
        const msalConfig = {
            auth: {
                clientId: clientId,
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
    async function getOfficeToken() {
        try {
            //@ts-ignore
            return await OfficeRuntime.auth.getAccessToken();
        }
        catch (error) {
            console.log("Error : ", error);
        }
    }
    async function getTokenWithSSO(email, tenantId) {
        const msalConfig = {
            auth: {
                clientId: clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
                redirectUri: redirectUri,
                navigateToLoginRequestUrl: true,
            },
            cache: {
                cacheLocation: "ExcelAddIn",
                storeAuthStateInCookie: true
            }
        };
        try {
            //@ts-ignore
            const response = await msal.PublicClientApplication(msalConfig).ssoSilent({
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
    function openLoginWindow() {
        const loginUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}&scope=https://graph.microsoft.com/.default`;
        // Open in a new window (only works if triggered by user action)
        const authWindow = window.open(loginUrl, "_blank", "width=500,height=600");
        if (!authWindow) {
            console.error("Popup blocked! Please allow popups.");
        }
    }
    // Function to handle login and acquire token
    async function loginAndGetToken() {
        const msalConfig = {
            auth: {
                clientId: clientId,
                authority: "https://login.microsoftonline.com/common",
                redirectUri: redirectUri
            },
            cache: {
                cacheLocation: "ExcelInvoicing", // Specify cache location
                storeAuthStateInCookie: true // Set this to true for IE 11
            }
        };
        //@ts-ignore
        const msalInstance = new msal.PublicClientApplication(msalConfig);
        return await acquire();
        async function acquire() {
            try {
                const response = await msalInstance.handleRedirectPromise();
                if (response !== null) {
                    console.log("Login successful:", response);
                    return response.accessToken;
                }
                const accounts = msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    const tokenResponse = await msalInstance.acquireTokenSilent({
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
                    msalInstance.acquireTokenRedirect({
                        scopes: ["https://graph.microsoft.com/.default"]
                    });
                }
            }
        }
        return;
        try {
            const loginRequest = {
                scopes: ["Files.ReadWrite"] // OneDrive scopes
            };
            await msalInstance.loginRedirect(loginRequest);
            return handleRedirectResponse();
        }
        catch (error) {
            console.error("Login error:", error);
            return undefined;
        }
        // Function to handle redirect response
        async function handleRedirectResponse() {
            try {
                const authResult = await msalInstance.handleRedirectPromise();
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
    // Function to get access token silently
    async function acquireTokenSilently(account) {
        try {
            const tokenRequest = {
                account: account,
                scopes: ["Files.ReadWrite"], // OneDrive scopes
            };
            const tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest);
            if (tokenResponse && tokenResponse.accessToken) {
                console.log("Token acquired silently :", tokenResponse.accessToken);
                return tokenResponse.accessToken;
            }
        }
        catch (error) {
            console.error("Token silent acquisition error:", error);
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
function getInputByIndex(inputs, index) {
    return inputs.find(input => Number(input.dataset.index) === index);
}
/**
 * Returns the dataset.index value of the input as a number
 * @param {HTMLInputElement} input - the input with a dataset.index attribute
 * @returns {number} - the dataset.index value of the input as a number
 */
function getIndex(input) {
    return Number(input.dataset.index);
}
// Utility function: Convert Blob to Base64
async function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result.toString().split(",")[1]);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}
// Utility function: Convert Base64 to Blob
function base64ToBlob(base64) {
    const byteCharacters = atob(base64);
    const byteNumbers = new Array(byteCharacters.length).fill(0).map((_, i) => byteCharacters.charCodeAt(i));
    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}
function getInputValue(index, inputs) {
    return inputs.find(input => Number(input.dataset.index) === index)?.value || '';
}
function settings() {
    const form = document.getElementById('form');
    if (!form)
        return;
    form.innerHTML = '';
    const inputs = [
        {
            label: 'Workbook Path: ',
            name: 'excelPath'
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
}
//# sourceMappingURL=main.js.map