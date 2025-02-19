"use strict";
// Authentication
//const accessToken = getAccessToken();
function getAccessToken() {
    const clientId = "157dd297-447d-4592-b2d3-76b643b97132";
    const redirectUri = "https://mbibawi.github.io/ExcelInvoicingAddIn"; //!must be the same domain as the app
    const msalConfig = {
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
    return getTokenWithMSAL(clientId, redirectUri, msalConfig);
}
// Fetch OneDrive File by Path
async function fetchOneDriveFileByPath(filePathAndName, accessToken) {
    try {
        const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePathAndName}:/content`;
        const response = await fetch(url, {
            method: "GET",
            headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Content-Type": "application/json"
            }
        });
        const data = await response.arrayBuffer();
        return data;
    }
    catch (error) {
        console.error("Error fetching OneDrive file:", error);
    }
}
async function addNewEntry(add = false) {
    accessToken = await getAccessToken() || '';
    (async function show() {
        if (add)
            return;
        TableRows = await fetchExcelTable(accessToken, excelPath, tableName);
        if (!TableRows)
            return;
        showForm(TableRows[0]);
    })();
    (async function addEntry() {
        if (!add)
            return;
        const inputs = Array.from(document.getElementsByTagName('input')); //all inputs
        const nature = getInputByIndex(inputs, 2)?.value || '';
        const date = getInputByIndex(inputs, 3)?.valueAsDate || undefined;
        const debit = ['Honoraire', 'Débours/Dépens', 'Débours/Dépens non facturables', 'Rétrocession d\'honoraires'].includes(nature); //We check if we need to change the value sign 
        const row = inputs.map(input => {
            const index = getIndex(input);
            if ([3, 4].includes(index))
                return getISODate(date); //Those are the 2 date columns
            else if ([5, 6].includes(index))
                return getTime([input]); //time start and time end columns
            else if (index === 7)
                return getTime([getInputByIndex(inputs, 5), getInputByIndex(inputs, 6)]); //Total time column
            else if (debit && index === 9)
                return input.valueAsNumber * -1 || 0; //This is the amount if negative
            else if ([8, 9, 10].includes(index))
                return input.valueAsNumber || 0; //Hourly Rate, Amount, VAT
            else
                return input.value;
        });
        const stop = 'You must at least provide the client, matter, nature, date and the amount. If you provided a time start, you must provide a time end, and an hourly rate. Please review your fields';
        if (row.filter((el, i) => (i < 4 || i === 9) && !el).length > 0)
            return alert(stop); //if client name, matter, nature, date or amount are missing
        else if (row[5] && (!row[6] || !row[8]))
            return alert(stop); //if startTime is provided but without endTime or without hourly rate
        else if (row[6] && (!row[5] || !row[8]))
            return alert(stop); //if endTime is provided but without startTime or without hourly rate
        await addRowToExcelTable([row], TableRows.length - 2, excelFilePath, tableName, accessToken);
        [0, 1].forEach(async (index) => {
            await filterExcelTable(excelFilePath, tableName, TableRows[0][index], row[index].toString(), accessToken);
        });
        function getISODate(date) {
            //@ts-ignore
            return [date?.getFullYear(), date?.getMonth() + 1, date?.getDate()].map(el => el.toString().padStart(2, '0')).join('-');
        }
        function getTime(inputs) {
            const day = (1000 * 60 * 60 * 24);
            if (inputs.length === 1 && inputs[0])
                return inputs[0].valueAsNumber / day || 0;
            const from = inputs[0]?.valueAsNumber; //this gives the time in milliseconds
            const to = inputs[1]?.valueAsNumber;
            if (!from || !to)
                return 0;
            let time = (to - from) / day;
            if (time < 0)
                time = (to + day - from) / day; //It means we started on one day and finished the next day 
            return time;
        }
    })();
    function showForm(title) {
        const form = document.getElementById('form');
        if (!form)
            return;
        form.innerHTML = '';
        title.forEach((title, index) => {
            if (![4, 7, 15].includes(index))
                form.appendChild(createLable(title, index)); //We exclued the labels for "Total Time" and for "Year"
            form.appendChild(createInput(index));
        });
        for (const t of title) { //!We could not use for(let i=0; i<title.length; i++) because the await does not work properly inside this loop
        }
        ;
        (function addBtn() {
            const btnIssue = document.createElement('button');
            btnIssue.innerText = 'Add Entry';
            btnIssue.classList.add('button');
            btnIssue.onclick = () => addNewEntry(true);
            form.appendChild(btnIssue);
        })();
        function createLable(title, i) {
            const label = document.createElement('label');
            label.htmlFor = 'input' + i.toString();
            label.innerHTML = title + ':';
            return label;
        }
        function createInput(index) {
            const css = 'field';
            const input = document.createElement('input');
            const id = 'input' + index.toString();
            input.classList.add(css);
            input.id = id;
            input.name = id;
            input.autocomplete = "on";
            input.dataset.index = index.toString();
            input.type = 'text';
            if ([8, 9, 10].includes(index))
                input.type = 'number';
            else if (index === 3)
                input.type = 'date';
            else if ([5, 6].includes(index))
                input.type = 'time';
            else if ([4, 7, 15].includes(index))
                input.style.display = 'none'; //We hide those 3 columns: 'Total Time' and the 'Year' and 'Link to a File'
            else if ([0, 1, 2, 11, 12, 13, 16].includes(index)) {
                //We add a dataList for those fields
                input.setAttribute('list', input.id + 's');
                input.onchange = () => inputOnChange(index, TableRows.slice(1, -1), false);
                if (![1, 16].includes(index))
                    createDataList(input.id, getUniqueValues(index, TableRows.slice(1, -1), tableName)); //We don't create the data list for columns 'Matter' (1) and 'Adress' (16) because it will be created when the 'Client' field is updated
            }
            return input;
        }
    }
}
// Update Word Document
async function invoice(issue = false) {
    accessToken = await getAccessToken() || '';
    (async function show() {
        if (issue)
            return;
        TableRows = await fetchExcelTable(accessToken, excelPath, tableName);
        if (!TableRows)
            return;
        insertInvoiceForm(TableRows);
    })();
    (async function issueInvoice() {
        if (!issue)
            return;
        const inputs = Array.from(document.getElementsByTagName('input'));
        const criteria = inputs.filter(input => Number(input.dataset.index) >= 0);
        const lang = inputs.find(input => input.type === 'checkbox' && input.checked === true)?.dataset.language || 'FR';
        const filtered = filterExcelData(TableRows, criteria, lang);
        console.log('filtered table = ', filtered);
        const date = new Date();
        const invoice = {
            number: getInvoiceNumber(date),
            clientName: getInputValue(0, criteria),
            matters: getArray(getInputValue(1, criteria)),
            adress: Array.from(new Set(filtered.map(row => row[16]))),
            lang: lang
        };
        const contentControls = getContentControlsValues(invoice, date);
        const filePath = `${destinationFolder}/${newWordFileName(invoice.clientName, invoice.matters, invoice.number)}`;
        await createAndUploadXmlDocument(filtered, contentControls, accessToken, filePath, lang);
    })();
}
/**
 * Updates the data list of the other fields according to the value of the input that has been changed
 * @param {number} index - the dataset.index of the input that has been changed
 * @param {any[][]} table - the table that will be filtered
 * @param {boolean} invoice - If true, it means that we called the function in order to generate an invoice. If false, we called it in order to add a new entry in the table
 * @returns
 */
function inputOnChange(index, table, invoice) {
    let inputs = Array.from(document.getElementsByTagName('input'));
    if (invoice)
        inputs = inputs.filter(input => input.dataset.index && Number(input.dataset.index) < 3); //Those are all the inputs that serve to filter the table (first 3 columns only)
    else
        inputs = inputs.filter(input => [0, 1, 16].includes(getIndex(input))); //Those are all the inputs that have data lists associated with them
    const filledInputs = inputs
        .filter(input => input.value && getIndex(input) <= index)
        .map(input => getIndex(input)); //Those are all the inputs that the user filled with data
    const boundInputs = inputs.filter(input => getIndex(input) > index); //Those are the inputs for which we want to create  or update their data lists
    if (filledInputs.length < 1 || boundInputs.length < 1)
        return;
    boundInputs.forEach(input => input.value = '');
    const filtered = filterOnInput(inputs, filledInputs, table); //We filter the table based on the filled inputs
    if (filtered.length < 1)
        return;
    boundInputs.map(input => createDataList(input?.id, getUniqueValues(getIndex(input), filtered, tableName), invoice));
    if (invoice) {
        const nature = getInputByIndex(inputs, 2); //We get the nature input in order to fill automaticaly its values by a ', ' separated string
        if (!nature)
            return;
        nature.value = Array.from(document.getElementById(nature?.id + 's')?.children)?.map((option) => option.value).join(', ');
    }
    function filterOnInput(inputs, filled, table) {
        let filtered = table;
        for (let i = 0; i < filled.length; i++) {
            filtered = filtered.filter(row => row[filled[i]].toString() === getInputByIndex(inputs, filled[i])?.value);
        }
        return filtered;
    }
}
;
async function extractInvoiceData(lang) {
    const inputs = Array.from(document.getElementsByTagName('input'));
    const criteria = inputs.filter(input => input.dataset.index);
    //@ts-expect-error
    const tableData = filterRows(0, await fetchExcelData());
    return getRowsData(tableData, lang);
    function filterRows(i, tableData) {
        while (i < criteria.length) {
            const input = criteria[i];
            tableData = tableData.filter(row => row[Number(input.dataset.index)].toString() === input.value);
            i++;
        }
        return tableData;
    }
}
function dateFromExcel(excelDate) {
    const date = new Date((excelDate - 25569) * (60 * 60 * 24) * 1000); //This gives the days converted from milliseconds. 
    const dateOffset = date.getTimezoneOffset() * 60 * 1000; //Getting the difference in milleseconds
    return new Date(date.getTime() + dateOffset);
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
            text: { FR: 'Facturen n° : ', EN: 'Invoice No.:' }[invoice.lang] || '',
        },
        number: {
            title: 'RTInvoiceNumber',
            text: invoice.number,
        },
        subjectLable: {
            title: 'LabelSubject',
            text: { FR: 'Objet : ', EN: 'Subject: ' }[invoice.lang] || '',
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
            text: invoice.adress.join('/n'),
        },
    };
    return Object.keys(fields).map(key => [fields[key].title, fields[key].text]);
}
function getMSGraphClient(accessToken) {
    //@ts-expect-error
    return MicrosoftGraph.Client.init({
        //@ts-expect-error
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}
// Save Word Document to Another Location
async function saveWordDocumentToNewLocation(invoice, accessToken, originalFilePath = templatePath, newFilePath = destinationFolder) {
    const date = new Date();
    const fileName = newWordFileName(invoice.clientName, invoice.matters, invoice.number);
    newFilePath = `${destinationFolder}/${fileName}`;
    try {
        const fileContent = await fetchOneDriveFileByPath(originalFilePath, accessToken);
        const resp = await getMSGraphClient(accessToken).api(`/me/drive/root:${newFilePath}:/content`).put(fileContent);
        console.log('put response = ', resp);
        return newFilePath;
    }
    catch (error) {
        console.error("Error saving Word document:", error);
    }
}
function getNewExcelRow(inputs) {
    return inputs.map(input => {
        input.value;
    });
}
async function addRowToExcelTable(row, index, filePath, tableName, accessToken) {
    await clearFliter(); //We start by clearing the filter of the table, otherwise the insertion will fail
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/rows`;
    const body = {
        index: index,
        values: row,
    };
    const response = await fetch(url, {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });
    if (response.ok) {
        alert("Row added successfully!");
        return await response.json();
    }
    else {
        alert(`Error adding row: ${await response.text()}`);
    }
    async function clearFliter() {
        // First, clear filters on the table (optional step)
        await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/clearFilters`, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Content-Type": "application/json"
            }
        });
    }
}
async function filterExcelTable(filePath, tableName, columnName, filterValue, accessToken) {
    // Step 3: Apply filter using the column name
    const filterUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/columns/${columnName}/filter/apply`;
    const body = {
        criteria: {
            filterOn: "custom",
            criterion1: `="${filterValue}"`,
        }
    };
    const filterResponse = await fetch(filterUrl, {
        method: "POST",
        headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });
    if (filterResponse.ok) {
        alert(`Filter applied to column ${columnName} successfully!`);
    }
    else {
        alert(`Error applying filter: ${await filterResponse.text()}`);
    }
}
//# sourceMappingURL=pwaVersion.js.map