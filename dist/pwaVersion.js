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
        const amount = getInputByIndex(inputs, 9);
        const rate = getInputByIndex(inputs, 8)?.valueAsNumber;
        const debit = ['Honoraire', 'Débours/Dépens', 'Débours/Dépens non facturables', 'Rétrocession d\'honoraires'].includes(nature); //We check if we need to change the value sign 
        const row = inputs.map(input => {
            const index = getIndex(input);
            if ([3, 4].includes(index))
                return getISODate(date); //Those are the 2 date columns
            else if ([5, 6].includes(index))
                return getTime([input]); //time start and time end columns
            else if (index === 7) {
                //!This is a hidden input
                const totalTime = getTime([getInputByIndex(inputs, 5), getInputByIndex(inputs, 6)]); //Total time column
                if (totalTime > 0 && rate && amount && !amount.valueAsNumber)
                    amount.valueAsNumber = totalTime * 24 * rate; // making the amount equal the rate * totalTime
                return totalTime;
            }
            else if (debit && index === 9)
                return input.valueAsNumber * -1 || 0; //This is the amount if negative
            else if ([8, 9, 10].includes(index))
                return input.valueAsNumber || 0; //Hourly Rate, Amount, VAT
            else
                return input.value;
        });
        const stop = 'You must at least provide the client, matter, nature, date and the amount. If you provided a time start, you must provide a time end, and an hourly rate. Please review your fields';
        if (missing())
            return alert(stop);
        function missing() {
            if (row[5] === row[6])
                return false; //If the total time = 0 we do not need to alert if the hourly rate is missing
            else if (row.filter((el, i) => (i < 4 || i === 9) && !el).length > 0)
                return true; //if client name, matter, nature, date or amount are missing
            //else if (row[9]) return [5, 6,7,8].map(index => row[index] = 0).length < 1;//This means the amount has been provided and does not  depend on the time spent or the hourly rate. We set the values of the startTime and endTime to 0, and return false (length<1 must return false)
            else if (row[5] && (!row[6] || !row[8]))
                return true; //if startTime is provided but without endTime or without hourly rate
            else if (row[6] && (!row[5] || !row[8]))
                return true; //if endTime is provided but without startTime or without hourly rate
        }
        ;
        await addRowToExcelTable([row], TableRows.length - 2, excelFilePath, tableName, accessToken);
        [0, 1].map(async (index) => {
            //!We use map because forEach doesn't await
            //@ts-ignore
            await filterExcelTable(excelFilePath, tableName, TableRows[0][index], row[index].toString(), accessToken);
        });
        alert('Row aded and table was filtered');
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
            const quarter = 15 * 60 * 1000; //quarter of an hour
            let time = to - from;
            time = Math.round(time / quarter) * quarter; //We are rounding the time by 1/4 hours
            time = time / day;
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
            (function append() {
                input.classList.add(css);
                input.id = id;
                input.name = id;
                input.autocomplete = "on";
                input.dataset.index = index.toString();
                input.type = 'text';
            })();
            (function customize() {
                if ([8, 9, 10].includes(index))
                    input.type = 'number';
                else if (index === 3)
                    input.type = 'date';
                else if ([5, 6].includes(index))
                    input.type = 'time';
                else if ([4, 7, 15].includes(index))
                    input.style.display = 'none'; //We hide those 3 columns: 'Total Time' and the 'Year' and 'Link to a File'
                else if (index < 3 || index > 10) {
                    //We add a dataList for those fields
                    input.setAttribute('list', input.id + 's');
                    input.onchange = () => inputOnChange(index, TableRows.slice(1, -1), false);
                    if (![1, 16].includes(index))
                        createDataList(input.id, getUniqueValues(index, TableRows.slice(1, -1), tableName)); //We don't create the data list for columns 'Matter' (1) and 'Adress' (16) because it will be created when the 'Client' field is updated
                }
                if (index > 4 && index < 11)
                    //Those are the "Start Time", "End Time", "Total Time", "Hourly Rate", "Amount", "VAT" columns . The "Hourly Rate" input is hidden, so it can't be changed by the user. We will add the onChange event to it by simplicity
                    input.onchange = () => inputOnChange(index, undefined, false); //!We are passing the table[][] argument as undefined, and the invoice argument as false 
            })();
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
    if (!table && !invoice) {
        const boundInputs = [5, 6, 7, 9, 10]; //Those are "Start Time" (5), "End Time" (6), "Total Time" (7, although it is hidden), "Amount" (9), "VAT" (10) columns. We exclude the "Hourly Rate" column (8). We let the user rest it if he wants
        boundInputs
            .forEach(i => i > index ? reset(i) : i = i);
        if (index === 9)
            boundInputs
                .forEach(i => i < index ? reset(i) : i = i);
        function reset(i) {
            const input = getInputByIndex(inputs, i);
            if (!input)
                return;
            input.valueAsNumber = 0;
            input.value = '';
        }
    }
    if (!table)
        return;
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
        console.log("Row added successfully!");
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
            criterion1: `=${filterValue}`,
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
        console.log(`Filter applied to column ${columnName} successfully!`);
    }
    else {
        alert(`Error applying filter: ${await filterResponse.text()}`);
    }
}
//# sourceMappingURL=pwaVersion.js.map