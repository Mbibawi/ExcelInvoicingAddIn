"use strict";
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Excel-specific initialization code goes here
        console.log("Excel is ready!");
        loadMsalScript();
    }
});
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
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const table = sheet.tables.getItem('LivreJournal');
        const header = table.getHeaderRowRange();
        header.load('values');
        table.columns.load('items');
        const dataBodyRange = table.getDataBodyRange();
        dataBodyRange.load('values');
        await context.sync();
        const title = header.values[0];
        if (id === 'entry')
            await addingEntry(table.columns.items, title, dataBodyRange.values);
        else if (id === 'invoice')
            await invoice(title);
    });
    async function invoice(title) {
        const inputs = await Promise.all(insertInputsAndLables([0, 1, 2, 3]));
        form.innerHTML += `<button onclick="generateInvoice()"> Filter Table</button>`;
        inputs.forEach(input => input?.addEventListener('focusout', async () => await inputOnChange(input), { passive: true }));
        function insertInputsAndLables(indexes) {
            const id = 'input';
            return indexes.map(async (index) => {
                if (!index)
                    return;
                const input = document.createElement('input');
                input.type = "text";
                if (index === 3)
                    input.type = 'date';
                input.id = id + index.toString();
                input.name = input.id;
                input.dataset.index = index.toString();
                input.setAttribute('list', input.id + 's'),
                    input.autocomplete = "on";
                const label = document.createElement('label');
                label.htmlFor = input.id;
                label.innerText = title[index];
                form.appendChild(label);
                form.appendChild(input);
                if (index === 0)
                    createDataList(input?.id, await getUniqueValues(index)); //We create a unique values dataList for the 'Client' input
                return input;
            });
        }
        ;
        async function inputOnChange(input, unfilter = false) {
            const index = Number(input.dataset.index);
            if (index === 0)
                unfilter = true; //If this is the 'Client' column, we remove any filter from the table;
            //We filter the table accordin to the input's value and return the visible cells
            const visibleCells = await filterTable(undefined, [{ column: index, value: getArray(input.value) }], unfilter);
            if (visibleCells.length < 1)
                return alert('There are no visible cells in the filtered table');
            //We create (or update) the unique values dataList for the next input 
            const nextInput = getNextInput(input);
            if (!nextInput)
                return;
            createDataList(nextInput?.id || '', await getUniqueValues(Number(nextInput.dataset.index), visibleCells));
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
                const nature = new Set((await filterTable(undefined, undefined, false)).map(row => row[index]));
                nature.forEach(el => form.appendChild(createCheckBox(undefined, el)));
            }
        }
        ;
    }
    async function addingEntry(columns, title, dataBodyRange) {
        await filterTable(undefined, undefined, true);
        for (const col of columns) {
            const i = columns.indexOf(col);
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
            .forEach(input => input.style.display = 'none');
        async function onFoucusOut(input) {
            debugger;
            const i = Number(input.dataset.index);
            const criteria = [{ column: i, value: getArray(input.value) }];
            let unfilter = false;
            if (i === 0)
                unfilter = true;
            await filterTable(undefined, criteria, unfilter);
            if (i < 1)
                createDataList('input' + String(i + 1), await getUniqueValues(i + 1));
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
                createDataList(input.id, await getUniqueValues(i, dataBodyRange));
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
    /**
     * Creates a dataList with the provided id from the unique values of the column which index is passed as parameter
     * @param {string} id - the id of the dataList that will be created
     * @param {number} index - the index of the column from which the unique values of the datalist will be retrieved
     *
    */
    function createDataList(id, uniqueValues) {
        //const uniqueValues = Array.from(new Set(visible.map(row => row[i])));
        if (!id || !uniqueValues || uniqueValues.length < 1)
            return;
        id += 's';
        // Create a new datalist element
        let dataList = Array.from(document.getElementsByTagName('datalist')).find(list => list.id === id);
        if (dataList)
            dataList.remove();
        dataList = document.createElement('datalist');
        //@ts-ignore
        dataList.id = id;
        // Append options to the datalist
        uniqueValues.forEach(option => {
            const optionElement = document.createElement('option');
            optionElement.value = option;
            dataList?.appendChild(optionElement);
        });
        // Attach the datalist to the body or a specific element
        document.body.appendChild(dataList);
    }
    ;
}
/**
 * Filters the Excel table based on a criteria
 * @param {[[number, string[]]]} criteria - the first element is the column index, the second element is the values[] based on which the column will be filtered
 */
async function filterTable(tableName = 'LivreJournal', criteria, clearFilter = false) {
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
async function generateInvoice(lang = 'FR') {
    const inputs = Array.from(document.getElementsByName('input'));
    if (!inputs)
        return;
    const visible = await filterTable(undefined, undefined, false);
    const clientId = "157dd297-447d-4592-b2d3-76b643b97132"; //the new one
    const tenantId = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
    const redirectUri = "https://mbibawi.github.io/ExcelInvoicingAddIn"; //!must be the same domain as the app
    // MSAL configuration
    const msalConfig = {
        auth: {
            clientId: clientId,
            authority: "https://login.microsoftonline.com/common",
            redirectUri: redirectUri,
        },
        cache: {
            cacheLocation: "ExcelAddIn",
            storeAuthStateInCookie: true
        }
    };
    const invoice = {
        clientName: visible.map(row => String(row[0]))[0] || 'CLIENT',
        matters: (await getUniqueValues(1, visible)).map(el => String(el)),
        adress: (await getUniqueValues(15, visible)).map(el => String(el)),
        lang: lang
    };
    const drive = {
        clientId: clientId,
        redirectUri: redirectUri,
        msalConfig: msalConfig,
        tenantId: tenantId
    };
    await uploadWordDocument(getData(), invoice, drive);
    function getData() {
        const lables = {
            totalFees: {
                nature: 'Honoraire',
                FR: 'Total honoraires',
                EN: 'Total Fees'
            },
            totalExpenses: {
                nature: 'Débours/Dépens',
                FR: 'Total débours et frais',
                EN: 'Total Expenses'
            },
            totalPayments: {
                nature: 'Provision/Règlement',
                FR: 'Total provisions reçues',
                EN: 'Total Payments'
            },
            totalDue: {
                FR: 'Montant dû',
                EN: 'Total Due'
            },
            hourlyBilled: {
                nature: '',
                FR: 'facturation au temps passé : ',
                EN: 'hourly billed: ',
            },
            hourlyRate: {
                nature: '',
                FR: ' au taux horaire de : ',
                EN: ' at an hourly rate of: ',
            },
            totalTimeSpent: {
                FR: 'Total des heures facturables (hors prestations facturées au forfait) ',
                EN: 'Total billable hours (other than lump-sum billed services)'
            },
            decimal: {
                nature: '',
                FR: ',',
                EN: '.'
            },
        };
        const amount = 9, vat = 10, hours = 7, rate = 8, nature = 2, descr = 14;
        const data = visible.map(row => {
            const date = dateFromExcel(row[3]);
            const time = getTimeSpent(row[hours]);
            let description = `${String(row[nature])} : ${String(row[descr])}`; //Column Nature + Column Description;
            //If the billable hours are > 0
            if (time)
                //@ts-ignore
                description += `(${lables.hourlyBilled[lang]} ${time} ${lables.hourlyRate[lang]} ${Math.abs(row[rate]).toString()} €)`;
            const rowValues = [
                [date.getDate(), date.getMonth() + 1, date.getFullYear()].join('/'), //Column Date
                description,
                (row[amount] * -1).toFixed(2) + ' €', //Column "Amount": we inverse the +/- sign for all the values 
                Math.abs(row[vat]).toFixed(2) + ' €', //Column VAT: always a positive value
            ];
            return rowValues;
        });
        pushTotalsRows();
        return data;
        function pushTotalsRows() {
            //Adding rows for the totals of the different categories and amounts
            const totalFee = getTotals(amount, lables.totalFees.nature);
            const totalFeeVAT = getTotals(vat, lables.totalFees.nature);
            const totalPayments = getTotals(amount, lables.totalPayments.nature);
            const totalPaymentsVAT = getTotals(vat, lables.totalPayments.nature);
            const totalExpenses = getTotals(amount, lables.totalExpenses.nature);
            const totalExpensesVAT = getTotals(vat, lables.totalExpenses.nature);
            const totalTimeSpent = getTotals(hours, null);
            const totalDueVAT = getTotals(vat, null);
            const totalDue = totalFee + totalExpenses - totalPayments;
            if (totalFee > 0)
                pushSumRow(lables.totalFees, totalFee, totalFeeVAT);
            if (totalExpenses > 0)
                pushSumRow(lables.totalExpenses, totalExpenses, totalExpensesVAT);
            if (totalPayments > 0)
                pushSumRow(lables.totalPayments, totalPayments, totalPaymentsVAT);
            if (totalTimeSpent > 0)
                pushSumRow(lables.totalTimeSpent, totalTimeSpent); //!We don't pass the vat argument in order to get the corresponding cell of the Word table empty
            pushSumRow(lables.totalDue, totalDue, totalDueVAT);
            function pushSumRow(label, amount, vat) {
                if (!amount)
                    return;
                data.push([
                    //@ts-ignore
                    label[lang],
                    '',
                    amountString(Math.abs(amount)), //The total amount can be a negative number, that's why we use Math.abs() in order to get the absolute number without the negative sign
                    //@ts-ignore
                    Number(vat) >= 0 ? Math.abs(vat).toFixed(2).replace('.', lables.decimal[lang]) + ' €' : '' //!We must check not only that vat is a number, but that it is >=0 in order to avoid getting '' each time the vat is = 0, because we need to show 0 vat values
                ]);
                function amountString(amount) {
                    if (label === lables.totalTimeSpent)
                        return getTimeSpent(amount);
                    //@ts-ignore
                    return amount.toFixed(2).replace('.', lables.decimal[lang]) + ' €';
                }
            }
            function getTotals(index, nature) {
                const total = visible.filter(row => nature ? row[2] === nature : row[2] === row[2])
                    .map(row => Number(row[index]));
                let sum = 0;
                for (let i = 0; i < total.length; i++) {
                    sum += total[i];
                }
                if (index === 7)
                    console.log('this is the hourly rate'); //!need to something to adjust the time spent format
                return sum;
            }
        }
        function getTimeSpent(time) {
            if (!time || time <= 0)
                return undefined;
            const minute = 60 * 1000;
            const minutes = Math.floor(time / minute);
            const hours = Math.floor(minutes / 60);
            return [hours, minutes % 60]
                .map(el => el.toString().padStart(2, '0'))
                .join(':');
        }
        function dateFromExcel(excelDate) {
            const date = new Date((excelDate - 25569) * (60 * 60 * 24) * 1000); //This gives the days converted from milliseconds. 
            const dateOffset = date.getTimezoneOffset() * 60 * 1000; //Getting the difference in milleseconds
            return new Date(date.getTime() + dateOffset);
        }
    }
}
async function getUniqueValues(index, array, tableName = 'LivreJournal') {
    if (!array)
        array = await filterTable(tableName, undefined, false);
    if (!array)
        array = [];
    return Array.from(new Set(array.map(row => row[index])));
}
;
async function uploadWordDocument(data, invoice, drive) {
    //const accessToken = await authenticateUser();
    const accessToken = await getTokenWithMSAL(drive.clientId, drive.redirectUri, drive.msalConfig);
    if (accessToken) {
        console.log("Successfully retrieved token:", accessToken);
        //Office.context.ui.messageParent(`Access token: ${accessToken}`);
    }
    else {
        return console.log("Failed to retrieve token.");
    }
    const date = new Date();
    const fileName = `Test_Facture_${invoice.clientName}_${Array.from(invoice.matters).join('&')}_${[date.getFullYear(), date.getMonth() + 1, date.getDate()].join('')}@${[date.getHours(), date.getMinutes()].join(':')}.docx`;
    const path = "Legal\\Mon Cabinet d'Avocat\\Comptabilité\\Factures\\";
    const templatePath = path + "FactureTEMPLATE [NE PAS MODIFIDER].dotm";
    const newPath = path + `Clients\\${fileName}`;
    await createWordDocumentFromTemplate(templatePath, newPath, accessToken, drive.tenantId);
    async function createWordDocumentFromTemplate(templatePath, newDocumentPath, accessToken, tenantId) {
        if (!accessToken)
            return;
        const headers = {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": "application/json",
        };
        // Fetch the template file from OneDrive
        const templateResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${templatePath}:/content`, {
            method: 'GET',
            headers: headers
        });
        const templateBlob = await templateResponse.blob();
        const templateArrayBuffer = await templateBlob.arrayBuffer();
        const templateBase64 = btoa(String.fromCharCode(...new Uint8Array(templateArrayBuffer)));
        // Create the new document with the template content
        const newDocumentResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${newDocumentPath}:/content`, {
            method: 'PUT',
            headers: headers,
            body: JSON.stringify({
                "@microsoft.graph.conflictBehavior": "rename",
                "file": {
                    "@odata.type": "#microsoft.graph.file"
                },
                "fileSystemInfo": {},
                "contentBytes": templateBase64
            })
        });
        const newDocument = await newDocumentResponse.json();
        console.log("Document created:", newDocument);
        editDocument(newDocument, data);
        // Function to open a Word document
    }
    async function editDocument(newDocument, data) {
        // Open the newly created Word document for editing
        await Word.run(async (context) => {
            const docUrl = newDocument["@microsoft.graph.downloadUrl"];
            const document = context.application.createDocument(docUrl);
            document.open();
            context.sync();
            console.log("Word document opened for editing:", document);
            const tables = context.document.body.tables;
            const contentControls = context.document.body.contentControls;
            context.load(tables);
            context.load(contentControls);
            await context.sync();
            const table = tables.items[0];
            if (!table)
                return;
            data.forEach(dataRow => table.addRows("End", 1, [dataRow]));
            await editRichTextContentControls();
            async function editRichTextContentControls() {
                const fields = {
                    dateLabel: {
                        title: 'LabelParisLe',
                        text: { FR: 'Paris le ', EN: 'Paris on ' },
                    },
                    date: {
                        title: 'RTInvoiceDate',
                        text: [date.getDay(), date.getMonth() + 1, date.getFullYear()].join('/'),
                    },
                    numberLabel: {
                        title: 'LabelInvoiceNumber',
                        text: { FR: 'Facturen n° : ', EN: 'Invoice No.:' },
                    },
                    number: {
                        title: 'RTInvoiceNumber',
                        text: [date.getDay(), date.getMonth() + 1, date.getFullYear() - 2000].join('') + '/' + [date.getHours(), date.getMinutes].join(''),
                    },
                    subjectLable: {
                        title: 'LabelSubject',
                        text: { FR: 'Objet : ', EN: 'Subject: ' },
                    },
                    subject: {
                        title: 'RTMatter',
                        text: invoice.matters.join(' & '),
                    },
                    amount: {
                        title: 'LabelTableHeadingMontantTTC',
                        text: { FR: 'Montant TTC', EN: 'Amount VAT Included' },
                    },
                    vat: {
                        title: 'LabelTableHeadingTVA',
                        text: { FR: 'TVA', EN: 'VAT' },
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
                //@ts-ignore
                await edit(fields.dateLabel.title, fields.dateLabel.text[lang] || fields.dateLabel.text.FR); //Date Label
                await edit(fields.date.title, fields.date.text); //Date
                //@ts-ignore
                await edit(fields.subjectLable.title, fields.subjectLable.text[lang] || fields.subjectLable.text.FR); //Subject Label
                await edit(fields.subject.title, fields.subject.text); //Subject
                //@ts-ignore
                await edit(fields.numberLabel.title, fields.numberLabel.text[lang] || fields.numberLabel.text.FR); //Invoice Number Label
                await edit(fields.number.text, fields.number.title); //Invoice Number
                await edit(fields.disclaimer.title, fields.disclaimer.text); //Disclaimer
                await edit(fields.adress.title, fields.adress.text); //Client adress
                await edit(fields.clientName.title, fields.clientName.text); //Client name
                //@ts-ignore
                await edit(fields.amount.title, fields.amount.text[lang] || fields.amount.text.FR); // Table Header 'Amount'
                //@ts-ignore
                await edit(fields.vat.title, fields.vat.text[lang] || fields.vat.text.FR); //Table Header 'VAT'
                async function edit(title, text) {
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
}
;
async function addEntry(tableName = 'LivreJournal') {
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
    //@ts-ignore
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
//# sourceMappingURL=main.js.map