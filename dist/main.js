"use strict";
const excelPath = "Legal/Mon Cabinet d'Avocat/Comptabilité/Comptabilité de Mon Cabinet_15 10 2023.xlsm";
const path = "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/";
const templatePath = path + 'FactureTEMPLATE [NE PAS MODIFIDER].dotm';
const destinationFolder = path + 'Clients';
const excelFilePath = "Legal/Mon Cabinet d'Avocat/Comptabilité/Comptabilité de Mon Cabinet_15 10 2023.xlsm";
const tenantId = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
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
    let table;
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        table = sheet.tables.getItem('LivreJournal');
        const header = table.getHeaderRowRange();
        header.load('text');
        await context.sync();
        const body = table.getDataBodyRange();
        body.load('text');
        await context.sync();
        const headers = header.text[0];
        const clientUniqueValues = await getUniqueValues(0, body.text);
        if (id === 'entry')
            await addingEntry(headers, clientUniqueValues);
        else if (id === 'invoice')
            await invoice(headers, clientUniqueValues);
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
    async function addingEntry(title, uniqueValues) {
        await filterTable(undefined, undefined, true);
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
                createDataList(input.id, uniqueValues);
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
async function generateInvoice() {
    const inputs = Array.from(document.getElementsByName('input'));
    if (!inputs)
        return;
    const lang = inputs.find(input => input.type === 'checkbox' && input.checked === true)?.id.slice(0, 3).toUpperCase() || 'FR';
    const visible = await filterTable(undefined, undefined, false);
    const invoiceDetails = {
        clientName: visible.map(row => String(row[0]))[0] || 'CLIENT',
        matters: (await getUniqueValues(1, visible)).map(el => String(el)),
        adress: (await getUniqueValues(15, visible)).map(el => String(el)),
        lang: lang
    };
    await uploadWordDocument(getData(), getContentControlsValues(invoiceDetails), await getAccessToken() || '', destinationFolder, newWordFileName(new Date(), invoiceDetails.clientName, invoiceDetails.matters));
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
                getAmountString(row[amount] * -1), //Column "Amount": we inverse the +/- sign for all the values 
                getAmountString(Math.abs(row[vat])), //Column VAT: always a positive value
            ];
            return rowValues;
        });
        pushTotalsRows();
        return data;
        function getAmountString(value) {
            //@ts-ignore
            return value.toFixed(2).replace('.', lables.decimal[lang] || '.') + ' €' || '';
        }
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
                amount = Math.abs(amount);
                data.push([
                    //@ts-ignore
                    label[lang],
                    '',
                    label === lables.totalTimeSpent ? getTimeSpent(amount) || '' : getAmountString(amount) || '', //The total amount can be a negative number, that's why we use Math.abs() in order to get the absolute number without the negative sign
                    //@ts-ignore
                    Number(vat) >= 0 ? getAmountString(Math.abs(vat)) : '' //!We must check not only that vat is a number, but that it is >=0 in order to avoid getting '' each time the vat is = 0, because we need to show 0 vat values
                ]);
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
            time = time * (60 * 60 * 24); //84600 is the number in seconds per day. Excel stores the time as fraction number of days like "1.5" which is = 36 hours 0 minutes 0 seconds;
            const minutes = Math.floor(time / 60);
            const hours = Math.floor(minutes / 60);
            return [hours, minutes % 60, 0]
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
async function uploadWordDocument(data, contentControls, accessToken, destinationFolder, fileName) {
    if (!accessToken)
        return;
    return await createAndEditNewXmlDoc();
    async function createAndEditNewXmlDoc() {
        const blob = await fetchBlobFromFile(templatePath, accessToken);
        if (!blob)
            return;
        const zip = await convertBlobIntoXML(blob);
        const doc = zip.xmlDoc;
        if (!doc)
            return;
        const table = getXMLTable(doc, 0);
        data.forEach(row => {
            const newXmlRow = addRowToXMLTable(doc, table);
            if (!newXmlRow)
                return;
            row.forEach(el => addCellToXMLTableRow(doc, newXmlRow, el));
        });
        contentControls
            .forEach(([title, text]) => {
            const control = findXMLContentControlByTitle(doc, title);
            if (!control)
                return;
            editXMLContentControl(control, text);
        });
        const newBlob = await convertXMLIntoBlob(doc, zip.zip);
        await uploadToOneDrive(newBlob, destinationFolder, fileName, accessToken);
    }
    //await editDocumentWordJSAPI(await copyTemplate()?.id, accessToken, data, getContentControlsValues(invoice.lang))
    async function copyTemplate() {
        // 1. Create a new Word document from the template
        const createResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${templatePath}:/copy`, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${accessToken}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                parentReference: { path: `/drive/root:/${destinationFolder}` },
                name: fileName
            })
        });
        if (!createResponse.ok)
            throw new Error("Failed to create document");
        // 2. Poll for the File ID (since copy is async)
        const file = await getFile(0);
        if (!file || !file.id)
            throw new Error("Failed to retrieve new document ID. File may not be ready.");
        return file;
        async function getFile(i) {
            if (i > 7)
                return;
            const checkFilesResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${destinationFolder}:/children`, {
                method: "GET",
                headers: {
                    "Authorization": `Bearer ${accessToken}`
                }
            });
            if (!checkFilesResponse.ok) {
                await new Promise(res => setTimeout(res, 3000));
                getFile(i + 1);
            }
            ;
            const files = await checkFilesResponse.json();
            const file = files.value.find((f) => f.name === fileName);
            if (file) {
                console.log('file =', file);
                return file;
            }
            else {
                await new Promise(res => setTimeout(res, 3000));
                await getFile(i + 1);
            }
        }
    }
    async function fetchBlobFromFile(templatePath, accessToken) {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${templatePath}:/content`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
            },
        });
        if (!response.ok)
            throw new Error('Failed to fetch the Word file from OneDrive');
        return await response.blob();
    }
    async function convertBlobIntoXML(blob) {
        //@ts-ignore
        const zip = new JSZip();
        const arrayBuffer = await blob.arrayBuffer();
        await zip.loadAsync(arrayBuffer);
        const documentXml = await zip.file("word/document.xml").async("string");
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(documentXml, "application/xml");
        return { xmlDoc, zip };
    }
    //@ts-expect-error
    async function convertXMLIntoBlob(editedXml, zip) {
        debugger;
        const serializer = new XMLSerializer();
        let modifiedDocumentXml = serializer.serializeToString(editedXml);
        //modifiedDocumentXml = `<?xml version="1.0" encoding="UTF-8"?>\n` + modifiedDocumentXml;
        zip.file("word/document.xml", modifiedDocumentXml);
        return await zip.generateAsync({ type: "blob" });
    }
    function getXMLTable(xmlDoc, index) {
        const tables = xmlDoc.getElementsByTagName("w:tbl");
        return tables[index];
    }
    function addRowToXMLTable(xmlDoc, table) {
        if (!table)
            return;
        const row = createTableElement(xmlDoc, "w:tr");
        table.appendChild(row);
        return row;
    }
    function addCellToXMLTableRow(xmlDoc, row, text) {
        if (!xmlDoc || !row)
            return;
        const cell = createTableElement(xmlDoc, "w:tc");
        row.appendChild(cell);
        const parag = createTableElement(xmlDoc, "w:p");
        cell.appendChild(parag);
        const newRun = createTableElement(xmlDoc, "w:r");
        parag.appendChild(newRun);
        if (!text)
            return;
        const newText = createTableElement(xmlDoc, "w:t");
        newText.textContent = text;
        newRun.appendChild(newText);
    }
    function createTableElement(xmlDoc, tag) {
        return xmlDoc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', tag);
    }
    function findXMLContentControlByTitle(xmlDoc, title) {
        const contentControls = Array.from(xmlDoc.getElementsByTagName("w:sdt"));
        return contentControls.find(control => control.getElementsByTagName("w:alias")[0]?.getAttribute("w:val") === title);
    }
    function editXMLContentControl(control, text) {
        if (!control)
            return;
        const textElement = control.getElementsByTagName("w:t")[0];
        if (!textElement)
            return;
        textElement.textContent = text;
    }
}
;
async function uploadToOneDrive(blob, folderPath, fileName, accessToken) {
    const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${folderPath}/${fileName}:/content`;
    const response = await fetch(endpoint, {
        method: 'PUT',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', // Correct MIME type for Word docs
        },
        body: blob, // Use the template's content as the new document's content
    });
    response.ok ? console.log('succefully uploaded the new file') : console.log('failed to upload the file to onedrive error = ', await response.json());
}
;
function newWordFileName(date, clientName, matters) {
    // return 'test file name for now.docx'
    return `_Test_Facture_${clientName}_${Array.from(matters).join('&')}_${[date.getFullYear(), date.getMonth() + 1, date.getDate()].join('')}@${[date.getHours(), date.getMinutes()].join(':')}.docx`;
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
//# sourceMappingURL=main.js.map