"use strict";
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
/**
 *
 * @param {boolean} add - If false, the function will only show a form containing input fields for the user to provide the data for the new row to be added to the Excel Table. If true, the function will parse the values from the input fields in the form, and will add them as a new row to the Excel Table. Its default value is false.
 * @param {any[]} row - If provided, the function will add the row directly to the Excel Table without needing to retrieve the data from the inputs.
 */
async function addNewEntry(add = false, row) {
    accessToken = await getAccessToken() || '';
    (async function show() {
        if (add)
            return;
        if (!workbookPath || !tableName)
            return alert('The Excel Workbook Path or the name of the Excel table are not valid');
        TableRows = await fetchExcelTableWithGraphAPI(accessToken, workbookPath, tableName);
        if (!TableRows)
            return;
        insertAddForm(TableRows[0]);
    })();
    (async function addEntry() {
        if (!add)
            return;
        if (row)
            return await addRow(row); //If a row is already passed, we will add them directly
        await addRow(parseInputs() || undefined, true);
        function parseInputs() {
            const stop = (missing) => alert(`${missing} missing. You must at least provide the client, matter, nature, date and the amount. If you provided a time start, you must provide the end time and the hourly rate. Please review your iputs`);
            const inputs = Array.from(document.getElementsByTagName('input')); //all inputs
            const nature = getInputByIndex(inputs, 2)?.value;
            if (!nature)
                return stop('The matter is');
            const date = getInputByIndex(inputs, 3)?.valueAsDate;
            if (!date)
                return stop('The invoice date is');
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
            if (missing())
                return stop('Some of the required fields are');
            return row;
            function missing() {
                if (row.filter((value, i) => (i < 4 || i === 9) && !value).length > 0)
                    return true; //if client name, matter, nature, date or amount are missing
                //else if (row[9]) return [5, 6,7,8].map(index => row[index] = 0).length < 1;//This means the amount has been provided and does not  depend on the time spent or the hourly rate. We set the values of the startTime and endTime to 0, and return false (length<1 must return false)
                if (row[5] === row[6])
                    return false; //If the total time = 0 we do not need to alert if the hourly rate is missing
                else if (row[5] && (!row[6] || !row[8]))
                    return true; //if startTime is provided but without endTime or without hourly rate
                else if (row[6] && (!row[5] || !row[8]))
                    return true; //if endTime is provided but without startTime or without hourly rate
            }
            ;
        }
        async function addRow(row, filter = false) {
            if (!row)
                return;
            await addRowToExcelTableWithGraphAPI([row], TableRows.length - 2, workbookPath, tableName, accessToken);
            if (!filter)
                return;
            [0, 1].map(async (index) => {
                //!We use map because forEach doesn't await
                await filterExcelTable(workbookPath, tableName, TableRows[0]?.[index], [row[index]?.toString()] || [], accessToken);
            });
            alert('Row aded and table was filtered');
        }
    })();
    function insertAddForm(title) {
        const form = document.getElementById('form');
        if (!form)
            return;
        form.innerHTML = '';
        title.forEach((title, index) => {
            if (![4, 7].includes(index))
                form.appendChild(createLable(title, index)); //We exclued the labels for "Total Time" and for "Year"
            form.appendChild(createInput(index));
        });
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
                else if ([4, 7].includes(index))
                    input.style.display = 'none'; //We hide those 3 columns: 'Total Time' and the 'Year' and 'Link to a File'
                else if (index < 3 || index > 10) {
                    //We add a dataList for those fields
                    input.setAttribute('list', input.id + 's');
                    input.onchange = () => inputOnChange(index, TableRows.slice(1, -1), false);
                    if (![1, 15].includes(index))
                        createDataList(input.id, getUniqueValues(index, TableRows.slice(1, -1))); //We don't create the data list for columns 'Matter' (1) and 'Adress' (16) because it will be created when the 'Client' field is updated
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
        if (!workbookPath || !tableName)
            return alert('The Excel Workbook path and/or the name of the Excel Table are missing or invalid');
        TableRows = await fetchExcelTableWithGraphAPI(accessToken, workbookPath, tableName);
        if (!TableRows)
            return;
        insertInvoiceForm(TableRows);
    })();
    (async function issueInvoice() {
        if (!issue)
            return;
        if (!templatePath || !destinationFolder)
            return alert('The full path of the Word Invoice Template and/or the destination folder where the new invoice will be saved, are either missing or not valid');
        const inputs = Array.from(document.getElementsByTagName('input'));
        const criteria = inputs.filter(input => Number(input.dataset.index) >= 0);
        const discount = parseInt(inputs.find(input => input.id === 'discount')?.value || '0%');
        const lang = inputs.find(input => input.dataset.language && input.checked === true)?.dataset.language || 'FR';
        TableRows = await fetchExcelTableWithGraphAPI(accessToken, workbookPath, tableName); //We fetch the table again in case there where changes made since it was fetched the first time when the userform was inserted
        const [wordRows, totalsRows, filtered] = filterExcelData(TableRows, criteria, discount, lang);
        const date = new Date();
        const invoice = {
            number: getInvoiceNumber(date),
            clientName: getInputValue(0, criteria),
            matters: getArray(getInputValue(1, criteria)),
            adress: getUniqueValues(15, filtered),
            lang: lang
        };
        const contentControls = getContentControlsValues(invoice, date);
        const fileName = getInvoiceFileName(invoice.clientName, invoice.matters, invoice.number);
        let filePath = `${destinationFolder}/${fileName}`;
        filePath = prompt(`The file will be saved in ${destinationFolder}, and will be named : ${fileName}./nIf you want to change the path or the name, provide the full file path and name of your choice without any sepcial characters`, filePath) || filePath;
        await createAndUploadXmlDocument(wordRows, contentControls, accessToken, templatePath, filePath, totalsRows);
        (async function filterTable() {
            await clearFilterExcelTableGraphAPI(workbookPath, tableName, accessToken); //We start by clearing the filter of the table, otherwise the insertion will fail
            [0, 1].map(async (index) => {
                await filterExcelTable(workbookPath, tableName, TableRows[0][index], getUniqueValues(index, filtered), accessToken);
            });
        })();
        /**
         * Filters the Excel table according to the values of each inputs, then returns the values of the Word table rows that will be added to the Word table in the invoice template document
         * @param {any[][]} data - The Excel table rows that will be filtered
         * @param {HTMLInputElement[]} criteria - the html inputs containing the values based on which the table will be filtered
         * @param {string} lang - The language in which the invoice will be issued
         * @returns {string[][]} - The values of the rows that will be added to the Word table in the invoice template
         */
        function filterExcelData(data, criteria, discount, lang) {
            //Filtering by Client (criteria[0])
            data = data.filter(row => row[getIndex(criteria[0])] === criteria[0].value);
            const adress = getUniqueValues(15, data); //!We must retrieve the adresses at this stage before filtering by "Matter" or any other column
            [1, 2].forEach(index => {
                //!Matter and Nature inputs (from columns 2 & 3 of the Excel table) may include multiple entries separated by ', ' not only one entry.
                const list = criteria[index].value.split(',').map(el => el.trimStart().trimEnd()); //We generate a string[] from the input.value
                data = data.filter(row => list.includes(row[index])); //We filter the data
            });
            //We finaly filter by date
            data = filterByDate(data);
            return [...getRowsData(data, discount, lang), data];
            function filterByDate(data) {
                const convertDate = (date) => dateFromExcel(Number(date)).getTime();
                const [from, to] = criteria
                    .filter(input => getIndex(input) === 3)
                    .map(input => input.valueAsDate?.getTime());
                if (from && to)
                    return data.filter(row => convertDate(row[3]) >= from && convertDate(row[3]) <= to); //we filter by the date
                else if (from)
                    return data.filter(row => convertDate(row[3]) >= from); //we filter by the date
                else if (to)
                    return data.filter(row => convertDate(row[3]) <= to); //we filter by the date
                else
                    return data.filter(row => convertDate(row[3]) <= new Date().getTime()); //we filter by the date
            }
        }
    })();
    function insertInvoiceForm(excelTable) {
        const form = document.getElementById('form');
        if (!form)
            return;
        form.innerHTML = '';
        const title = excelTable[0];
        insertInputsAndLables([0, 1, 2, 3, 3], 'input'); //Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice
        insertInputsAndLables(['Discount'], 'discount', false)[0].value = '0%'; //Inserting a discount percentage input and setting its default value to 0%
        insertInputsAndLables(['Français', 'English'], 'lang', true); //Inserting languages checkboxes
        (function customizeDateLabels() {
            const [from, to] = Array.from(document.getElementsByTagName('label'))
                ?.filter(label => label.htmlFor.endsWith('3'));
            if (from)
                from.innerText += ' From (included)';
            if (to)
                to.innerText += ' To/Before (included)';
        })();
        (function addIssueInvoiceBtn() {
            const btnIssue = document.createElement('button');
            btnIssue.innerText = 'Generate Invoice';
            btnIssue.classList.add('button');
            btnIssue.onclick = () => invoice(true);
            form.appendChild(btnIssue);
        })();
        function insertInputsAndLables(indexes, id, checkBox = false) {
            let css = 'field';
            if (checkBox)
                css = 'checkBox';
            return indexes.map((index) => {
                appendLable(index);
                return appendInput(index);
            });
            function appendInput(index) {
                const input = document.createElement('input');
                input.classList.add(css);
                !isNaN(Number(index)) ? input.id = id + index.toString() : input.id = id;
                (function setType() {
                    if (checkBox)
                        input.type = 'checkbox';
                    else if (isNaN(Number(index)) || Number(index) < 3)
                        input.type = 'text';
                    else
                        input.type = 'date';
                })();
                (function notCheckBox() {
                    if (isNaN(Number(index)) || checkBox)
                        return; //If the index is not a number or the input is a checkBox, we return;
                    index = Number(index);
                    input.name = input.id;
                    input.dataset.index = index.toString();
                    input.setAttribute('list', input.id + 's');
                    input.autocomplete = "on";
                    if (index < 2)
                        input.onchange = () => inputOnChange(Number(input.dataset.index), excelTable.slice(1, -1), true);
                    if (index < 1)
                        createDataList(input.id, getUniqueValues(0, TableRows.slice(1, -1))); //We create a unique values dataList for the 'Client' input
                })();
                (function isCheckBox() {
                    if (!checkBox)
                        return;
                    input.dataset.language = index.toString().slice(0, 2).toUpperCase();
                    input.onchange = () => Array.from(document.getElementsByTagName('input'))
                        .filter((checkBox) => checkBox.dataset.language && checkBox !== input)
                        .forEach(checkBox => checkBox.checked = false);
                })();
                form?.appendChild(input);
                return input;
            }
            function appendLable(index) {
                const label = document.createElement('label');
                isNaN(Number(index)) || checkBox ? label.innerText = index.toString() : label.innerText = title[Number(index)];
                !isNaN(Number(index)) ? label.htmlFor = id + index.toString() : label.htmlFor = id;
                form?.appendChild(label);
            }
        }
        ;
    }
}
async function issueLetter(create = false) {
    accessToken = await getAccessToken() || '';
    const templatePath = '';
    (function showForm() {
        if (create)
            return;
        const form = document.getElementById('form');
        if (!form)
            return;
        form.innerHTML = '';
        const input = document.createElement('textarea');
        (function inputAttributes() {
            input.id = 'textInput';
            input.classList.add('field');
            form.appendChild(input);
        })();
        (function generateBtn() {
            const btn = document.createElement('button');
            form?.appendChild(btn);
            btn.classList.add('button');
            btn.innerText = 'Créer lettre';
            btn.onclick = () => issueLetter(true);
        })();
    })();
    (async function generate() {
        if (!create)
            return;
        const input = document.getElementById('textInput');
        if (!input)
            return;
        const templatePath = "Legal/Mon Cabinet d'Avocat/Administratif/Modèles Actes/Template_Lettre With Letter Head [DO NOT MODIFY].docx";
        const fileName = prompt('Provide the file name without special characthers');
        if (!fileName)
            return;
        const filePath = `${prompt('Provide the destination folder', "Legal/Mon Cabinet d'Avocat/Clients")}/${fileName}.docx`;
        if (!filePath)
            return;
        const contentControls = [['RTCoreText', input.value], ['RTReference', 'Référence'], ['RTClientName', 'Nom du Client'], ['RTEmail', 'Email du client']];
        createAndUploadXmlDocument(undefined, contentControls, accessToken, templatePath, filePath);
    })();
}
/**
 * Updates the data list or the value of bound inputs according to the value of the input that has been changed
 * @param {number} index - the dataset.index of the input that has been changed
 * @param {any[][]} table - the table that will be filtered. If undefined, it means that no data list will be updated.
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
            input.value = '';
            if (input.valueAsNumber)
                input.valueAsNumber = 0;
        }
    }
    if (!table)
        return;
    if (invoice)
        inputs = inputs.filter(input => input.dataset.index && Number(input.dataset.index) < 3); //Those are all the inputs that serve to filter the table (first 3 columns only)
    else
        inputs = inputs.filter(input => [0, 1, 15].includes(getIndex(input))); //Those are all the inputs that have data lists associated with them
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
    boundInputs.map(input => createDataList(input?.id, getUniqueValues(getIndex(input), filtered), invoice));
    function filterOnInput(inputs, filled, table) {
        let filtered = table;
        for (let i = 0; i < filled.length; i++) {
            filtered = filtered.filter(row => row[filled[i]].toString() === getInputByIndex(inputs, filled[i])?.value);
        }
        return filtered;
    }
}
;
async function createAndUploadXmlDocument(rows, contentControls, accessToken, templatePath, filePath, totals = []) {
    if (!accessToken || !templatePath || !filePath)
        return;
    const schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const blob = await fetchFileFromOneDriveWithGraphAPI(accessToken, templatePath);
    if (!blob)
        return;
    const [doc, zip] = await convertBlobIntoXML(blob);
    if (!doc)
        return;
    (function editTable() {
        if (!rows)
            return;
        const table = getXMLElements(doc, "w:tbl", 1);
        if (!table)
            return;
        rows.forEach((row, index) => {
            const newXmlRow = insertRowToXMLTable();
            if (!newXmlRow)
                return;
            const isTotal = totals.includes(row[0]);
            const isLast = index === rows.length - 1;
            row.forEach((text, index) => {
                addCellToXMLTableRow(newXmlRow, getStyle(index, isTotal && !isLast), [isTotal, isLast].includes(true), text);
            });
        });
        function insertRowToXMLTable(after = -1) {
            if (!table)
                return;
            const row = createTableElement("w:tr");
            after >= 0 ? getXMLElements(table, 'w:tr', after)?.insertAdjacentElement('afterend', row) :
                table.appendChild(row);
            return row;
        }
        function getStyle(cell, isTotal = false) {
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
        function setStyle(targetElement, style, backGroundColor = '') {
            // Create or find the run properties element
            //const styleProps = createAndAppend(runElement, "w:rPr", false);
            const tag = targetElement.tagName.toLocaleLowerCase();
            (function cell() {
                if (tag !== 'w:tc')
                    return;
                const cellProp = createAndAppend(targetElement, 'w:tcPr', false);
                createAndAppend(cellProp, 'w:vAlign').setAttribute('w:val', "center");
                //createAndAppend(cellProp, 'w:tcStyle').setAttribute('w:val', 'InvoiceCellCentered');
                if (!backGroundColor)
                    return;
                const background = createAndAppend(cellProp, 'w:shd'); //Adding background color to cell
                background.setAttribute('w:val', "clear");
                background.setAttribute('w:fill', backGroundColor);
            })();
            (function parag() {
                if (tag !== 'w:p')
                    return;
                if (!style)
                    return;
                const props = createAndAppend(targetElement, "w:pPr", false);
                createAndAppend(props, "w:pStyle").setAttribute("w:val", style);
            })();
            function createAndAppend(parent, tag, append = true) {
                const newElement = createTableElement(tag);
                if (append)
                    parent.appendChild(newElement);
                else
                    parent.insertBefore(newElement, parent.firstChild);
                return newElement;
            }
        }
        function addCellToXMLTableRow(row, style, isTotal, text) {
            if (!row)
                return;
            const cell = createTableElement("w:tc"); //new table cell
            row.appendChild(cell);
            if (isTotal)
                setStyle(cell, style, 'D9D9D9'); //We set the background color of the cell
            else
                setStyle(cell, style, '');
            const parag = createTableElement("w:p"); //new table paragraph
            cell.appendChild(parag);
            setStyle(parag, style, '');
            const newRun = createTableElement("w:r"); // new run
            parag.appendChild(newRun);
            if (!text)
                return;
            const newText = createTableElement("w:t");
            newText.textContent = text;
            newRun.appendChild(newText);
        }
        function createTableElement(tag) {
            return doc.createElement(tag);
        }
    })();
    (function editContentControls() {
        if (!contentControls)
            return;
        const ctrls = getXMLElements(doc, "w:sdt");
        contentControls
            .forEach(([title, text]) => {
            const control = findXMLContentControlByTitle(ctrls, title);
            if (!control)
                return;
            editXMLContentControl(control, text);
        });
        function findXMLContentControlByTitle(ctrls, title) {
            return ctrls.find(control => control.getElementsByTagName("w:alias")[0]?.getAttribute("w:val") === title);
        }
        function editXMLContentControl(control, text) {
            if (!text)
                return control.remove();
            const textElement = control.getElementsByTagName("w:t")[0];
            if (!textElement)
                return; //!need to insert a text element instead of returning
            textElement.textContent = text;
        }
    })();
    await convertXMLToBlobAndUpload(doc, zip, filePath, accessToken);
    function getXMLElements(xmlDoc, tag, index) {
        const elements = xmlDoc.getElementsByTagName(tag);
        if (index)
            return elements[index];
        return Array.from(elements);
    }
}
;
/**
 * Converts the blob of a Word document into an XML
 * @param blob - the blob of the file to be converted
 * @returns {[XMLDocument, JSZip]} - The xml document, and the zip containing all the xml files
 */
//@ts-expect-error
async function convertBlobIntoXML(blob) {
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
 * Converts an XML Word document into a Blob, and uploads it to OneDrive using the Graph API
 * @param {XMLDocument} doc
 * @param {JSZip} zip
 * @param {string} filePath - the full OneDrive file path (including file name and extension) of the file that will be uploaded
 * @param {string} accessToken - the Graph API accessToken
 */
//@ts-expect-error
async function convertXMLToBlobAndUpload(doc, zip, filePath, accessToken) {
    const blob = await convertXMLIntoBlob();
    if (!blob)
        return;
    await uploadFileToOneDriveWithGraphAPI(blob, filePath, accessToken);
    async function convertXMLIntoBlob() {
        const serializer = new XMLSerializer();
        let modifiedDocumentXml = serializer.serializeToString(doc);
        zip.file("word/document.xml", modifiedDocumentXml);
        return await zip.generateAsync({ type: "blob" });
    }
}
;
/**
 * Convert the date in an Excel row into a javascript date (in milliseconds)
 * @param {number} excelDate - The date retrieved from an Excel cell
 * @returns {Date} - a javascript format of the date
 */
function dateFromExcel(excelDate) {
    const date = new Date((excelDate - 25569) * (60 * 60 * 24) * 1000); //This gives the days converted from milliseconds. 
    const dateOffset = date.getTimezoneOffset() * 60 * 1000; //Getting the difference in milleseconds
    return new Date(date.getTime() + dateOffset);
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
function getNewExcelRow(inputs) {
    return inputs.map(input => {
        input.value;
    });
}
async function addRowToExcelTableWithGraphAPI(row, index, filePath, tableName, accessToken) {
    await clearFilterExcelTableGraphAPI(filePath, tableName, accessToken); //We start by clearing the filter of the table, otherwise the insertion will fail
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
}
async function filterExcelTable(filePath, tableName, columnName, values, accessToken) {
    if (!accessToken)
        return;
    // Step 3: Apply filter using the column name
    const filterUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/columns/${columnName}/filter/apply`;
    const body = {
        criteria: {
            filterOn: "values",
            values: values,
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