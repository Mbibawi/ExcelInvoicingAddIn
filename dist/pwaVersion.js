"use strict";
showMainUI();
function showMainUI(homeBtn) {
    const container = byID('btns');
    if (!container)
        return;
    container.innerHTML = "";
    if (homeBtn)
        return appendBtn('home', 'Back to Main', showMainUI);
    appendBtn('entry', 'Add Entry', addNewEntry);
    appendBtn('invoice', 'Invoice', invoice);
    appendBtn('letter', 'Letter', issueLetter);
    appendBtn('lease', 'Leases', issueLeaseLetter);
    appendBtn('search', 'Search Files', searchFiles);
    appendBtn('settings', 'Settings', saveSettings);
    function appendBtn(id, text, onClick) {
        const btn = document.createElement('button');
        btn.id = id;
        btn.classList.add("ms-Button");
        btn.innerText = text;
        btn.onclick = () => onClick();
        container?.appendChild(btn);
        return btn;
    }
}
;
/**
 *
 * @param {boolean} add - If false, the function will only show a form containing input fields for the user to provide the data for the new row to be added to the Excel Table. If true, the function will parse the values from the input fields in the form, and will add them as a new row to the Excel Table. Its default value is false.
 * @param {boolean} display - If provided, the function will show the visible rows in the UI after the new row has been added.
 * @param {any[]} row - If provided, the function will add the row directly to the Excel Table without needing to retrieve the data from the inputs.
 */
async function addNewEntry(add = false, row) {
    spinner(true); //We show the spinner
    const findSetting = (name, settings) => settings?.find(setting => setting.name === name);
    const stored = getSavedSettings() || undefined;
    if (!stored)
        return;
    const workbookPath = findSetting(settingsNames.invoices.workBook, stored)?.value;
    if (!workbookPath)
        return alert('Could not get a valid workbook path from the localStorage');
    const tableName = findSetting(settingsNames.invoices.tableName, stored)?.value;
    if (!tableName)
        return alert('Could not get the name of the Excel table from the localStorage');
    if (!workbookPath || !tableName)
        throw new Error('The Excel Workbook path and/or the name of the Excel Table are missing or invalid');
    const graph = new GraphAPI('', workbookPath);
    const TableRows = await graph.fetchExcelTable(tableName, true);
    if (!TableRows?.length)
        return alert('Failed to retrieve the Excel table');
    const tableTitles = TableRows[0];
    (async function showAddNewForm() {
        if (add)
            return;
        document.querySelector('table')?.remove();
        try {
            await createForm();
            spinner(false); //We hide the sinner
        }
        catch (error) {
            spinner(false); //We hide the sinner
            alert(error);
        }
        async function createForm() {
            const sessionId = await graph.createFileSession();
            if (!sessionId)
                throw new Error('There was an issue with the creation of the file cession. Check the console.log for more details');
            const tableBody = TableRows.slice(1, -1);
            const inputs = [];
            const bound = (indexes) => inputs.filter(input => indexes.includes(getIndex(input))).map(input => [input, getIndex(input)]);
            insertAddForm(tableTitles);
            await graph.closeFileSession(sessionId);
            function insertAddForm(titles) {
                if (!titles)
                    throw new Error('The table titles are missing. Check the console.log for more details');
                const form = byID();
                if (!form)
                    throw new Error('Could not find the form element');
                form.innerHTML = '';
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
                    ]
                        .forEach(group => newDiv(NaN, divs.filter(div => group.includes(Number(div.dataset.block)))));
                })();
                (function addBtn() {
                    const btnIssue = document.createElement('button');
                    btnIssue.innerText = 'Add Entry';
                    btnIssue.classList.add('button');
                    btnIssue.onclick = () => addNewEntry(true);
                    form.appendChild(btnIssue);
                })();
                (function homeBtn() {
                    showMainUI(true);
                })();
                function newDiv(i, divs, css = "block") {
                    if (divs)
                        return groupDivs();
                    else
                        return create();
                    function create() {
                        const div = document.createElement('div');
                        div.dataset.block = i.toString();
                        form?.appendChild(div);
                        div.classList.add(css);
                        return div;
                    }
                    function groupDivs() {
                        const div = newDiv(i, undefined, "group");
                        divs?.forEach(el => div.appendChild(el));
                        form?.children[3]?.insertAdjacentElement('afterend', div);
                        return div;
                    }
                }
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
                        inputs.push(input);
                    })();
                    (function customize() {
                        if ([8, 9, 10].includes(index))
                            input.type = 'number';
                        else if (index === 3)
                            input.type = 'date';
                        else if ([5, 6].includes(index))
                            input.type = 'time';
                        else if ([4, 7].includes(index))
                            input.style.display = 'none'; //We hide those 2 columns: 'Total Time' and the 'Year'
                        (function addDataLists() {
                            const updateNext = [0, 1, 8, 15]; //Those are the indexes of the inputs (i.e; the columns numbers) that need to get an onChange event in order to update the dataLists of the next inputs when the current input is changed: "Client"(0), "Affaire"(1), "Taux Horaire"(8), "Adresses"(15)
                            if (updateNext.includes(index))
                                input.onchange = () => inputOnChange(index, bound(updateNext), tableBody, false);
                            if (![0, 2, 11, 12, 13].includes(index))
                                return; //We will initially populate the "Client"(0), Nature(2), "Payment Method"(11), "Bank Account"(12), "Third Party"(13) lists only, the other inputs will be populate when the onChange function will be called
                            populateSelectElement(input, getUniqueValues(index, tableBody));
                        })();
                        (function addRestOnChange() {
                            if (index < 5 || index > 10)
                                return;
                            //Only for the  "Start Time", "End Time", "Total Time", "Hourly Rate", "Amount", "VAT" columns . The "Total Time" input (7) is hidden, so it can't be changed by the user. We will add the onChange event to it by simplicity
                            const reset = [5, 6, 7, 9, 10]; //Those are "Start Time" (5), "End Time" (6), "Total Time" (7, although it is hidden), "Amount" (9), "VAT" (10) columns. We exclude the "Hourly Rate" column (8). We let the user rest it if he wants
                            input.onchange = () => resetInputs(bound(reset), index); //!We are passing the table[][] argument as undefined, and the invoice argument as false which means that the function will only reset the bound inputs without updating any data list
                        })();
                    })();
                    return input;
                }
                function resetInputs(inputs, index) {
                    inputs
                        .filter(([input, index]) => index > index)
                        .forEach(([input, index]) => reset(input)); //We reset any input which dataset-index is > than the dataset-index of the input that has been changed
                    if (index === 9)
                        inputs
                            .filter(([input, index]) => index < index)
                            .forEach(([input], index) => reset(input)); //If the input is the input for the "Montant" column of the Excel table, we also reset the "Start Time" (5), "End Time" (6) and "Hourly Rate" (7) columns' inputs. We do this because we assume that if the user provided the amount, it means that either this is not a fee, or the fee is not hourly billed.
                    function reset(input) {
                        if (!input)
                            return;
                        input.value = '';
                        if (input.valueAsNumber)
                            input.valueAsNumber = 0;
                    }
                }
                ;
            }
        }
    })();
    (async function addEntry() {
        if (!add)
            return;
        const display = !row?.length;
        if (!row)
            row = parseInputs() || [];
        try {
            const visibleCells = await addRow(row);
            if (visibleCells?.length)
                displayVisibleCells(visibleCells, display);
            spinner(false); //We hide the spinner
        }
        catch (error) {
            spinner(false); //We hide the spinner
            alert(error);
        }
        function parseInputs() {
            const colNature = 2, colDate = 3, colStart = 5, colEnd = 6, colRate = 8, colAmount = 9, colVAT = 10;
            const stop = (missing) => alert(`${missing} missing. You must at least provide the client, matter, nature, date and the amount. If you provided a time start, you must provide the end time and the hourly rate. Please review your iputs`);
            const inputs = Array.from(document.getElementsByTagName('input')); //all inputs
            const nature = getInputByIndex(inputs, colNature)?.value;
            if (!nature)
                return stop('The matter is');
            const date = getInputByIndex(inputs, colDate)?.valueAsDate;
            if (!date)
                return stop('The invoice date is');
            const amount = getInputByIndex(inputs, colAmount);
            const rate = getInputByIndex(inputs, colRate)?.valueAsNumber || 0;
            const debit = ['Honoraire', 'Débours/Dépens', 'Débours/Dépens non facturables', 'Rétrocession d\'honoraires', 'Charges déductibles'].includes(nature); //We check if we need to change the value sign
            const row = inputs.map((input, index) => getInputValue(index)); //!CAUTION: The html inputs are not arranged according to their dataset.index values. If we follow their order, some values will be assigned to the wrong column of the Excel table. That's why we do not pass the input itself or the dataset.index of the input to getInputValue(), but instead we pass the index of the column for which we want to retrieve the value from the relevant input.
            if (missing())
                return stop('Some of the required fields are');
            return row;
            function getInputValue(index) {
                const input = getInputByIndex(inputs, index);
                if ([colDate, colDate + 1].includes(index))
                    return getISODate(date); //Those are the 2 date columns
                else if ([colStart, colEnd].includes(index))
                    return getTime([input]); //time start and time end columns
                else if (index === 7) {
                    //!This is a hidden input
                    const timeInputs = [colStart, colEnd].map(i => getInputByIndex(inputs, i));
                    const totalTime = getTime(timeInputs); //Total time column
                    if (totalTime && rate && !amount.valueAsNumber)
                        amount.valueAsNumber = totalTime * 24 * rate; // making the amount equal the rate * totalTime
                    return totalTime;
                }
                else if (debit && index === colAmount)
                    return -input.valueAsNumber || 0; //This is the amount if negative
                else if ([colRate, colAmount, colVAT].includes(index))
                    return input.valueAsNumber || 0; //Hourly Rate, Amount, VAT
                else
                    return input.value;
            }
            function missing() {
                if (row.filter((value, i) => (i < colDate + 1 || i === colAmount) && !value).length > 0)
                    return true; //if client name, matter, nature, date or amount are missing
                //else if (row[9]) return [5, 6,7,8].map(index => row[index] = 0).length < 1;//This means the amount has been provided and does not  depend on the time spent or the hourly rate. We set the values of the startTime and endTime to 0, and return false (length<1 must return false)
                if (row[colStart] === row[colEnd])
                    return false; //If the total time = 0 we do not need to alert if the hourly rate is missing
                else if (row[colStart] && (!row[colEnd] || !row[colRate]))
                    return true; //if startTime is provided but without endTime or without hourly rate
                else if (row[colEnd] && (!row[colStart] || !row[colRate]))
                    return true; //if endTime is provided but without startTime or without hourly rate
            }
            ;
        }
        async function addRow(row) {
            if (!row)
                throw new Error('The row is not valid');
            const visibleCells = await graph.addRowToExcelTable(row, TableRows.length - 2, tableName, tableTitles);
            if (!visibleCells?.length)
                return alert('There was an issue with the adding or the filtering, check the console.log for more details');
            alert('Row aded and the table was filtered');
            return visibleCells;
        }
        ;
        function displayVisibleCells(visibleCells, display) {
            if (!display)
                return;
            const tableDiv = createDivContainer();
            const table = document.createElement('table');
            table.classList.add('table');
            tableDiv.appendChild(table);
            const columns = [0, 1, 2, 3, 7, 8, 9, 10, 14, 15]; //The columns that will be displayed in the table;
            const rowClass = 'excelRow';
            (function insertTableHeader() {
                if (!tableTitles)
                    throw new Error('No Table Titles');
                const headerRow = document.createElement('tr');
                headerRow.classList.add(rowClass);
                const thead = document.createElement('thead');
                table.appendChild(thead);
                thead.appendChild(headerRow);
                tableTitles.forEach((cell, index) => {
                    if (!columns.includes(index))
                        return;
                    addTableCell(headerRow, cell, 'th');
                });
            })();
            (function insertTableRows() {
                const tbody = document.createElement('tbody');
                table.appendChild(tbody);
                visibleCells.forEach((row, index) => {
                    if (index < 1)
                        return; //We exclude the header row
                    if (!row)
                        return;
                    const tr = document.createElement('tr');
                    tr.classList.add(rowClass);
                    tbody.appendChild(tr);
                    row.forEach((cell, index) => {
                        if (!columns.includes(index))
                            return;
                        addTableCell(tr, cell, 'td');
                    });
                });
            })();
            const form = byID();
            if (!form)
                throw new Error('The form element was not found');
            if (form) {
                form?.insertAdjacentElement('afterend', tableDiv);
            }
            function createDivContainer() {
                const id = 'retrieved';
                let tableDiv = byID(id);
                if (tableDiv) {
                    tableDiv.innerHTML = '';
                    return tableDiv;
                }
                ;
                tableDiv = document.createElement('div');
                tableDiv.classList.add('table-div');
                tableDiv.id = id;
                return tableDiv;
            }
            function addTableCell(parent, text, tag) {
                const cell = document.createElement(tag);
                //   cell.classList.add(css);
                cell.textContent = text;
                parent.appendChild(cell);
            }
        }
        ;
    })();
}
;
// Update Word Document
async function invoice(issue = false) {
    spinner(true); //We show the spinner
    const findSetting = (name, settings) => settings?.find(setting => setting.name === name);
    const stored = getSavedSettings() || undefined;
    if (!stored)
        return;
    const workbookPath = findSetting(settingsNames.invoices.workBook, stored)?.value || prompt('Provide the Excel workbook path');
    if (!workbookPath)
        return alert('Could not retrieve the path of the Excel workbook from the localStorage');
    const tableName = findSetting(settingsNames.invoices.tableName, stored)?.value;
    if (!tableName)
        return alert('Could not get the name of the Excel table from the localStorage');
    const templatePath = findSetting(settingsNames.invoices.template, stored)?.value || prompt('Provide the path for the Word invoice template');
    if (!templatePath)
        return alert('Could not get a valid Word tempalte path  from the localStorage');
    const saveTo = findSetting(settingsNames.invoices.saveTo, stored)?.value || prompt('Provide teh path for the folder where the invoice should be saved') || 'MISSING PATH';
    const graph = new GraphAPI('', workbookPath);
    const TableRows = await graph.fetchExcelTable(tableName, true);
    if (!TableRows?.length)
        return alert('Failed to retrieve the Excel table');
    const tableTitles = TableRows[0];
    (async function showInvoiceForm() {
        if (issue)
            return;
        if (!workbookPath || !tableName)
            return alert('The Excel Workbook path and/or the name of the Excel Table are missing or invalid');
        document.querySelector('table')?.remove();
        try {
            await createForm();
            spinner(false); //We hide the spinner
        }
        catch (error) {
            alert(error);
            spinner(false); //We hide the spinner
        }
        async function createForm() {
            const sessionId = await graph.createFileSession() || '';
            if (!sessionId)
                throw new Error('There was an issue with the creation of the file cession. Check the console.log for more details');
            insertInvoiceForm(tableTitles);
            await graph.closeFileSession(sessionId);
        }
        function insertInvoiceForm(tableTitles) {
            if (!tableTitles || !TableRows)
                throw new Error('The table titles or the table rows are missing. Check the console.log for more details');
            const form = byID();
            if (!form)
                throw new Error('The form element was not found');
            form.innerHTML = '';
            const tableBody = TableRows.slice(1, -1);
            const boundInputs = [];
            insertInputsAndLables([0, 1, 2, 3, 3], 'input'); //Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice
            insertInputsAndLables(['Discount'], 'discount')[0].value = '0%'; //Inserting a discount percentage input and setting its default value to 0%
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
            (function homeBtns() {
                showMainUI(true);
            })();
            function insertInputsAndLables(indexes, id, checkBox = false) {
                let css = 'field';
                if (checkBox)
                    css = 'checkBox';
                return indexes.map((index) => {
                    const div = newDiv(String(index));
                    appendLable(index, div);
                    return appendInput(index, div);
                });
                function appendInput(index, div) {
                    const NaN = isNaN(Number(index));
                    const input = document.createElement('input');
                    input.classList.add(css);
                    !NaN ? input.id = id + index.toString() : input.id = id;
                    (function setType() {
                        if (checkBox)
                            input.type = 'checkbox';
                        else if (NaN || index < 3)
                            input.type = 'text';
                        else
                            input.type = 'date';
                    })();
                    (function notCheckBox() {
                        if (NaN || checkBox)
                            return; //If the index is not a number or the input is a checkBox, we return;
                        index = Number(index);
                        input.name = input.id;
                        input.dataset.index = index.toString();
                        if (index < 3)
                            boundInputs.push([input, index]); //Fields "Client"(0), "Affaire"(1), "Nature"(2) are the inputs that will need to get their dataList created or updated each time the previous input is changed.
                        if (index < 2)
                            input.onchange = () => inputOnChange(index, boundInputs, tableBody, true); //We add onChange on "Client" (0) and "Affaire" (1) columns
                        if (index < 1)
                            populateSelectElement(input, getUniqueValues(0, tableBody)); //We create a unique values dataList for the "Client" (0) input
                    })();
                    (function isCheckBox() {
                        if (!checkBox)
                            return;
                        input.dataset.language = index.toString().slice(0, 2).toUpperCase();
                        input.onchange = () => Array.from(document.getElementsByTagName('input'))
                            .filter((checkBox) => checkBox.dataset.language && checkBox !== input)
                            .forEach(checkBox => checkBox.checked = false);
                    })();
                    div.appendChild(input);
                    return input;
                }
                function appendLable(index, div) {
                    const label = document.createElement('label');
                    isNaN(Number(index)) || checkBox ? label.innerText = index.toString() : label.innerText = tableTitles[Number(index)];
                    !isNaN(Number(index)) ? label.htmlFor = id + index.toString() : label.htmlFor = id;
                    div?.appendChild(label);
                }
                function newDiv(i, css = "block") {
                    const div = document.createElement('div');
                    div.dataset.block = i;
                    form?.appendChild(div);
                    div.classList.add(css);
                    return div;
                }
            }
            ;
        }
    })();
    (async function issueInvoice() {
        if (!issue)
            return;
        try {
            await editInvoice();
            spinner(false); //We hide the spinner
        }
        catch (error) {
            spinner(false); //We hide the sinner
            alert(error);
        }
    })();
    async function editInvoice() {
        const client = tableTitles[0], matter = tableTitles[1]; //Those are the 'Client' and 'Matter' columns of the Excel table
        const sessionId = await graph.createFileSession(true) || ''; //!persist must be = true because we might add a new row if there is a discount. If we don't persist the session, the table will be filtered and the new row will not be added.
        if (!sessionId)
            throw new Error('There was an issue with the creation of the file cession. Check the console.log for more details');
        const inputs = Array.from(document.getElementsByTagName('input'));
        const criteria = inputs.filter(input => getIndex(input) >= 0);
        const discount = parseInt(inputs.find(input => input.id === 'discount')?.value || '0%');
        const lang = inputs.find(input => input.dataset.language && input.checked === true)?.dataset.language || 'FR';
        const date = new Date(); //We need to generate the date at this level and pass it down to all the functions that need it
        const invoiceNumber = getInvoiceNumber(date);
        const data = await filterExcelData(criteria, discount, lang, invoiceNumber);
        if (!data)
            throw new Error('Could not retrieve the filtered Excel table');
        const [wordRows, totalsLabels, clientName, matters, adress] = data;
        const invoice = {
            number: invoiceNumber,
            clientName: clientName,
            matters: matters,
            adress: adress,
            lang: lang
        };
        const contentControls = getContentControlsValues(invoice, date);
        const fileName = getInvoiceFileName(clientName, matters, invoiceNumber);
        let savePath = `${saveTo}/${fileName}`;
        savePath = prompt(`The file will be saved in ${saveTo}, and will be named : ${fileName}.\nIf you want to change the path or the name, provide the full file path and name of your choice without any sepcial characters`, savePath) || savePath;
        (async function editInvoiceFilterExcelClose() {
            await graph.createAndUploadDocumentFromTemplate(templatePath, savePath, lang, [['Invoice', wordRows, 1]], contentControls, totalsLabels);
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
        async function filterExcelData(inputs, discount, lang, invoiceNumber) {
            const matterCol = 1, dateCol = 3, addressCol = 15; //Indexes of the 'Matter' and 'Date' columns in the Excel table
            const clientName = getInputByIndex(inputs, 0)?.value || '';
            const matters = getArray(getInputByIndex(inputs, matterCol)?.value) || []; //!The Matter input may include multiple entries separated by ', ' not only one entry.
            if (!clientName || !matters?.length)
                throw new Error('could not retrieve the client name or the matter/matters list from the inputs');
            await graph.clearFilterExcelTable(tableName, sessionId); //We unfilter the table;
            //Filtering by Client (criteria[0])
            await graph.filterExcelTable(tableName, client, [clientName], sessionId);
            let visible = await graph.getVisibleCells(tableName, sessionId);
            if (!visible) {
                return alert('Could not retrieve the visible cells of the Excel table');
            }
            visible = visible.slice(1, -1); //We exclude the first and the last rows of the table. Since we are calling the "range" endpoint, we get the whole table including the headers. The first row is the header, and the last row is the total row.
            const adresses = getUniqueValues(addressCol, visible); //!We must retrieve the adresses at this stage before filtering by "Matter" or any other column
            visible = visible.filter(row => matters.includes(row[matterCol]));
            //We finaly filter by date
            visible = filterByDate(visible, dateCol);
            const [wordRows, totalLabels] = getRowsData(visible, discount, lang, invoiceNumber);
            return [wordRows, totalLabels, clientName, matters, adresses];
            function filterByDate(visible, date) {
                const convertDate = (date) => dateFromExcel(Number(date)).getTime();
                const [from, to] = inputs
                    .filter(input => getIndex(input) === date)
                    .map(input => input.valueAsDate?.getTime());
                if (from && to)
                    return visible.filter(row => convertDate(row[date]) >= from && convertDate(row[3]) <= to); //we filter by the date
                else if (from)
                    return visible.filter(row => convertDate(row[date]) >= from); //we filter by the date
                else if (to)
                    return visible.filter(row => convertDate(row[date]) <= to); //we filter by the date
                else
                    return visible.filter(row => convertDate(row[date]) <= new Date().getTime()); //we filter by the date
            }
        }
    }
}
async function issueLetter(create = false) {
    spinner(true); //We show the spinner
    (function showForm() {
        if (create)
            return;
        document.querySelector('table')?.remove();
        const form = byID();
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
        (function homeBtn() {
            showMainUI(true);
            spinner(false); //We hide the spinner
        })();
    })();
    (async function generate() {
        if (!create)
            return;
        const input = byID('textInput');
        if (!input)
            return;
        const stored = getSavedSettings();
        const templatePath = stored?.find(setting => setting.name === settingsNames.letter.template).value;
        if (!templatePath)
            return;
        const fileName = prompt('Provide the file name without special characthers');
        if (!fileName)
            return;
        const saveTo = stored?.find(setting => setting.name === settingsNames.letter.saveTo).value || 'NO STORED DEFAULT FOLDER PATH FOUND';
        const saveToPath = `${prompt('Provide the destination folder', saveTo)}/${fileName}.docx`;
        if (!saveToPath)
            return;
        const contentControls = [['RTCoreText', input.value], ['RTReference', 'Référence'], ['RTClientName', 'Nom du Client'], ['RTEmail', 'Email du client']];
        try {
            new GraphAPI('', saveToPath).createAndUploadDocumentFromTemplate(templatePath, saveToPath, 'FR', undefined, contentControls);
            spinner(false); //We hide the spinner
        }
        catch (err) {
            console.log(`There was an error: ${err}`);
            spinner(false); //We hide the spinner
        }
    })();
}
async function issueLeaseLetter(create = false) {
    spinner(true); //We show the spinner
    const findSetting = (name, settings) => settings?.find(setting => setting.name === name);
    const stored = getSavedSettings() || undefined;
    if (!stored)
        return;
    const workbookPath = findSetting(settingsNames.leases.workBook, stored)?.value || prompt('Provide the Excel workbook path');
    if (!workbookPath)
        return alert('Could not retrieve the path of the Excel workbook from the localStorage');
    const tableName = findSetting(settingsNames.leases.tableName, stored)?.value;
    if (!tableName)
        return alert('Could not get the name of the Excel table from the localStorage');
    const graph = new GraphAPI('', workbookPath);
    const tableRows = await graph.fetchExcelTable(tableName, false); //We are calling the "/rows" endPoint, so we will get the tableBody without the headers
    const Ctrls = {
        owner: { title: 'RTBailleur', col: 0, label: 'Nom du Bailleur', type: 'select', value: '' },
        adress: { title: 'RTAdresseDestinataire', label: 'Adresse du bien loué', col: 1, type: 'select', value: '' },
        tenant: { title: 'RTLocataire', label: 'Nom du Locataire', col: 2, type: 'select', value: '' },
        leaseDate: { title: 'RTDateBail', label: 'Date du Bail', col: 3, type: 'date', value: '' },
        leaseType: { title: 'RTNature', label: 'Nature du Bail', col: 4, type: 'text', value: '' },
        initialIndex: { title: 'RTIndiceInitial', label: 'Indice initial', col: 5, type: 'number', value: '' },
        indexQuarter: { title: 'RTTrimestre', label: 'Trimestre de l\'indice', col: 6, type: 'number', value: '' },
        initialIndexDate: { title: 'RTIndiceInitialDate', label: 'Date de l\'indice initial', col: 7, type: 'date', value: '' },
        baseIndex: { title: 'RTIndiceBase', label: 'Indice de référence', col: 8, type: 'number', value: '' },
        baseIndexDate: { title: 'RTDateIndiceBase', label: 'Date de l\'indice de référence', col: 9, type: 'date', value: '' },
        index: { title: 'RTIndice', label: 'Indice de révision', col: 10, type: 'number', value: '' },
        indexDate: { title: 'RTDateIndice', label: 'Date de l\'indice de révision', col: 11, type: 'date', value: '' },
        currentLease: { title: 'RTLoyerActuel', label: 'Loyer Actuel (ou révisé)', col: 12, type: 'number', value: '' },
        revisionDate: { title: 'RTDateRévision', label: 'Date de la dernière Révision', col: 13, type: 'date', value: '' },
        anniversaryDate: { title: 'RTDateAnniversaire', value: '' },
        initialYear: { title: 'RTIndiceInitialAnnée', value: '' },
        baseYear: { title: 'RTIndiceBaseAnnée', value: '' },
        revisionYear: { title: 'RTIndiceAnnée', value: '' },
        newLease: { title: 'RTLoyerNouveau', value: '' },
        nextRevision: { title: 'RTProchaineRevision', value: '' },
    };
    const ctrls = Object.values(Ctrls);
    const findRT = (id) => ctrls.find(RT => RT.title === id);
    let row, rowIndex = null;
    (async function showForm() {
        if (create)
            return;
        const inputs = [];
        const findInput = (id) => inputs.find(([input, col]) => input.id === id)?.[0];
        if (!tableRows)
            return;
        document.querySelector('table')?.remove();
        const form = byID();
        if (!form)
            return;
        form.innerHTML = '';
        const divs = [];
        (function insertInputs() {
            const unvalid = (values) => values.find(value => !value || isNaN(Number(value)));
            ctrls
                .filter(RT => !isNaN(RT.col))
                .map(RT => inputs.push([createInput(RT), RT.col]));
            const owner = findInput(Ctrls.owner.title);
            if (owner)
                populateSelectElement(owner, getUniqueValues(Ctrls.owner.col, tableRows), false);
            (function inputsOnChange() {
                const filled = inputs.filter(([input, col]) => col <= Ctrls.tenant.col);
                filled.forEach(([input, col]) => input.onchange = () => row = inputOnChange(col, inputs, tableRows, false));
                const index = findInput(Ctrls.index.title);
                if (index)
                    index.onchange = () => {
                        if (!row?.length)
                            return alert('No lease having owner name, property adress and tenant name as in the inputs was found');
                        rowIndex = tableRows.indexOf(row);
                        const initial = row[Ctrls.initialIndex.col]; //This is the value of the inital index
                        const base = row[Ctrls.baseIndex.col] || initial; //This is the value of the base index
                        const latestIndex = index.value; //this is the latest index
                        const currentLease = row[Ctrls.currentLease.col]; //This is the value of the current lease
                        if (unvalid([base, latestIndex, currentLease]))
                            return alert('Please make sure that the values of the current lease, the base indice and the new indice are all provided and valid numbers');
                        const newLease = (Number(currentLease) * (Number(latestIndex) / Number(base))).toFixed(2).toString();
                        const currentLeaseInput = findInput(Ctrls.currentLease.title);
                        currentLeaseInput.value = newLease; //This will show the value of the new lease after applying the calculation
                        Ctrls.baseIndex.value = latestIndex; //We replace the value of the base index with the latest index
                        Ctrls.newLease.value = newLease; //We update the new lease RT
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
            ]
                .forEach((group, index) => groupDivs(divs.filter(div => group.includes(getIndex(div))), index));
            function groupDivs(divs, i) {
                const div = document.createElement('div');
                div.classList.add("group");
                div.dataset.block = i.toString();
                divs?.forEach(el => div.appendChild(el));
                form?.appendChild(div);
                return div;
            }
        })();
        (function generateBtn() {
            const btn = document.createElement('button');
            form?.appendChild(btn);
            btn.classList.add('button');
            btn.innerText = 'Créer lettre';
            btn.onclick = () => generate(inputs, row);
        })();
        (function homeBtn() {
            showMainUI(true);
            spinner(false); //We hide the spinner
        })();
        function createInput(RT, className = 'field') {
            const id = RT.title;
            const div = document.createElement('div');
            form?.appendChild(div);
            const append = (el) => div.appendChild(el);
            (function appendLabel() {
                if (!RT.label)
                    return;
                const label = document.createElement('label');
                label.htmlFor = id;
                label.innerText = RT.label;
                append(label);
            })();
            return appendInput();
            function appendInput() {
                const input = document.createElement('input');
                input.type = RT.type || 'text';
                input.id = id;
                input.classList.add(className);
                const col = RT.col.toString();
                input.dataset.index = col;
                div.dataset.index = col;
                append(input);
                divs.push(div);
                return input;
            }
            ;
        }
        ;
    })();
    async function generate(inputs, row) {
        if (!inputs.length)
            return alert('Either the inputs collection or the lease or the lease index are missing');
        const findInputById = (RT) => inputs.find(([input, col]) => input.id === RT.title)?.[0];
        const templatePath = "Legal/Mon Cabinet d'Avocat/Administratif/Modèles Actes/Template_Révision de loyer [DO NOT MODIFY].docx";
        const date = new Date();
        const fileName = prompt('Provide the file name without special characthers');
        if (!fileName)
            return;
        const savePath = `${prompt('Provide the destination folder', "Legal/Mon Cabinet d'Avocat/Clients")}/${fileName}_${date.getFullYear()}${date.getMonth() + 1}${date.getDate()}@${date.getHours()}-${date.getMinutes()}.docx`;
        if (!savePath)
            return;
        inputs.map(([input, col]) => {
            const id = input.id;
            if (id === Ctrls.currentLease.title)
                return; //!We don't update the value of current lease from the input because the value of the input is the new lease not the old one
            const RT = findRT(id);
            if (!RT)
                return;
            if (RT.type === 'date')
                RT.value = getDateString(input.valueAsDate);
            else
                RT.value = input.value;
        });
        (function setMissingValues() {
            if (!row)
                return alert('The values in the input did not identifiy a unique lease in the Excel table');
            const leaseDate = dateFromExcel(row[Ctrls.leaseDate.col]);
            const getYear = (date) => dateFromExcel(date).getFullYear().toString();
            const anniversary = (year) => [leaseDate.getDate(), leaseDate.getMonth() + 1, year].join('/');
            const year = date.getFullYear();
            Ctrls.initialYear.value = getYear(row[Ctrls.initialIndexDate.col]);
            Ctrls.baseYear.value = getYear(row[Ctrls.baseIndexDate.col]);
            Ctrls.anniversaryDate.value = anniversary(year);
            Ctrls.revisionDate.value = getDateString(date);
            Ctrls.revisionYear.value = year.toString();
            Ctrls.nextRevision.value = anniversary(year + 1);
        })();
        const contentControls = ctrls.map(RT => [RT.title, RT.value]);
        try {
            await graph.createAndUploadDocumentFromTemplate(templatePath, savePath, 'FR', undefined, contentControls);
            await updateExcelTable();
            spinner(false); //We hide the spinner
        }
        catch (error) {
            console.log(error);
            alert(error);
            spinner(false); //We hide the spinner
        }
        async function updateExcelTable() {
            if (!tableName)
                return;
            const TableRows = await graph.fetchExcelTable(tableName, true);
            if (!TableRows?.length)
                return alert('Failed to retrieve the Excel table');
            const tableTitles = TableRows[0];
            (async function updateRow() {
                if (!row || !rowIndex)
                    return;
                inputs.forEach(input => update(row, input));
                await graph.updateExcelTableRow(tableName, rowIndex, row);
                return;
                const col = Ctrls.revisionDate.col;
                const revisionDate = findInputById(Ctrls.revisionDate)?.valueAsDate || undefined;
                row[col] = getISODate(revisionDate);
            })();
            (async function newRow() {
                if (row || rowIndex)
                    return; //This a scenario where no row has ever been found for the specified lease
                row = Array(inputs.length);
                inputs.forEach(input => update(row, input));
                await graph.addRowToExcelTable(row, rowIndex, tableName);
            })();
            function update(row, [input, col]) {
                if (input.type === 'date')
                    row[col] = getISODate(input.valueAsDate || undefined);
                else if (input.type === 'number')
                    row[col] = input.valueAsNumber;
                else
                    row[col] = input.value;
            }
        }
        ;
    }
}
/**
 * Updates the data list or the value of bound inputs according to the value of the input that has been changed
 * @param {number} index - the dataset.index of the input that has been changed
 * @param {any[][]} table - The table that will be filtered to update the data list of the button. If undefined, it means that the data list will not be updated.
 * @param {boolean} invoice - If true, it means that we called the function in order to generate an invoice. If false, we called it in order to add a new entry in the table
 * @returns
 */
function inputOnChange(index, inputs, table, invoice) {
    if (!table?.length)
        return;
    const filledInputs = inputs
        .filter(([input, col]) => input.value && col <= index); //Those are all the inputs that the user filled with data
    const filtered = filterTableByInputsValues(filledInputs, table); //We filter the table based on the filled inputs
    if (!filtered.length)
        return;
    const boundInputs = inputs.filter(([input, col]) => col > index); //Those are the inputs for which we want to create  or update their data lists
    for (const [input, col] of boundInputs) {
        input.value = ''; //We reset the value of all bound inputs.
        const list = getUniqueValues(col, filtered);
        const row = fillBound(list, input);
        if (row)
            return row; //!If the function returns true, it means that we filled the value of all the bound inputs, so we break the loop. If it returns false, it means that there is more than one value in the list, so we need to create or update the data list of the input.
        //if (fillBound(list, input)) break;//!If the function returns true, it means that we filled the value of all the bound inputs, so we break the loop. If it returns false, it means that there is more than one value in the list, so we need to create or update the data list of the input.
        const combine = (invoice && [1, 2].includes(col)); //For the "Matter" and "Nature" lists, we add a new element combining all the values separated by ","
        populateSelectElement(input, list, combine);
    }
    function fillBound(list, input) {
        if (list.length > 1)
            return;
        const value = list[0], found = filtered.length < 2;
        if (!found)
            return setValue(input, value); //If the filtered array contains more than one row with the same unique value in the corresponding column, we will not fill the next inputs
        const row = filtered[0]; //This is the unique row in the filtered list, we will use it to fill all the other inputs
        boundInputs.forEach(([input, col]) => setValue(input, row[col]));
        return row;
    }
    function setValue(input, value) {
        if (input.type === "date")
            input.valueAsDate = dateFromExcel(value); //!We must convert the dates from Excel
        else
            input.value = value?.toString() || '';
    }
    ;
}
;
/**
 * Filters the table according to the values of the inputs. The value of each input is compared to the value of the cell in the corresponding column in the table. If the value of the input is included in the cell value, it means that this row matches the criteria of this input. For a row to be included in the resulting filtered table, it must match the criteria of all the inputs.
 * @param {HTMLInputElement[]} inputs - the html inputs containing the values based on which the table will be filtered
 * @param {any[][]} table - The table that will be filtered
 * @returns {any[][]} - The resulting filtered table
 */
function filterTableByInputsValues(inputs, table) {
    const values = inputs.map(([input, index]) => [index, input.value.split(splitter)]); //!some inputs may contain multiple comma separated values if the user has selected more than one option in the data list. So we split the input value by ", " and we check if the cell value is included in the resulting array.
    return table.filter(row => values.every(([index, value]) => value.includes(row[index])));
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
    return inputs.map(input => input.value);
}
function searchFiles() {
    spinner(true); //We show the spinner
    const graph = new GraphAPI('');
    (function showForm() {
        const form = byID('form');
        if (!form)
            return;
        form.innerHTML = '';
        if (localStorage.folderPath)
            fetchAllDriveFiles(form, localStorage.folderPath); //We will delete the record for this folder path from the database
        (function RegExpInput() {
            const regexp = document.createElement('input');
            regexp.id = 'search';
            regexp.classList.add('field');
            regexp.placeholder = 'Enter your file name search as a regular expression';
            regexp.onkeydown = (e) => e.key === 'Enter' ? fetchAllDriveFiles(form) : e.key;
            form.appendChild(regexp);
        })();
        (function dateAfterInput() {
            const after = document.createElement('input');
            after.type = 'date';
            after.id = 'after';
            after.classList.add('field');
            after.title = 'You can proivde the date after which the file was created';
            form.appendChild(after);
        })();
        (function dateAfterInput() {
            const before = document.createElement('input');
            before.type = 'date';
            before.id = 'before';
            before.title = 'You can provide the date before which the file was created';
            before.classList.add('field');
            form.appendChild(before);
        })();
        (function fileTypeInput() {
            const mime = document.createElement('input');
            mime.classList.add('field');
            mime.placeholder = 'Enter the mime type of the file';
            form.appendChild(mime);
        })();
        (function folderPathInput() {
            const folder = document.createElement('input');
            folder.id = 'folder';
            folder.placeholder = "Proide the path for the folder";
            folder.classList.add('field');
            if (localStorage.folderPath)
                folder.value = localStorage.folderPath;
            form.appendChild(folder);
        })();
        (function searchBtn() {
            const btn = document.createElement('button');
            form.appendChild(btn);
            btn.classList.add('button');
            btn.innerText = 'Search';
            btn.onclick = () => fetchAllDriveFiles(form);
        })();
        (function insertTable() {
            document.querySelector('table')?.remove();
            const table = document.createElement('table');
            form.insertAdjacentElement('afterend', table);
        })();
    })();
    async function fetchAllDriveFiles(form, record) {
        if (record)
            return manageFilesDatabase([], record, true); //We delete the record for the folder path
        try {
            await fetchAndFilter();
            spinner(false); //Hide the spinner
        }
        catch (error) {
            spinner(false); //Hide the spinner
            console.log(error);
            alert(error);
        }
        async function fetchAndFilter() {
            const files = await fetchAllFilesByBatches();
            if (!files)
                throw new Error('Could not fetch the files list from onedrive');
            const search = form.querySelector('#search');
            if (!search)
                throw new Error('Did not find the serch input');
            // Filter files matching regex pattern
            const matchingFiles = filterFiles(files, search.value);
            // Get reference to the table
            const table = document.querySelector('table');
            if (!table)
                throw new Error('The table element was not found');
            table.innerHTML = "<tr class =\"fileTitle\"><th>File Name</th><th>Created Date</th><th>Last Modified</th></tr>"; // Reset table
            const docFragment = new DocumentFragment();
            docFragment.appendChild(table); //We move the table to the docFragment in order to avoid the slow down related to the insertion of the rows directly in the DOM 
            for (const file of matchingFiles) {
                // Populate table with matching files
                const row = table.insertRow();
                row.classList.add('fileRow');
                row.insertCell(0).textContent = file.name;
                row.insertCell(1).textContent = new Date(file.createdDateTime).toLocaleString();
                row.insertCell(2).textContent = new Date(file.lastModifiedDateTime).toLocaleString();
                const link = await getDownloadLink(file.id);
                // Add double-click event listener to open file
                row.addEventListener("dblclick", () => {
                    window.open(link, "_blank");
                });
            }
            form.insertAdjacentElement('afterend', table);
            console.log(`Fetched ${files.length} items, displaying ${matchingFiles.length} matching files.`);
        }
        async function getDownloadLink(fileId) {
            const data = await JSONFromGETRequest(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`);
            return data.webUrl;
        }
        async function fetchAllFilesByBatches() {
            const path = byID('folder')?.value;
            if (!path)
                throw new Error('The file path could not be retrieved');
            const allFiles = [];
            const existing = await manageFilesDatabase(allFiles, path);
            if (existing.length)
                return existing;
            localStorage.folderPath = path;
            const select = '$select=name,id,folder,file,createdDateTime,lastModifiedDateTime';
            const top = '$top=900';
            await fetchAllFilesByPath(path);
            return await manageFilesDatabase(allFiles, path);
            async function fetchAllFilesByPath(path) {
                // Step 1: Get root-level files & folders
                path = path.replace('\\', '/');
                const topLevelItems = await fetchTopLevelFiles(path);
                const [files, folders] = getFilesAndFolders(topLevelItems);
                allFiles.push(...files);
                // Step 2: Filter folders & fetch their contents using $batch
                const folderIds = folders.map((f) => f.id);
                await fetchSubfolderContents(folderIds);
                console.log(`Fetched ${allFiles.length} files.`);
                return allFiles;
            }
            async function fetchTopLevelFiles(path) {
                const id = await getFolderIdByPath(path);
                const url = `https://graph.microsoft.com/v1.0/me/drive/items/${id}/children?${top}&${select}`;
                const data = await JSONFromGETRequest(url);
                return data.value; // Returns an array of files & folders
                async function getFolderIdByPath(path) {
                    const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${path}`;
                    const data = await JSONFromGETRequest(endpoint);
                    return data.id; // Folder ID
                }
            }
            async function fetchSubfolderContents(folderIds) {
                const batchUrl = "https://graph.microsoft.com/v1.0/$batch";
                // Create batch request for each folder
                const batchRequests = folderIds.map((folderId, index) => ({
                    id: `${index + 1}`,
                    method: "GET",
                    url: `/me/drive/items/${folderId}/children?${top}&${select}`,
                }));
                const limit = 20;
                for (let i = 0; i < batchRequests.length; i += limit) {
                    const batchData = await fetchRequests(batchRequests.slice(i, i + limit));
                    await processItems(batchData);
                }
                async function fetchRequests(requests) {
                    const body = { requests: requests };
                    const response = await graph.sendRequest(batchUrl, 'POST', body, undefined, "application/json", "Error fetching subfolders");
                    if (!response?.ok)
                        return;
                    return await response?.json();
                }
                async function processItems(data) {
                    // Extract file lists from batch responses
                    const items = data.responses.map((res) => res.body.value).flat();
                    const [files, folders] = getFilesAndFolders(items);
                    allFiles.push(...files);
                    const subfolderIds = folders.map((f) => f.id);
                    await fetchSubfolderContents(subfolderIds);
                }
            }
        }
        ;
        function getFilesAndFolders(items) {
            return [getFiles(items), subFolders(items)];
        }
        function subFolders(items) {
            return items.filter(item => item?.folder);
        }
        function getFiles(items) {
            return items.filter(item => item?.file);
        }
        async function JSONFromGETRequest(url) {
            const response = await graph.sendRequest(url, 'GET', undefined, undefined, undefined, 'Error fetching items from endpoint');
            if (!response?.ok)
                return;
            return await response.json();
        }
        ;
        function filterFiles(files, search) {
            const byName = files.filter((item) => RegExp(search, 'i').test(item.name));
            const created = (file) => new Date(file.createdDateTime);
            const after = form.querySelector('#after')?.valueAsDate;
            const before = form.querySelector('#before')?.valueAsDate;
            if (after && before)
                return byName.filter(file => created(file).getTime() > after.getTime() && created(file).getTime() < before.getTime());
            else if (before)
                return byName.filter(file => created(file).getTime() < before.getTime());
            else if (after)
                return byName.filter(file => created(file).getTime() > after.getTime());
            else
                return byName;
        }
        async function manageFilesDatabase(files, path, deleteRecord = false) {
            const dbName = "FileDatabase";
            const storeName = "Files";
            const dbVersion = 1;
            // Open (or create) the database
            const db = await new Promise((resolve, reject) => {
                const request = indexedDB.open(dbName, dbVersion);
                request.onupgradeneeded = function (event) {
                    const db = event.target?.result;
                    if (db.objectStoreNames.contains(storeName))
                        return;
                    db.createObjectStore(storeName, { keyPath: "path" });
                    console.log("Object store created successfully.");
                };
                request.onsuccess = function (event) {
                    resolve(event.target?.result);
                };
                request.onerror = function (event) {
                    reject("Failed to open database: " + event.target?.error);
                };
            });
            // Retrieve or add the entry
            return new Promise((resolve, reject) => {
                const transaction = db.transaction(storeName, "readwrite");
                const store = transaction.objectStore(storeName);
                // Check if an entry with the given path exists
                const getRequest = store.get(path);
                getRequest.onsuccess = function (event) {
                    const existingEntry = event.target?.result;
                    if (existingEntry && deleteRecord) {
                        const deleteRequest = store.delete(path);
                        deleteRequest.onsuccess = function () {
                            console.log('successfuly deleted the record');
                            resolve(files);
                        };
                        deleteRequest.onerror = function () {
                            reject("Failed to delete the specified record: " + event.target?.error);
                        };
                    }
                    else if (existingEntry) {
                        console.log("Entry found for path:", path);
                        resolve(existingEntry.files); // Return the existing files array
                    }
                    else if (!files.length) {
                        resolve(files); //We return the empty array
                    }
                    else {
                        // Add a new entry if it doesn't exist
                        const data = { path: path, files: files };
                        const addRequest = store.put(data);
                        addRequest.onsuccess = function () {
                            console.log("New entry added for path:", path);
                            resolve(files); // Return the newly added files array
                        };
                        addRequest.onerror = function (event) {
                            reject("Failed to add new entry: " + event.target?.error);
                        };
                    }
                };
                getRequest.onerror = function (event) {
                    reject("Failed to retrieve entry: " + event.target?.error);
                };
            });
        }
    }
}
//# sourceMappingURL=pwaVersion.js.map