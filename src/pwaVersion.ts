showMainUI();
function showMainUI(homeBtn?: boolean) {
    const container = byID('btns');
    if (!container) return;
    container.innerHTML = "";
    if (homeBtn) return appendBtn('home', 'Back to Main', showMainUI)

    appendBtn('entry', 'Add Entry', addNewEntry);
    appendBtn('invoice', 'Invoice', invoice);
    appendBtn('letter', 'Letter', issueLetter);
    appendBtn('lease', 'Leases', issueLeaseLetter);
    appendBtn('search', 'Search Files', searchFiles);
    appendBtn('settings', 'Settings', settings);

    function appendBtn(id: string, text: string, onClick: Function) {
        const btn = document.createElement('button');
        btn.id = id;
        btn.classList.add("ms-Button");
        btn.innerText = text;
        btn.onclick = () => onClick();
        container?.appendChild(btn);
        return btn
    }
};


async function getAccessToken() {
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

async function setLocalStorageTitles(graph: GraphAPI) {
    TableRows = await graph.fetchExcelTable(tableName, true) as string[][];//!We fetch the entire table inlcuding the headers row (we call the "/range" endpoint not the "/rows" endpoint)

    tableTitles = TableRows?.[0];
    if (!tableTitles) return [];
    localStorage.tableTitles = JSON.stringify(tableTitles);
    return tableTitles;
}
/**
 * 
 * @param {boolean} add - If false, the function will only show a form containing input fields for the user to provide the data for the new row to be added to the Excel Table. If true, the function will parse the values from the input fields in the form, and will add them as a new row to the Excel Table. Its default value is false.
 * @param {boolean} display - If provided, the function will show the visible rows in the UI after the new row has been added.
 * @param {any[]} row - If provided, the function will add the row directly to the Excel Table without needing to retrieve the data from the inputs.
 */
async function addNewEntry(add: boolean = false, row?: any[]) {
    const workbookPath = getAccountsWorkBookPath();
    if (!workbookPath) return alert('Could not get a valid workbook path from the localStorage');
    accessToken = await getAccessToken() || '';
    if (!accessToken) return alert('The access token is missing. Check the console.log for more details');
    const graph = new GraphAPI(accessToken, workbookPath);
    if (!workbookPath || !tableName) throw new Error('The Excel Workbook path and/or the name of the Excel Table are missing or invalid');

    (async function showAddNewForm() {
        if (add) return;
        document.querySelector('table')?.remove();
        spinner(true);//We show the spinner
        try {
            await createForm();
        } catch (error) {
            spinner(false);//We hide the sinner
            alert(error);
        }

        async function createForm() {
            const sessionId = await graph.createFileSession();
            if (!sessionId) throw new Error('There was an issue with the creation of the file cession. Check the console.log for more details');
            if (!tableTitles || !TableRows) tableTitles = await setLocalStorageTitles(graph);
            const tableBody = TableRows.slice(1, -1);
            const inputs: HTMLInputElement[] = [];
            const bound = (indexes: number[]) => inputs.filter(input => indexes.includes(getIndex(input))).map(input => [input, getIndex(input)]) as InputCol[];

            insertAddForm(tableTitles);
            await graph.closeFileSession(sessionId);
            spinner(false);//We hide the spinner


            function insertAddForm(titles: string[]) {
                if (!titles) throw new Error('The table titles are missing. Check the console.log for more details');


                const form = byID();
                if (!form) throw new Error('Could not find the form element');
                form.innerHTML = '';

                const divs = titles.map((title, index) => {
                    const div = newDiv(index);
                    if (![4, 7].includes(index))
                        div.appendChild(createLable(title, index));//We exclued the labels for "Total Time" and for "Year"
                    div.appendChild(createInput(index));
                    return div;
                });

                (function groupDivs() {
                    [
                        [11, 12, 13],//"Moyen de Paiement", "Compte", "Tiers"
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

                function newDiv(i: number, divs?: HTMLDivElement[], css: string = "block") {
                    if (divs) return groupDivs();
                    else return create();

                    function create() {
                        const div = document.createElement('div');
                        div.dataset.block = i.toString();
                        form?.appendChild(div);
                        div.classList.add(css);
                        return div;
                    }

                    function groupDivs() {
                        const div = newDiv(i, undefined, "group") as HTMLDivElement;
                        divs?.forEach(el => div.appendChild(el));
                        form?.children[3]?.insertAdjacentElement('afterend', div);
                        return div
                    }
                }

                function createLable(title: string, i: number) {
                    const label = document.createElement('label');
                    label.htmlFor = 'input' + i.toString();
                    label.innerHTML = title + ':';
                    return label
                }


                function createInput(index: number) {
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
                        else if ([4, 7].includes(index)) input.style.display = 'none';//We hide those 2 columns: 'Total Time' and the 'Year'

                        (function addDataLists() {
                            const updateNext = [0, 1, 8, 15]//Those are the indexes of the inputs (i.e; the columns numbers) that need to get an onChange event in order to update the dataLists of the next inputs when the current input is changed: "Client"(0), "Affaire"(1), "Taux Horaire"(8), "Adresses"(15)

                            if (updateNext.includes(index)) input.onchange = () => inputOnChange(index, bound(updateNext), tableBody, false);

                            if (![0, 2, 11, 12, 13].includes(index)) return;//We will initially populate the "Client"(0), Nature(2), "Payment Method"(11), "Bank Account"(12), "Third Party"(13) lists only, the other inputs will be populate when the onChange function will be called
                            populateSelectElement(input, getUniqueValues(index, tableBody));
                        })();

                        (function addRestOnChange() {
                            if (index < 5 || index > 10) return;
                            //Only for the  "Start Time", "End Time", "Total Time", "Hourly Rate", "Amount", "VAT" columns . The "Total Time" input (7) is hidden, so it can't be changed by the user. We will add the onChange event to it by simplicity
                            const reset = [5, 6, 7, 9, 10];//Those are "Start Time" (5), "End Time" (6), "Total Time" (7, although it is hidden), "Amount" (9), "VAT" (10) columns. We exclude the "Hourly Rate" column (8). We let the user rest it if he wants
                            input.onchange = () => resetInputs(bound(reset), index);//!We are passing the table[][] argument as undefined, and the invoice argument as false which means that the function will only reset the bound inputs without updating any data list

                        })();
                    })();

                    return input
                }

                function resetInputs(inputs: [HTMLInputElement, number][], index: number) {
                    inputs
                        .filter(([input, index]) => index > index)
                        .forEach(([input, index]) => reset(input));//We reset any input which dataset-index is > than the dataset-index of the input that has been changed

                    if (index === 9)
                        inputs
                            .filter(([input, index]) => index < index)
                            .forEach(([input], index) => reset(input));//If the input is the input for the "Montant" column of the Excel table, we also reset the "Start Time" (5), "End Time" (6) and "Hourly Rate" (7) columns' inputs. We do this because we assume that if the user provided the amount, it means that either this is not a fee, or the fee is not hourly billed.

                    function reset(input: HTMLInputElement) {
                        if (!input) return;
                        input.value = '';
                        if (input.valueAsNumber) input.valueAsNumber = 0;
                    }
                };
            }
        }
    })();

    (async function addEntry() {
        if (!add) return;
        spinner(true);//We show the spinner
        const display = !row?.length;
        if (!row) row = parseInputs() || [];
        try {
            const visibleCells = await addRow(row);
            if (visibleCells?.length) displayVisibleCells(visibleCells, display);
            spinner(false);//We hide the spinner
        } catch (error) {
            spinner(false);//We hide the spinner
            alert(error)
        }

        function parseInputs() {
            const colNature = 2, colDate = 3, colStart = 5, colEnd = 6, colRate = 8, colAmount = 9, colVAT = 10;
            const stop = (missing: string) => alert(`${missing} missing. You must at least provide the client, matter, nature, date and the amount. If you provided a time start, you must provide the end time and the hourly rate. Please review your iputs`);

            const inputs = Array.from(document.getElementsByTagName('input')) as HTMLInputElement[];//all inputs

            const nature = getInputByIndex(inputs, colNature)?.value;
            if (!nature) return stop('The matter is')
            const date = getInputByIndex(inputs, colDate)?.valueAsDate as Date | undefined;
            if (!date) return stop('The invoice date is');
            const amount = getInputByIndex(inputs, colAmount) as HTMLInputElement;
            const rate = getInputByIndex(inputs, colRate)?.valueAsNumber || 0;

            const debit = ['Honoraire', 'Débours/Dépens', 'Débours/Dépens non facturables', 'Rétrocession d\'honoraires', 'Charges déductibles'].includes(nature);//We check if we need to change the value sign

            const row =
                inputs.map((input, index) => getInputValue(index));//!CAUTION: The html inputs are not arranged according to their dataset.index values. If we follow their order, some values will be assigned to the wrong column of the Excel table. That's why we do not pass the input itself or the dataset.index of the input to getInputValue(), but instead we pass the index of the column for which we want to retrieve the value from the relevant input.

            if (missing()) return stop('Some of the required fields are');

            return row

            function getInputValue(index: number) {
                const input = getInputByIndex(inputs, index) as HTMLInputElement;
                if ([colDate, colDate + 1].includes(index))
                    return getISODate(date);//Those are the 2 date columns
                else if ([colStart, colEnd].includes(index))
                    return getTime([input]);//time start and time end columns
                else if (index === 7) {
                    //!This is a hidden input
                    const timeInputs = [colStart, colEnd].map(i => getInputByIndex(inputs, i));
                    const totalTime = getTime(timeInputs);//Total time column
                    if (totalTime && rate && !amount.valueAsNumber)
                        amount.valueAsNumber = totalTime * 24 * rate// making the amount equal the rate * totalTime
                    return totalTime
                }
                else if (debit && index === colAmount)
                    return -input.valueAsNumber || 0;//This is the amount if negative
                else if ([colRate, colAmount, colVAT].includes(index))
                    return input.valueAsNumber || 0;//Hourly Rate, Amount, VAT
                else return input.value;

            }

            function missing() {
                if (row.filter((value, i) => (i < colDate + 1 || i === colAmount) && !value).length > 0) return true;//if client name, matter, nature, date or amount are missing
                //else if (row[9]) return [5, 6,7,8].map(index => row[index] = 0).length < 1;//This means the amount has been provided and does not  depend on the time spent or the hourly rate. We set the values of the startTime and endTime to 0, and return false (length<1 must return false)

                if (row[colStart] === row[colEnd]) return false;//If the total time = 0 we do not need to alert if the hourly rate is missing
                else if (row[colStart] && (!row[colEnd] || !row[colRate]))
                    return true//if startTime is provided but without endTime or without hourly rate
                else if (row[colEnd] && (!row[colStart] || !row[colRate]))
                    return true//if endTime is provided but without startTime or without hourly rate
            };
        }

        async function addRow(row: any[] | undefined) {
            if (!row) throw new Error('The row is not valid');
            const visibleCells = await graph.addRowToExcelTable(row, TableRows.length - 2, tableName, true);

            if (!visibleCells?.length)
                return alert('There was an issue with the adding or the filtering, check the console.log for more details');

            alert('Row aded and the table was filtered');
            return visibleCells

        };
        function displayVisibleCells(visibleCells: string[][], display: boolean) {
            if (!display) return;
            const tableDiv = createDivContainer();
            const table = document.createElement('table');
            table.classList.add('table');
            tableDiv.appendChild(table);

            const columns = [0, 1, 2, 3, 7, 8, 9, 10, 14, 15];//The columns that will be displayed in the table;
            const rowClass = 'excelRow';
            (function insertTableHeader() {
                if (!tableTitles) throw new Error('No Table Titles');
                const headerRow = document.createElement('tr');
                headerRow.classList.add(rowClass);
                const thead = document.createElement('thead');
                table.appendChild(thead);
                thead.appendChild(headerRow);
                tableTitles.forEach((cell, index) => {
                    if (!columns.includes(index)) return;
                    addTableCell(headerRow, cell, 'th');
                });
            })();
            (function insertTableRows() {
                const tbody = document.createElement('tbody');
                table.appendChild(tbody);
                visibleCells.forEach((row, index) => {
                    if (index < 1) return;//We exclude the header row
                    if (!row) return;
                    const tr = document.createElement('tr');
                    tr.classList.add(rowClass);
                    tbody.appendChild(tr);
                    row.forEach((cell, index) => {
                        if (!columns.includes(index)) return;
                        addTableCell(tr, cell, 'td');
                    });
                });
            })();


            const form = byID();
            if (!form) throw new Error('The form element was not found');
            if (form) {
                form?.insertAdjacentElement('afterend', tableDiv);
            }

            function createDivContainer() {
                const id = 'retrieved';
                let tableDiv = byID(id)
                if (tableDiv) {
                    tableDiv.innerHTML = '';
                    return tableDiv;
                };
                tableDiv = document.createElement('div');
                tableDiv.classList.add('table-div');
                tableDiv.id = id;
                return tableDiv;
            }

            function addTableCell(parent: HTMLElement, text: string, tag: string) {
                const cell = document.createElement(tag);
                //   cell.classList.add(css);
                cell.textContent = text;
                parent.appendChild(cell);
            }
        };

    })();

};

// Update Word Document
async function invoice(issue: boolean = false) {
    const workbookPath = getAccountsWorkBookPath();
    if (!workbookPath) return alert('Could not get a valid workbook path from the localStorage');
    accessToken = await getAccessToken() || '';
    if (!accessToken) return alert('The access token is missing. Check the console.log for more details');
    const graph = new GraphAPI(accessToken, workbookPath);
    (async function showInvoiceForm() {
        if (issue) return;
        if (!workbookPath || !tableName) return alert('The Excel Workbook path and/or the name of the Excel Table are missing or invalid');
        document.querySelector('table')?.remove();
        spinner(true);//We show the spinner
        try {
            await createForm();
        }
        catch (error) {
            spinner(false);//We hide the spinner
            alert(error)
        }

        async function createForm() {
            const sessionId = await graph.createFileSession() || '';
            if (!sessionId) throw new Error('There was an issue with the creation of the file cession. Check the console.log for more details');
            if (!tableTitles || !TableRows) tableTitles = await setLocalStorageTitles(graph);

            insertInvoiceForm(tableTitles);
            await graph.closeFileSession(sessionId);
            spinner(false);//We hide the spinner
        }

        function insertInvoiceForm(tableTitles: string[]) {
            if (!tableTitles) throw new Error('The table titles are missing. Check the console.log for more details');
            const form = byID();
            if (!form) throw new Error('The form element was not found');
            form.innerHTML = '';
            const tableBody = TableRows.slice(1, -1);
            const boundInputs: InputCol[] = [];

            insertInputsAndLables([0, 1, 2, 3, 3], 'input');//Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice

            insertInputsAndLables(['Discount'], 'discount')[0].value = '0%'; //Inserting a discount percentage input and setting its default value to 0%

            insertInputsAndLables(['Français', 'English'], 'lang', true); //Inserting languages checkboxes

            (function customizeDateLabels() {
                const [from, to] = Array.from(document.getElementsByTagName('label'))
                    ?.filter(label => label.htmlFor.endsWith('3'));
                if (from) from.innerText += ' From (included)';
                if (to) to.innerText += ' To/Before (included)';
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

            function insertInputsAndLables(indexes: (number | string)[], id: string, checkBox: boolean = false): HTMLInputElement[] {
                let css = 'field';
                if (checkBox) css = 'checkBox';
                return indexes.map((index) => {
                    const div = newDiv(String(index));
                    appendLable(index, div);
                    return appendInput(index, div);
                });

                function appendInput(index: number | string, div: HTMLDivElement) {
                    const NaN = isNaN(Number(index));
                    const input = document.createElement('input');
                    input.classList.add(css);
                    !NaN ? input.id = id + index.toString() : input.id = id;

                    (function setType() {
                        if (checkBox) input.type = 'checkbox';
                        else if (NaN || index as number < 3) input.type = 'text';
                        else input.type = 'date';
                    })();

                    (function notCheckBox() {
                        if (NaN || checkBox) return;//If the index is not a number or the input is a checkBox, we return;
                        index = Number(index);
                        input.name = input.id;
                        input.dataset.index = index.toString();
                        if (index < 3)
                            boundInputs.push([input, index]);//Fields "Client"(0), "Affaire"(1), "Nature"(2) are the inputs that will need to get their dataList created or updated each time the previous input is changed.
                        if (index < 2)
                            input.onchange = () => inputOnChange(index as number, boundInputs, tableBody, true);//We add onChange on "Client" (0) and "Affaire" (1) columns
                        if (index < 1)
                            populateSelectElement(input, getUniqueValues(0, tableBody));//We create a unique values dataList for the "Client" (0) input
                    })();

                    (function isCheckBox() {
                        if (!checkBox) return;
                        input.dataset.language = index.toString().slice(0, 2).toUpperCase();
                        input.onchange = () =>
                            Array.from(document.getElementsByTagName('input'))
                                .filter((checkBox: HTMLInputElement) => checkBox.dataset.language && checkBox !== input)
                                .forEach(checkBox => checkBox.checked = false);
                    })();
                    div.appendChild(input);
                    return input;
                }

                function appendLable(index: number | string, div: HTMLDivElement) {
                    const label = document.createElement('label');
                    isNaN(Number(index)) || checkBox ? label.innerText = index.toString() : label.innerText = tableTitles[Number(index)];
                    !isNaN(Number(index)) ? label.htmlFor = id + index.toString() : label.htmlFor = id;
                    div?.appendChild(label);
                }

                function newDiv(i: string, css: string = "block") {
                    const div = document.createElement('div');
                    div.dataset.block = i;
                    form?.appendChild(div);
                    div.classList.add(css);
                    return div;
                }
            };

        }

    })();

    (async function issueInvoice() {
        if (!issue) return;
        if (!templatePath || !destinationFolder) return alert('The full path of the Word Invoice Template and/or the destination folder where the new invoice will be saved, are either missing or not valid');
        spinner(true);//We show the spinner

        try {
            await editInvoice();
        } catch (error) {
            spinner(false);//We hide the sinner
            alert(error)
        }

    })();

    async function editInvoice() {
        const client = tableTitles[0], matter = tableTitles[1];//Those are the 'Client' and 'Matter' columns of the Excel table
        const sessionId = await graph.createFileSession(true) || '';//!persist must be = true because we might add a new row if there is a discount. If we don't persist the session, the table will be filtered and the new row will not be added.
        if (!sessionId) throw new Error('There was an issue with the creation of the file cession. Check the console.log for more details');

        const inputs = Array.from(document.getElementsByTagName('input'));
        const criteria = inputs.filter(input => getIndex(input) >= 0);

        const discount = parseInt(inputs.find(input => input.id === 'discount')?.value || '0%');

        const lang = inputs.find(input => input.dataset.language && input.checked === true)?.dataset.language || 'FR';

        const date = new Date();//We need to generate the date at this level and pass it down to all the functions that need it
        const invoiceNumber = getInvoiceNumber(date);
        const data = await filterExcelData(criteria, discount, lang, invoiceNumber);

        if (!data) throw new Error('Could not retrieve the filtered Excel table');

        const [wordRows, totalsLabels, clientName, matters, adress] = data;

        const invoice = {
            number: invoiceNumber,
            clientName: clientName,
            matters: matters,
            adress: adress,
            lang: lang
        }

        const contentControls = getContentControlsValues(invoice, date);

        const fileName = getInvoiceFileName(clientName, matters, invoiceNumber);
        let savePath = `${destinationFolder}/${fileName}`;

        savePath = prompt(`The file will be saved in ${destinationFolder}, and will be named : ${fileName}.\nIf you want to change the path or the name, provide the full file path and name of your choice without any sepcial characters`, savePath) || savePath;

        (async function editInvoiceFilterExcelClose() {
            await graph.createAndUploadWordDocument(templatePath, savePath, lang, 'Invoice', wordRows, contentControls, totalsLabels);

            await graph.filterExcelTable(tableName, matter, matters, sessionId);//We filter the table by the matters that were invoiced

            await graph.closeFileSession(sessionId);
            spinner(false);//We hide the spinner
        })();

        /**
         * Filters the Excel table according to the values of each inputs, then returns the values of the Word table rows that will be added to the Word table in the invoice template document 
         * @param {HTMLInputElement[]} inputs - the html inputs containing the values based on which the table will be filtered
         * @param {number} discount  - The discount percentage that will be applied to the amount of each invoiced row if any. It is a number between 0 and 100. If it is equal to 0, it means that no discount will be applied.
         * @param {string} lang - The language in which the invoice will be issued 
         * @returns {Promise<[string[][], string[], string[], string[]]>} - The values of the rows that will be added to the Word table in the invoice template
         */
        async function filterExcelData(inputs: HTMLInputElement[], discount: number, lang: string, invoiceNumber: string): Promise<[string[][], string[], string, string[], string[]] | void> {
            const matterCol = 1, dateCol = 3, addressCol = 15;//Indexes of the 'Matter' and 'Date' columns in the Excel table
            const clientName = getInputByIndex(inputs, 0)?.value || '';
            const matters =
                getArray(getInputByIndex(inputs, matterCol)?.value) || []; //!The Matter input may include multiple entries separated by ', ' not only one entry.

            if (!clientName || !matters?.length) throw new Error('could not retrieve the client name or the matter/matters list from the inputs');

            await graph.clearFilterExcelTable(tableName, sessionId);//We unfilter the table;

            //Filtering by Client (criteria[0])
            await graph.filterExcelTable(tableName, client, [clientName], sessionId);

            let visible = await graph.getVisibleCells(tableName, sessionId) as any[][];

            if (!visible) {
                return alert('Could not retrieve the visible cells of the Excel table');
            }
            visible = visible.slice(1, - 1);//We exclude the first and the last rows of the table. Since we are calling the "range" endpoint, we get the whole table including the headers. The first row is the header, and the last row is the total row.

            const adresses = getUniqueValues(addressCol, visible) as string[];//!We must retrieve the adresses at this stage before filtering by "Matter" or any other column

            visible = visible.filter(row => matters.includes(row[matterCol]));
            //We finaly filter by date
            visible = filterByDate(visible, dateCol);
            const [wordRows, totalLabels] = getRowsData(visible, discount, lang, invoiceNumber);

            return [wordRows, totalLabels, clientName, matters, adresses];

            function filterByDate(visible: string[][], date: number) {

                const convertDate = (date: string | number) => dateFromExcel(Number(date)).getTime();

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

async function issueLetter(create: boolean = false) {
    accessToken = await getAccessToken() || '';
    (function showForm() {
        if (create) return;
        document.querySelector('table')?.remove();
        const form = byID();
        if (!form) return;
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
            btn.innerText = 'Créer lettre'
            btn.onclick = () => issueLetter(true);
        })();

        (function homeBtn() {
            showMainUI(true);
        })();
    })();

    (async function generate() {
        if (!create) return;
        const input = byID('textInput') as HTMLTextAreaElement;
        if (!input) return;
        const templatePath = "Legal/Mon Cabinet d'Avocat/Administratif/Modèles Actes/Template_Lettre With Letter Head [DO NOT MODIFY].docx";
        const fileName = prompt('Provide the file name without special characthers');
        if (!fileName) return;
        const filePath = `${prompt('Provide the destination folder', "Legal/Mon Cabinet d'Avocat/Clients")}/${fileName}.docx`;
        if (!filePath) return;

        const contentControls = [['RTCoreText', input.value], ['RTReference', 'Référence'], ['RTClientName', 'Nom du Client'], ['RTEmail', 'Email du client']];

        new GraphAPI('', filePath).createAndUploadWordDocument(templatePath, filePath, 'FR', undefined, undefined, contentControls);
    })();

}
async function issueLeaseLetter(create: boolean = false) {
    spinner(true);
    accessToken = await getAccessToken() || '';
    if (!localStorage.leasesPath)
        localStorage.leasesPath = prompt('Please provide the OneDrive full path (including the file name and extension) for the Excel Workbook', "Legal/Mon Cabinet d'Avocat/Clients/LeasesDataBase.xlsm");
    const workbookPath = localStorage.leasesPath || alert('The excel Workbook path is not valid');
    const tableName: string = 'LEASES';
    const graph = new GraphAPI(accessToken, workbookPath);


    const Ctrls: LeaseCtrls = {
        owner: { title: 'RTBailleur', col: 0, label: 'Nom du Bailleur', type: 'select', value: '' },
        adress: { title: 'RTAdresseDestinataire', label: 'Adresse du bien loué', col: 1, type: 'select', value: '' },
        tenant: { title: 'RTLocataire', label: 'Nom du Locataire', col: 2, type: 'select', value: '' },
        leaseDate: { title: 'RTDateBail', label: 'Date du Bail', col: 3, type: 'date', value: '' },
        leaseType: { title: 'RTNature', label: 'Nature du Bail', col: 4, type: 'text', value: '' },
        initialIndex: { title: 'RTIndiceInitial', label: 'Indice initial', col: 5, type: 'text', value: '' },
        indexQuarter: { title: 'RTTrimestre', label: 'Trimestre de l\'indice', col: 6, type: 'text', value: '' },
        initialIndexDate: { title: 'RTIndiceInitialDate', label: 'Date de l\'indice initial', col: 7, type: 'date', value: '' },
        baseIndex: { title: 'RTIndiceBase', label: `Indice de référence`, col: 8, type: 'text', value: '' },
        baseIndexDate: { title: 'RTDateIndiceBase', label: `Date de l'indice de référence`, col: 9, type: 'date', value: '' },
        index: { title: 'RTIndice', label: 'Indice de révision', col: 10, type: 'text', value: '' },
        indexDate: { title: 'RTDateIndice', label: 'Date de l\'indice de révision', col: 11, type: 'date', value: '' },
        currentLease: { title: 'RTLoyerActuel', label: 'Loyer Actuel (ou révisé)', col: 12, type: 'text', value: '' },
        revisionDate: { title: 'RTDate', label: 'Date de la dernière Révision', col: 13, type: 'date', value: '' },
        initialYear: { title: 'RTIndiceInitialAnnée', type: 'text', value: '' },
        revisionYear: { title: 'RTYear', type: 'text', value: '' },
        baseYear: { title: 'RTPreviousYear', type: 'text', value: '' },
        newLease: { title: 'RTLoyerNouveau', type: 'text', value: '' },
        nextRevision: { title: 'RTNextRevision', type: 'text', value: '' },
    };

    /**
     * This function casts the "col" property as "number" beacause the col property of some RTs is "undefined". So this function will mainly cast the col property to a number in order to avoid casting each time we retrieve the col property
     * @param {RT} RT 
     * @returns {number}
     */
    const column = (RT: RT) => RT.col as number;

    const ctrls = Object.values(Ctrls);

    const findRT = (id: string) => ctrls.find(RT => RT.title === id);

    let row: any[] = [], rowIndex: number | null = null;
    (async function showForm() {
        if (create) return;
        const inputs: InputCol[] = [];
        const findInput = (id: string) => inputs.find(([input, col]) => input.id === id)?.[0];
        const tableRows = await graph.fetchExcelTable(tableName, false, false) as any[][];//We are calling the "rows" endpoint which returns the table rows without the headers.
        if (!tableRows) return;
        document.querySelector('table')?.remove();
        const form = byID();
        if (!form) return;
        form.innerHTML = '';
        const divs: HTMLDivElement[] = [];

        (function insertInputs() {
            const unvalid = (values: (string | undefined)[]) => values.find(value => !value || isNaN(Number(value)));
             ctrls
                .filter(RT => !isNaN(column(RT)))
                .map(RT => inputs.push([createInput(RT), column(RT)] as const));
            
            const owner = findInput(Ctrls.owner.title);
            if(owner) populateSelectElement(owner, getUniqueValues(column(Ctrls.owner), tableRows), false);

            (function inputsOnChange() {
                const filled = inputs.filter(([input, col]) => col <= column(Ctrls.tenant));
                filled.forEach(([input, col]) => input.onchange = () => inputOnChange(col, inputs, tableRows, false));

                const index = findInput(Ctrls.index.title);
                if (index) index.onchange = () => {
                    const filtered = filterTableByInputsValues(filled, tableRows);
                    if (!filtered?.length) {
                        return prompt('No lease having owner name, property adress and tenant name as in the inputs was found')
                    } else if (filtered.length > 1) {
                        return prompt('Multiple leases having owner name, property adress and tenant name as in the inputs were found. Please provide the number of the lease you want to select :\n' + filtered.map((row, index) => `${index + 1} : ${row.join(', ')}`).join('\n'))
                    }
                    row = filtered[0];
                    rowIndex = tableRows.indexOf(row);
                    const base = row[column(Ctrls.baseIndex)] || row[column(Ctrls.initialIndex)];
                    const latestIndex = index.value;
                    const currentLease = row[column(Ctrls.currentLease)];
                    if (unvalid([base, latestIndex, currentLease])) return alert('Please make sure that the values of the current lease, the base indice and the new indice are all provided and valid numbers');
                    const currentLeaseInput = findInput(Ctrls.currentLease.title);
                    if (!currentLeaseInput) return alert('Current lease input not found');
                    const newLease = (Number(currentLease) * (Number(latestIndex) / Number(base))).toFixed(2).toString();
                    currentLeaseInput.value = newLease;//This will show the value of the new lease after applying the calculation
                    Ctrls.baseIndex.value = latestIndex;
                    Ctrls.newLease.value = newLease;//We update the new lease RT
                };
            })();

        })();

        (function groupDivs() {
            [
                [0, 1, 2],//"Bailleur"(0), "Adresse"(1), "Locataire"(2)
                [3, 4],//"Date du Bail"(3), "Nature du Bail"(4)
                [5, 6, 7],//"Indice Initial"(5), "Date de l'indice initial"(6), "Trimestre de l'indice"(7)
                [8, 9],//"Indice de référence"(8), "Date de l'indice de référence"(9)
                [10, 11],//"Indice de révision"(10), "Date de l'indice de révision"(11)
                [12, 13],//"Loyer Actuel (ou révisé)"(12), "Date de la dernière Révision"(13)
            ]
                .forEach((group, index) => groupDivs(divs.filter(div => group.includes(getIndex(div))), index));

            function groupDivs(divs: HTMLDivElement[], i: number) {
                const div = document.createElement('div');
                div.classList.add("group");
                div.dataset.block = i.toString();
                divs?.forEach(el => div.appendChild(el));
                form?.appendChild(div);
                return div
            }
        })();

        (function generateBtn() {
            const btn = document.createElement('button');
            form?.appendChild(btn);
            btn.classList.add('button');
            btn.innerText = 'Créer lettre'
            btn.onclick = () => generate(inputs);
        })();

        (function homeBtn() {
            showMainUI(true);
        })();
        spinner(false);

        function createInput(RT: RT, className: string = 'field') {
            const id = RT.title;
            const div = document.createElement('div');
            form?.appendChild(div);
            const append = (el: HTMLElement) => div.appendChild(el);

            (function appendLabel() {
                if (!RT.label) return;
                const label = document.createElement('label');
                label.htmlFor = id;
                label.innerText = RT.label;
                append(label);
            })();
            return appendInput();
            function appendInput() {
                const input = document.createElement('input') as HTMLInputElement;
                input.type = RT.type || 'text';
                input.id = id;
                input.classList.add(className);
                const col = column(RT).toString();
                input.dataset.index = col;
                div.dataset.index = col;
                append(input);
                divs.push(div);
                return input as HTMLInputElement
            };
        };
    })();

    async function generate(inputs: InputCol[]) {
        if (!inputs.length) return;
        const findInput = (id: string) => inputs.find(([input, col]) => input.id === id)?.[0];
        const templatePath = "Legal/Mon Cabinet d'Avocat/Administratif/Modèles Actes/Template_Révision de loyer [DO NOT MODIFY].docx";
        const date = new Date();
        const fileName = prompt('Provide the file name without special characthers');
        if (!fileName) return;
        const savePath = `${prompt('Provide the destination folder', "Legal/Mon Cabinet d'Avocat/Clients")}/${fileName}_${date.getFullYear()}${date.getMonth() + 1}${date.getDate()}@${date.getHours()}-${date.getMinutes()}.docx`;
        if (!savePath) return;

        inputs.map(([input, col]) => {
            const id = input.id
            if (id === Ctrls.currentLease.title) return;//!We don't update the value of current lease from the input because the value of the input is the new lease not the old one
            const RT = findRT(id);
            if (!RT) return;
            if ([Ctrls.leaseDate, Ctrls.indexDate, Ctrls.baseIndexDate, Ctrls.initialIndexDate].includes(RT)) RT.value = getDateString(input.valueAsDate as Date);
            else RT.value = input.value
        });
        const year = date.getFullYear();
        Ctrls.revisionDate.value = getDateString(date);
        Ctrls.revisionYear.value = year.toString();
        Ctrls.baseYear.value = (year - 1).toString();
        Ctrls.nextRevision.value = (year + 1).toString();
        const contentControls = ctrls.map(RT => [RT.title, RT.value]);


        graph.createAndUploadWordDocument(templatePath, savePath, 'FR', undefined, undefined, contentControls);

        (function updateLeasesTable() {

            (async function updateRow() {
                if (!rowIndex) return;
                const col = Ctrls.revisionDate.col as number;
                const revisionDate = Ctrls.initialIndexDate.value.replace(Ctrls.initialYear.value, Ctrls.revisionYear.value);
                row[col] = new Date(revisionDate);
                const input = findInput(Ctrls.revisionDate.title)
                if (input) input.value = revisionDate;
                await graph.updateExcelTableRow(tableName, rowIndex, row)
            })();
            (async function newRow() {
                if (rowIndex) return;
                row = [];
                inputs.forEach(([input, col]) => row[col] = input.value);
                await graph.addRowToExcelTable(row, rowIndex, tableName)
            });
        })();
        spinner(false);
    }

}


/**
 * Updates the data list or the value of bound inputs according to the value of the input that has been changed
 * @param {number} index - the dataset.index of the input that has been changed
 * @param {any[][]} table - The table that will be filtered to update the data list of the button. If undefined, it means that the data list will not be updated.
 * @param {boolean} invoice - If true, it means that we called the function in order to generate an invoice. If false, we called it in order to add a new entry in the table
 * @returns 
 */
function inputOnChange(index: number, inputs: InputCol[], table: any[][] | undefined, invoice: boolean) {
    if (!table?.length) return;

    const filledInputs =
        inputs
            .filter(([input, col]) => input.value && col <= index)//Those are all the inputs that the user filled with data

    const filtered = filterTableByInputsValues(filledInputs, table);//We filter the table based on the filled inputs

    if (!filtered.length) return;

    const boundInputs = inputs.filter(([input, col]) => col > index)//Those are the inputs for which we want to create  or update their data lists

    for (const [input, col] of boundInputs) {
        input.value = ''; //We reset the value of all bound inputs.
        const list = getUniqueValues(col, filtered);
        if (fillBound(list, input)) break;//!If the function returns true, it means that we filled the value of all the bound inputs, so we break the loop. If it returns false, it means that there is more than one value in the list, so we need to create or update the data list of the input.
        const combine = (invoice && [1, 2].includes(col))//For the "Matter" and "Nature" lists, we add a new element combining all the values separated by ","

        populateSelectElement(input, list, combine);
    }

    function fillBound(list: any[], input: HTMLInputElement): boolean | void {
        if (list.length > 1) return false;
        const value = list[0], found = filtered.length < 2;
        if (!found) return setValue(input, value);//If the filtered array contains more than one row with the same unique value in the corresponding column, we will not fill the next inputs
        const row = filtered[0];//This is the unique row in the filtered list, we will use it to fill all the other inputs
        boundInputs.forEach(([input, col]) => setValue(input, row[col]));
        return found;
    }

    function setValue(input: HTMLInputElement, value: any) {
        if (input.type === "date")
            input.valueAsDate = dateFromExcel(value);//!We must convert the dates from Excel
        else input.value = value?.toString() || '';
    };
};
/**
 * Filters the table according to the values of the inputs. The value of each input is compared to the value of the cell in the corresponding column in the table. If the value of the input is included in the cell value, it means that this row matches the criteria of this input. For a row to be included in the resulting filtered table, it must match the criteria of all the inputs.
 * @param {HTMLInputElement[]} inputs - the html inputs containing the values based on which the table will be filtered
 * @param {any[][]} table - The table that will be filtered
 * @returns {any[][]} - The resulting filtered table
 */
function filterTableByInputsValues(inputs: InputCol[], table: any[][]): any[][] {
    const values = inputs.map(([input, index]) => [index, input.value.split(splitter)] as const);//!some inputs may contain multiple comma separated values if the user has selected more than one option in the data list. So we split the input value by ", " and we check if the cell value is included in the resulting array.
    return table.filter(row => values.every(([index, value]) => value.includes(row[index])));
};


/**
 * Convert the date in an Excel row into a javascript date (in milliseconds)
 * @param {number} excelDate - The date retrieved from an Excel cell
 * @returns {Date} - a javascript format of the date
 */
function dateFromExcel(excelDate: number): Date {
    const date = new Date((excelDate - 25569) * (60 * 60 * 24) * 1000);//This gives the days converted from milliseconds. 

    const dateOffset = date.getTimezoneOffset() * 60 * 1000;//Getting the difference in milleseconds
    return new Date(date.getTime() + dateOffset);
}

function getMSGraphClient(accessToken: string) {
    //@ts-expect-error
    return MicrosoftGraph.Client.init({
        //@ts-expect-error
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}

function getNewExcelRow(inputs: HTMLInputElement[]) {
    return inputs.map(input => input.value)
}


function searchFiles() {
    (function showForm() {
        const form = byID('form') as HTMLDivElement;
        if (!form) return;
        form.innerHTML = '';
        if (localStorage.folderPath)
            fetchAllDriveFiles(form, localStorage.folderPath);//We will delete the record for this folder path from the database
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
            if (localStorage.folderPath) folder.value = localStorage.folderPath;
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

    async function fetchAllDriveFiles(form: HTMLDivElement, record?: string) {
        if (record) return manageFilesDatabase([], record, true);//We delete the record for the folder path
        if (!accessToken)
            accessToken = await getAccessToken() || '';
        if (!accessToken) return alert('The access token is missing. Check the console.log for more details');
        spinner(true);//We show the spinner

        try {
            await fetchAndFilter();
        } catch (error) {
            spinner(false);//Hide the spinner
            alert(error)
        }


        async function fetchAndFilter() {
            const files = await fetchAllFilesByBatches();
            if (!files) throw new Error('Could not fetch the files list from onedrive');
            const search = form.querySelector('#search') as HTMLInputElement;
            if (!search) throw new Error('Did not find the serch input');
            // Filter files matching regex pattern
            const matchingFiles = filterFiles(files, search.value);

            // Get reference to the table

            const table = document.querySelector('table');
            if (!table) throw new Error('The table element was not found');
            table.innerHTML = "<tr class =\"fileTitle\"><th>File Name</th><th>Created Date</th><th>Last Modified</th></tr>"; // Reset table
            const docFragment = new DocumentFragment();
            docFragment.appendChild(table);//We move the table to the docFragment in order to avoid the slow down related to the insertion of the rows directly in the DOM 

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
            spinner(false);//We hide the spinner


            console.log(`Fetched ${files.length} items, displaying ${matchingFiles.length} matching files.`);
        }

        async function getDownloadLink(fileId: string) {
            const data = await JSONFromGETRequest(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`);

            return data.webUrl;
        }

        async function fetchAllFilesByBatches() {
            const path = (byID('folder') as HTMLInputElement)?.value;
            if (!path) throw new Error('The file path could not be retrieved');

            const allFiles: fileItem[] = [];
            const existing = await manageFilesDatabase(allFiles, path);
            if (existing.length) return existing as fileItem[];

            localStorage.folderPath = path;
            const select = '$select=name,id,folder,file,createdDateTime,lastModifiedDateTime';
            const top = '$top=900';
            await fetchAllFilesByPath(path);

            return await manageFilesDatabase(allFiles, path);

            async function fetchAllFilesByPath(path: string) {
                // Step 1: Get root-level files & folders
                path = path.replace('\\', '/');
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
                return data.value as (fileItem | folderItem)[]; // Returns an array of files & folders

                async function getFolderIdByPath(path: string): Promise<string> {
                    const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${path}`;
                    const data = await JSONFromGETRequest(endpoint);
                    return data.id; // Folder ID
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
                    const batchData = await fetchRequests(batchRequests.slice(i, i + limit));
                    await processItems(batchData);

                }

                async function fetchRequests(requests: any[]) {
                    const response = await fetch(batchUrl, {
                        method: "POST",
                        headers: {
                            Authorization: `Bearer ${accessToken}`,
                            "Content-Type": "application/json",
                        },
                        body: JSON.stringify({ requests: requests }),
                    });

                    if (!response.ok) throw new Error(`Error fetching subfolders: ${await response.text()}`);

                    return await response.json();

                }

                async function processItems(data: any) {
                    // Extract file lists from batch responses
                    const items = data.responses.map((res: any) => res.body.value).flat() as (fileItem | folderItem)[];
                    const [files, folders] = getFilesAndFolders(items);
                    allFiles.push(...files);
                    const subfolderIds = folders.map((f) => f.id);
                    await fetchSubfolderContents(subfolderIds);
                }

            }

        };

        function getFilesAndFolders(items: (fileItem | folderItem)[]): [fileItem[], folderItem[]] {
            return [getFiles(items), subFolders(items)];
        }
        function subFolders(items: (fileItem | folderItem)[]) {
            return items.filter(item => (item as folderItem)?.folder) as folderItem[];
        }
        function getFiles(items: (fileItem | folderItem)[]) {
            return items.filter(item => (item as fileItem)?.file) as fileItem[];
        }
        async function JSONFromGETRequest(url: string) {
            const response = await fetch(url, {
                method: "GET",
                headers: { Authorization: `Bearer ${accessToken}` },
            });
            if (!response.ok) throw new Error(`Error fetching items from endpoint ${url}: \n${await response.text()}`);
            return await response.json();
        }
        function filterFiles(files: fileItem[], search: string) {
            const byName = files.filter((item: any) => RegExp(search, 'i').test(item.name));
            const created = (file: fileItem) => new Date(file.createdDateTime);

            const after = (form.querySelector('#after') as HTMLInputElement)?.valueAsDate;
            const before = (form.querySelector('#before') as HTMLInputElement)?.valueAsDate;

            if (after && before)
                return byName.filter(file => created(file).getTime() > after.getTime() && created(file).getTime() < before.getTime());
            else if (before)
                return byName.filter(file => created(file).getTime() < before.getTime());
            else if (after)
                return byName.filter(file => created(file).getTime() > after.getTime());
            else return byName

        }


        async function manageFilesDatabase(files: fileItem[], path: string, deleteRecord: boolean = false): Promise<fileItem[]> {
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
                    reject("Failed to open database: " + (event.target as IDBOpenDBRequest)?.error);
                };
            });

            // Retrieve or add the entry
            return new Promise((resolve, reject) => {
                const transaction = (db as IDBDatabase).transaction(storeName, "readwrite");
                const store = transaction.objectStore(storeName);

                // Check if an entry with the given path exists
                const getRequest = store.get(path);

                getRequest.onsuccess = function (event) {
                    const existingEntry = (event.target as IDBRequest)?.result;
                    if (existingEntry && deleteRecord) {
                        const deleteRequest = store.delete(path);
                        deleteRequest.onsuccess = function () {
                            console.log('successfuly deleted the record');
                            resolve(files);
                        }
                        deleteRequest.onerror = function () {
                            reject("Failed to delete the specified record: " + (event.target as IDBRequest)?.error);
                        }
                    } else if (existingEntry) {
                        console.log("Entry found for path:", path);
                        resolve(existingEntry.files as fileItem[]); // Return the existing files array
                    } else if (!files.length) {
                        resolve(files);//We return the empty array
                    } else {
                        // Add a new entry if it doesn't exist
                        const data = { path: path, files: files };
                        const addRequest = store.put(data);

                        addRequest.onsuccess = function () {
                            console.log("New entry added for path:", path);
                            resolve(files); // Return the newly added files array
                        };

                        addRequest.onerror = function (event) {
                            reject("Failed to add new entry: " + (event.target as IDBRequest)?.error);
                        };
                    }
                };

                getRequest.onerror = function (event) {
                    reject("Failed to retrieve entry: " + (event.target as IDBRequest)?.error);
                };
            });
        }

    }
}




