
function getAccessToken() {
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
    return getTokenWithMSAL(clientId, redirectUri, msalConfig)
}

async function setLocalStorageTitles(sessionId?: string) {
    if (!accessToken)
        accessToken = await getAccessToken() || '';
    if (!accessToken) return [];
    if (!sessionId)
        sessionId = await createFileCession(workbookPath, accessToken);
    if (!sessionId) return [];

    TableRows = await fetchExcelTableWithGraphAPI(sessionId, accessToken, workbookPath, tableName, true) as string[][];

    tableTitles = TableRows?.[0];
    if (!tableTitles) return [];
    localStorage.setItem('tableTitles', JSON.stringify(tableTitles));
    await closeFileSession(sessionId, workbookPath, accessToken);
    return tableTitles;
}
/**
 * 
 * @param {boolean} add - If false, the function will only show a form containing input fields for the user to provide the data for the new row to be added to the Excel Table. If true, the function will parse the values from the input fields in the form, and will add them as a new row to the Excel Table. Its default value is false.
 * @param {any[]} row - If provided, the function will add the row directly to the Excel Table without needing to retrieve the data from the inputs.
 */
async function addNewEntry(add: boolean = false, row?: any[]) {
    accessToken = await getAccessToken() || '';
    if (!accessToken) return alert('The access token is missing. Check the console.log for more details');


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
            const sessionId = await createFileCession(workbookPath, accessToken);
            if (!sessionId) throw new Error('There was an issue with the creation of the file cession. Check the console.log for more details');
            if (!workbookPath || !tableName) throw new Error('The Excel Workbook path and/or the name of the Excel Table are missing or invalid');
            if (!tableTitles || !TableRows) tableTitles = await setLocalStorageTitles(sessionId);

            insertAddForm(tableTitles);
            await closeFileSession(sessionId, workbookPath, accessToken);
            spinner(false);//We hide the spinner


            function insertAddForm(titles: string[]) {
                if (!titles) throw new Error('The table titles are missing. Check the console.log for more details');


                const form = document.getElementById('form');
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
                            if ([9, 10, 14, 16].includes(index)) return;//We exclude the "Montant" (9), "TVA" (10), "Description" (14), and the "Link to file" (16) columns;
                            else if (index > 2 && index < 8) return; //We exclude the "Date" (3), "Année" (4), "Start Time" (5), "End Time" (6), "Total Time" (7) columns

                            input.setAttribute('list', input.id + 's');

                            if ([1, 8, 15].includes(index)) return;
                            createDataList(input.id, getUniqueValues(index, TableRows.slice(1, -1)));//We don't create the data list for columns 'Matter' (1), "Hourly Rate" (8) and 'Adress' (15) because the data list will be created when the 'Client' input (0) is updated

                            if (index > 1) return;//We add onChange for "Client" (0) and "Affaire" (1) columns only.

                            input.onchange = () => inputOnChange(index, TableRows.slice(1, -1), false);
                        })();

                        (function addRestOnChange() {
                            if (index < 5 || index > 10) return;
                            //Only for the  "Start Time", "End Time", "Total Time", "Hourly Rate", "Amount", "VAT" columns . The "Total Time" input (7) is hidden, so it can't be changed by the user. We will add the onChange event to it by simplicity

                            input.onchange = () => inputOnChange(index, undefined, false);//!We are passing the table[][] argument as undefined, and the invoice argument as false which means that the function will only reset the bound inputs without updating any data list
                        })();
                    })();

                    return input
                }
            }
        }


    })();

    (async function addEntry() {
        if (!add) return;
        spinner(true);//We show the spinner
        try {
            if (row) await addRow(row);//If a row is already passed, we will add them directly
            else await addRow(parseInputs() || undefined, true);
        } catch (error) {
            spinner(false);//We hide the spinner
            alert(error)
        }



        function parseInputs() {
            const stop = (missing: string) => alert(`${missing} missing. You must at least provide the client, matter, nature, date and the amount. If you provided a time start, you must provide the end time and the hourly rate. Please review your iputs`);

            const inputs = Array.from(document.getElementsByTagName('input')) as HTMLInputElement[];//all inputs

            const nature = getInputByIndex(inputs, 2)?.value;
            if (!nature) return stop('The matter is')
            const date = getInputByIndex(inputs, 3)?.valueAsDate as Date | undefined;
            if (!date) return stop('The invoice date is');
            const amount = getInputByIndex(inputs, 9) as HTMLInputElement;
            const rate = getInputByIndex(inputs, 8)?.valueAsNumber || 0;

            const debit = ['Honoraire', 'Débours/Dépens', 'Débours/Dépens non facturables', 'Rétrocession d\'honoraires', 'Charges déductibles'].includes(nature);//We check if we need to change the value sign

            const row =
                inputs.map((input, index) => getInputValue(index));//!CAUTION: The html inputs are not arranged according to their dataset.index values. If we follow their order, some values will be assigned to the wrong column of the Excel table. That's why we do not pass the input itself or the dataset.index of the input to getInputValue(), but instead we pass the index of the column for which we want to retrieve the value from the relevant input.

            if (missing()) return stop('Some of the required fields are');

            return row

            function getInputValue(index: number) {
                const input = getInputByIndex(inputs, index) as HTMLInputElement;
                if ([3, 4].includes(index))
                    return getISODate(date);//Those are the 2 date columns
                else if ([5, 6].includes(index))
                    return getTime([input]);//time start and time end columns
                else if (index === 7) {
                    //!This is a hidden input
                    const totalTime = getTime([getInputByIndex(inputs, 5), getInputByIndex(inputs, 6)]);//Total time column
                    if (totalTime && rate && !amount.valueAsNumber)
                        amount.valueAsNumber = totalTime * 24 * rate// making the amount equal the rate * totalTime
                    return totalTime
                }
                else if (debit && index === 9)
                    return input.valueAsNumber * -1 || 0;//This is the amount if negative
                else if ([8, 9, 10].includes(index))
                    return input.valueAsNumber || 0;//Hourly Rate, Amount, VAT
                else return input.value;

            }

            function missing() {
                if (row.filter((value, i) => (i < 4 || i === 9) && !value).length > 0) return true;//if client name, matter, nature, date or amount are missing
                //else if (row[9]) return [5, 6,7,8].map(index => row[index] = 0).length < 1;//This means the amount has been provided and does not  depend on the time spent or the hourly rate. We set the values of the startTime and endTime to 0, and return false (length<1 must return false)

                if (row[5] === row[6]) return false;//If the total time = 0 we do not need to alert if the hourly rate is missing
                else if (row[5] && (!row[6] || !row[8]))
                    return true//if startTime is provided but without endTime or without hourly rate
                else if (row[6] && (!row[5] || !row[8]))
                    return true//if endTime is provided but without startTime or without hourly rate
            };
        }

        async function addRow(row: any[] | undefined, filter: boolean = false) {
            if (!row) throw new Error('The row is not valid');
            const visibleCells = await addRowToExcelTableWithGraphAPI(row, TableRows.length - 2, workbookPath, tableName, accessToken, filter);
            if (!visibleCells)
                return alert('There was an issue with the adding or the filtering, check the console.log for more details');

            alert('Row aded and the table was filtered');

            displayVisibleCells(visibleCells);
            spinner(false);//We hide the spinner

            function displayVisibleCells(visibleCells: string[][]) {
                const tableDiv = createDivContainer();
                const table = document.createElement('table');
                table.classList.add('table');
                tableDiv.appendChild(table);

                const columns = [0, 1, 2, 7, 8, 9, 10, 14];//The columns that will be displayed in the table;
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


                const form = document.getElementById('form');
                if (!form) throw new Error('The form element was not found');
                if (form) {
                    form?.insertAdjacentElement('afterend', tableDiv);
                }

                function createDivContainer() {
                    const id = 'retrieved';
                    let tableDiv = document.getElementById(id)
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

        };

    })();

};

// Update Word Document
async function invoice(issue: boolean = false) {
    accessToken = await getAccessToken() || '';
    if (!accessToken) return alert('The access token is missing. Check the console.log for more details');

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
            const sessionId = await createFileCession(workbookPath, accessToken) || '';
            if (!sessionId) throw new Error('There was an issue with the creation of the file cession. Check the console.log for more details');
            if (!tableTitles || !TableRows) tableTitles = await setLocalStorageTitles(sessionId);
            
            insertInvoiceForm(tableTitles);
            await closeFileSession(sessionId, workbookPath, accessToken);
            spinner(false);//We hide the spinner
        }
        
        function insertInvoiceForm(tableTitles: string[]) {
            if (!tableTitles) throw new Error('The table titles are missing. Check the console.log for more details');
            const form = document.getElementById('form');
            if (!form) throw new Error('The form element was not found');
            form.innerHTML = '';

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

            function insertInputsAndLables(indexes: (number | string)[], id: string, checkBox: boolean = false): HTMLInputElement[] {
                let css = 'field';
                if (checkBox) css = 'checkBox';
                return indexes.map((index) => {
                    const div = newDiv(String(index));
                    appendLable(index, div);
                    return appendInput(index, div);
                });

                function appendInput(index: number | string, div: HTMLDivElement) {
                    const input = document.createElement('input');
                    input.classList.add(css);
                    !isNaN(Number(index)) ? input.id = id + index.toString() : input.id = id;

                    (function setType() {
                        if (checkBox) input.type = 'checkbox';
                        else if (isNaN(Number(index)) || Number(index) < 3) input.type = 'text';
                        else input.type = 'date';
                    })();

                    (function notCheckBox() {
                        if (isNaN(Number(index)) || checkBox) return;//If the index is not a number or the input is a checkBox, we return;
                        index = Number(index);
                        input.name = input.id;
                        input.dataset.index = index.toString();
                        input.setAttribute('list', input.id + 's');
                        input.autocomplete = "on";

                        if (index < 2)
                            input.onchange = () => inputOnChange(Number(input.dataset.index), TableRows.slice(1, -1), true);

                        if (index < 1)
                            createDataList(input.id, getUniqueValues(0, TableRows.slice(1, -1)));//We create a unique values dataList for the 'Client' input
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

        const sessionId = await createFileCession(workbookPath, accessToken, true) || '';//!persist must be = true because we might add a new row if there is a discount. If we don't persist the session, the table will be filtered and the new row will not be added.
        if (!sessionId) throw new Error('There was an issue with the creation of the file cession. Check the console.log for more details');

        const inputs = Array.from(document.getElementsByTagName('input'));
        const criteria = inputs.filter(input => Number(input.dataset.index) >= 0);

        const discount = parseInt(inputs.find(input => input.id === 'discount')?.value || '0%');

        const lang = inputs.find(input => input.dataset.language && input.checked === true)?.dataset.language || 'FR';

        const data = await filterExcelData(criteria, discount, lang);

        if (!data) throw new Error('Could not retrieve the filtered Excel table');

        const [wordRows, totalsLabels, matters, adress] = data;

        const date = new Date();
        const invoice = {
            number: getInvoiceNumber(date),
            clientName: getInputValue(0, criteria),
            matters: getArray(getInputValue(1, criteria)),
            adress: adress,
            lang: lang
        }

        const contentControls = getContentControlsValues(invoice, date);

        const fileName = getInvoiceFileName(invoice.clientName, invoice.matters, invoice.number);
        let filePath = `${destinationFolder}/${fileName}`;

        filePath = prompt(`The file will be saved in ${destinationFolder}, and will be named : ${fileName}./nIf you want to change the path or the name, provide the full file path and name of your choice without any sepcial characters`, filePath) || filePath;

        (async function editInvoiceFilterExcelClose() {
            await createAndUploadXmlDocument(accessToken, templatePath, filePath, lang, 'Invoice', wordRows, contentControls, totalsLabels);

            await filterExcelTableWithGraphAPI(workbookPath, tableName, matter, matters, sessionId, accessToken);//We filter the table by the matters that were invoiced

            await closeFileSession(sessionId, workbookPath, accessToken);
            spinner(false);//We hide the spinner
        })();

        /**
         * Filters the Excel table according to the values of each inputs, then returns the values of the Word table rows that will be added to the Word table in the invoice template document 
         * @param {any[][]} data - The Excel table rows that will be filtered
         * @param {HTMLInputElement[]} criteria - the html inputs containing the values based on which the table will be filtered
         * @param {string} lang - The language in which the invoice will be issued 
         * @returns {string[][]} - The values of the rows that will be added to the Word table in the invoice template
         */
        async function filterExcelData(criteria: HTMLInputElement[], discount: number, lang: string): Promise<[string[][], string[], string[], string[]] | void> {

            await clearFilterExcelTableGraphAPI(workbookPath, tableName, sessionId, accessToken);//We unfilter the table;

            //Filtering by Client (criteria[0])
            await filterExcelTableWithGraphAPI(workbookPath, tableName, client, [criteria[0].value], sessionId, accessToken);

            let visible = await getVisibleCellsWithGraphAPI(workbookPath, tableName, sessionId, accessToken) as any[][];

            if (!visible) {
                return alert('Could not retrieve the visible cells of the Excel table');
            }
            visible = visible.slice(1, - 1);//We exclude the first and the last rows of the table. The first row is the header, and the last row is the total row.

            const adress = getUniqueValues(15, visible);//!We must retrieve the adresses at this stage before filtering by "Matter" or any other column

            const [matters, natures] = [1, 2].map(index => {
                //!Matter and Nature inputs (from columns 2 & 3 of the Excel table) may include multiple entries separated by ', ' not only one entry.
                const list = criteria[index].value.split(',').map(el => el.trimStart().trimEnd());//We generate a string[] from the input.value
                visible = visible.filter(row => list.includes(row[index]));
                return list;
            });
            //We finaly filter by date
            visible = filterByDate(visible);

            return [...getRowsData(visible, discount, lang), matters, adress];

            function filterByDate(visible: string[][]) {

                const convertDate = (date: string | number) => dateFromExcel(Number(date)).getTime();

                const [from, to] = criteria
                    .filter(input => getIndex(input) === 3)
                    .map(input => input.valueAsDate?.getTime());

                if (from && to)
                    return visible.filter(row => convertDate(row[3]) >= from && convertDate(row[3]) <= to); //we filter by the date
                else if (from)
                    return visible.filter(row => convertDate(row[3]) >= from); //we filter by the date
                else if (to)
                    return visible.filter(row => convertDate(row[3]) <= to); //we filter by the date
                else
                    return visible.filter(row => convertDate(row[3]) <= new Date().getTime()); //we filter by the date

            }

        }
    }

}

async function issueLetter(create: boolean = false) {
    accessToken = await getAccessToken() || '';
    const templatePath = '';
    (function showForm() {
        if (create) return;
        document.querySelector('table')?.remove();
        const form = document.getElementById('form');
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
    })();

    (async function generate() {
        if (!create) return;
        const input = document.getElementById('textInput') as HTMLTextAreaElement;
        if (!input) return;
        const templatePath = "Legal/Mon Cabinet d'Avocat/Administratif/Modèles Actes/Template_Lettre With Letter Head [DO NOT MODIFY].docx";
        const fileName = prompt('Provide the file name without special characthers');
        if (!fileName) return;
        const filePath = `${prompt('Provide the destination folder', "Legal/Mon Cabinet d'Avocat/Clients")}/${fileName}.docx`;
        if (!filePath) return;

        const contentControls = [['RTCoreText', input.value], ['RTReference', 'Référence'], ['RTClientName', 'Nom du Client'], ['RTEmail', 'Email du client']];

        createAndUploadXmlDocument(accessToken, templatePath, filePath, 'FR', undefined, undefined, contentControls);
    })();

}


/**
 * Updates the data list or the value of bound inputs according to the value of the input that has been changed
 * @param {number} index - the dataset.index of the input that has been changed
 * @param {any[][]} table - The table that will be filtered to update the data list of the button. If undefined, it means that the data list will not be updated.
 * @param {boolean} invoice - If true, it means that we called the function in order to generate an invoice. If false, we called it in order to add a new entry in the table
 * @returns 
 */
function inputOnChange(index: number, table: any[][] | undefined, invoice: boolean) {
    let inputs = Array.from(document.getElementsByTagName('input'))
        .filter(input => input.dataset.index) as HTMLInputElement[];

    (function resetInputs() {
        //In some cases, we only need to rest the values of other inputs bound to the input that has been changed. If the function is called for this purpose, we will just rest those inputs without updating their data list.
        if (table || invoice) return;
        const boundInputs = [5, 6, 7, 9, 10];//Those are "Start Time" (5), "End Time" (6), "Total Time" (7, although it is hidden), "Amount" (9), "VAT" (10) columns. We exclude the "Hourly Rate" column (8). We let the user rest it if he wants
        boundInputs
            .forEach(i => i > index ? reset(i) : i = i);//We reset any input which dataset-index is > than the dataset-index of the input that has been changed

        if (index === 9)
            boundInputs
                .forEach(i => i < index ? reset(i) : i = i);//If the input is the input for the "Montant" column of the Excel table, we also reset the "Start Time" (5), "End Time" (6) and "Hourly Rate" (7) columns' inputs. We do this because we assume that if the user provided the amount, it means that either this is not a fee, or the fee is not hourly billed.

        function reset(i: number) {
            const input = getInputByIndex(inputs, i);
            if (!input) return;
            input.value = '';
            if (input.valueAsNumber) input.valueAsNumber = 0;
        }
    })();

    if (!table) return;

    if (invoice)
        inputs = inputs.filter(input => getIndex(input) < 3); //Those are all the inputs that serve to filter the table (first 3 columns only) when we are invoicing the client
    else
        inputs = inputs.filter(input => [0, 1, 8, 15].includes(getIndex(input))); //Those are all the inputs that have data lists associated with them that need to be updated if an input calls inputOnChage(). Only the "Client" and "Affaire" inputs call this function in the context of adding a new entry, so index will always be <3

    const filledInputs =
        inputs
            .filter(input => input.value && getIndex(input) <= index)//Those are all the inputs that the user filled with data


    const boundInputs = inputs.filter(input => getIndex(input) > index);//Those are the inputs for which we want to create  or update their data lists


    if (filledInputs.length < 1 || boundInputs.length < 1) return;

    boundInputs.forEach(input => input.value = '');

    const filtered = filterOnInput(filledInputs, table);//We filter the table based on the filled inputs

    if (filtered.length < 1) return;

    boundInputs.map(input => {
        const dataList = createDataList(input?.id, getUniqueValues(getIndex(input), filtered), invoice) as HTMLDataListElement;
        if (dataList.options.length === 1)
            input.value = dataList.options[0].value
    });


    function filterOnInput(filled: HTMLInputElement[], table: any[][]) {
        filled
            .forEach(input => table = table.filter(row => row[getIndex(input)].toString() === input.value));
        return table
    }
};

/**
 * Creates an invoice Word document from the invoice Word template, then uploads it to the destination folder
 * @param {string} accessToken - The access token that will be used to authenticate the user
 * @param {string} templatePath - The full path of the Word invoice template
 * @param {string} filePath - The full path of the destination folder where the new invoice will be saved
 * @param {string} lang - The language in which the invoice will be issued
 * @param {string} tableTitle - The title of the table in the Word document that will be updated
 * @param {string[][]} rows - The rows that will be added to the table in the Word document
 * @param {string[][]} contentControls - The titles and text of each of the content controls that will be updated in the Word document
 * @param {string[]} totalsLabels - The labels of the rows that will be formatted as totals
 * @returns 
 */
async function createAndUploadXmlDocument(accessToken: string, templatePath: string, filePath: string, lang: string, tableTitle?: string, rows?: string[][] | undefined, contentControls?: string[][] | undefined, totalsLabels?: string[]) {

    if (!accessToken || !templatePath || !filePath) return;

    const blob = await fetchFileFromOneDriveWithGraphAPI(accessToken, templatePath);

    if (!blob) return;

    const [doc, zip] = await convertBlobIntoXML(blob);

    if (!doc) return;
    const schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    (function editTable() {
        if (!rows) return;
        const tables = getXMLElements(doc, "tbl") as Element[];
        const table = getXMLTableByTitle(tables, tableTitle);

        if (!table) return;
        const firstRow = getXMLElements(table, 'tr', 1) as Element;

        rows.forEach((row, index) => {
            const newXmlRow = insertRowToXMLTable(NaN, true) as Element || table.appendChild(createXMLElement('tr'));
            if (!newXmlRow) return;
            const isTotal = totalsLabels?.includes(row[0]);
            const isLast = index === rows.length - 1;
            return editCells(newXmlRow, row, isLast, isTotal);
        });

        firstRow.remove();//We remove the first row when we finish

        function editCells(tableRow: Element, values: string[], isLast: boolean = false, isTotal: boolean = false) {
            const cells = getXMLElements(tableRow, 'tc') as Element[] || values.map(v => tableRow.appendChild(createXMLElement('tc')));//getting all the cells in the row element

            cells.forEach((cell, index) => {
                const textElement = getXMLElements(cell, 't', 0) as Element || appendParagraph(cell);
                if (!textElement) return console.log('No text element was found !');
                const pPr = setTextLanguage(cell);//We call this here in order to set the language for all the cells. It returns the pPr element if any.
                textElement.textContent = values[index];

                (function totalsRowsFormatting() {
                    if (!isLast && !isTotal) return;
                    (function cellBackgroundColor() {
                        const tcPr = getXMLElements(cell, 'tcPr', 0) as Element || cell.prepend(createXMLElement('tcPr'));
                        const background = getXMLElements(tcPr, 'shd', 0) as Element || tcPr.appendChild(createXMLElement('shd') as Element);//Adding background color to cell
                        background.setAttributeNS(schema, 'val', "clear");
                        background.setAttributeNS(schema, 'fill', 'D9D9D9');
                    })();

                    (function paragraphStyle() {
                        if (!pPr) return console.log('No "w:pPr" or "w:rPr" property element was found !');
                        const style = getXMLElements(pPr, 'pStyle', 0) as Element || pPr.appendChild(createXMLElement('pStyle'));
                        style.setAttributeNS(schema, 'val', getStyle(index, isTotal && !isLast));
                    })();
                })();
            })
        }


        function insertRowToXMLTable(after: number = -1, clone: boolean = false) {
            if (clone) return cloneFirstRow();
            else return create();

            function create() {
                if (!table) return;
                const row = createXMLElement("tr");
                after >= 0 ? (getXMLElements(table, 'tr', after) as Element)?.insertAdjacentElement('afterend', row) :
                    table.appendChild(row);
                return row;
            }

            function cloneFirstRow() {
                const row = firstRow.cloneNode(true) as Element;
                table?.appendChild(row);
                return row
            };
        }

        function getStyle(cell: number, isTotal: boolean = false) {
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

        function getXMLTableByTitle(tables: Element[], title?: string, property: string = 'tblCaption') {
            if (!title) return;
            return tables
                .filter(table => tblCaption(table))
                .find(table => tblCaption(table).getAttribute('w:val') === title) as Element;

            function tblCaption(table: Element) {
                return getXMLElements(table, property, 0) as Element
            }
        }

    })();

    (function editContentControls() {
        if (!contentControls) return;
        const ctrls = getXMLElements(doc, "sdt") as Element[];
        contentControls
            .forEach(([title, text]) => {
                const control = findXMLContentControlByTitle(ctrls, title);
                if (!control) return;
                editXMLContentControl(control, text);
            });

        function findXMLContentControlByTitle(ctrls: Element[], title: string) {
            return ctrls.find(control => (getXMLElements(control, "alias", 0) as Element)?.getAttributeNS(schema, 'val') === title);
        }

        function editXMLContentControl(control: Element, text: string) {
            if (!text) return control.remove();
            const sdtContent = getXMLElements(control, "sdtContent", 0) as Element;
            if (!sdtContent) return;
            const paragTemplate = getParagraphOrRun(sdtContent) as Element;//This will set the language for the paragraph or the run
            if (!paragTemplate) return console.log('No template paragraph or run were found !');
            setTextLanguage(paragTemplate);//We amend the language element to the "w:pPr" or "r:pPr" child elements of paragTemplate

            text.split('\n')
                .forEach((parag, index) => editParagraph(parag, index));

            function editParagraph(parag: string, index: number) {
                let textElement: Element;
                if (index < 1)
                    textElement = getXMLElements(paragTemplate, 't', index) as Element;
                else textElement = appendParagraph(paragTemplate, sdtContent);//We pass sdtContent as parent argument

                if (!textElement) return console.log('No textElement was found !');

                textElement.textContent = parag;

            }
        }

    })();

    await convertXMLToBlobAndUpload(doc, zip, filePath, accessToken);

    /**
     * Adds a new paragraph XML element or appends a cloned paragraph, and in both cases, it returns the textElement of the paragraph
     * @param {Element} element - The element to which the new paragraph will be appended if the parent argument is not provided. If the parent argument is provided, the element will be cloned assuming that this is a pargraph element
     * @param {Elemenet} parent - If provided, element will be cloned and appended to parent.
     * @returns {Element} the textElemenet attached to the paragraph
     */
    function appendParagraph(element: Element, parent?: Element) {
        if (parent) return clone();
        else return create();
        function clone() {
            const parag = element?.cloneNode(true) as Element;
            parent?.appendChild(parag);
            return getXMLElements(parag, 't', 0) as Element
        }
        function create() {
            const parag = element.appendChild(createXMLElement('p'));
            parag.appendChild(createXMLElement('pPr'));
            const run = parag.appendChild(createXMLElement('r'));
            return run.appendChild(createXMLElement('t'));
        }
    }

    function createXMLElement(tag: string, parent?: HTMLElement) {
        return doc.createElementNS(schema, tag);
    }

    function getXMLElements(xmlDoc: XMLDocument | Element, tag: string, index: number = NaN): Element[] | Element {
        const elements = xmlDoc.getElementsByTagNameNS(schema, tag);
        if (!isNaN(index)) return elements?.[index];
        return Array.from(elements)
    }

    /**
     * Looks for a child "w:p" (paragraph) element, if it doesn't find any, it looks for a "w:r" (run) element.
     * @param {Element} parent - the parent XML of the paragraph or run element we want to retrieve. 
     * @returns {Element | undefined} - an XML element representing a "w:p" (paragraph) or, if not found, a "w:r" (run), or undefined
     */
    function getParagraphOrRun(parent: Element) {
        return getXMLElements(parent, 'p', 0) as Element || getXMLElements(parent, 'r', 0) as Element;
    }
    /**
     * Finds a "w:pPr" XML element (property element) which is a child of the XML parent element passed as argument. If does not find it, it looks for a "w:rPr" XML element. When it finds either a "w:pPr" or a "w:rPr" element, it appends a "w:lang" element to it, and sets its "w:val" attribute to the language passed as "lang"
     * @param {Element} parent - the XML element containing the paragraph or the run for which we want to set the language.
     * @returns {Element | undefined} - the "w:pPr" or "w:rPr" property XML element child of the parent element passed as argument
     */
    function setTextLanguage(parent: Element) {
        const pPr = getXMLElements(parent, 'pPr', 0) as Element ||
            getXMLElements(parent, 'rPr', 0) as Element;
        if (!pPr) return;
        pPr
            .appendChild(createXMLElement('lang'))//appending a "w:lang" element
            .setAttributeNS(schema, 'val', `${lang.toLowerCase()}-${lang.toUpperCase()}`);//setting the "w:val" attribute of "w:lang" to the appropriate language like "fr-FR"
        return pPr as Element
    }
};

/**
 * Converts the blob of a Word document into an XML
 * @param blob - the blob of the file to be converted
 * @returns {[XMLDocument, JSZip]} - The xml document, and the zip containing all the xml files
 */
//@ts-expect-error
async function convertBlobIntoXML(blob: Blob): Promise<[XMLDocument, JSZip]> {
    //@ts-ignore
    const zip = new JSZip();

    const arrayBuffer = await blob.arrayBuffer();

    await zip.loadAsync(arrayBuffer);

    const documentXml = await zip.file("word/document.xml").async("string");

    const parser = new DOMParser();

    const xmlDoc = parser.parseFromString(documentXml, "application/xml");

    return [xmlDoc, zip]
}

/**
 * Converts an XML Word document into a Blob, and uploads it to OneDrive using the Graph API
 * @param {XMLDocument} doc 
 * @param {JSZip} zip 
 * @param {string} filePath - the full OneDrive file path (including file name and extension) of the file that will be uploaded
 * @param {string} accessToken - the Graph API accessToken
 */
//@ts-expect-error
async function convertXMLToBlobAndUpload(doc: XMLDocument, zip: JSZip, filePath: string, accessToken: string) {
    const blob = await convertXMLIntoBlob();
    if (!blob) return;

    await uploadFileToOneDriveWithGraphAPI(blob, filePath, accessToken);

    async function convertXMLIntoBlob() {

        const serializer = new XMLSerializer();
        let modifiedDocumentXml = serializer.serializeToString(doc);

        zip.file("word/document.xml", modifiedDocumentXml);

        return await zip.generateAsync({ type: "blob" });
    }
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
    return inputs.map(input => {
        input.value

    })

}
/**
 * Adds a new row to the Excel table using the Grap API
 * @param {string} row - The row that will be added to the Excel table
 * @param {number} index - The index at which the row will be added
 * @param {string} filePath - The full path of the Excel file 
 * @param {string} tableName - The name of the Excel table
 * @param {string} accessToken - The Graph API access token 
 * @param {boolean} filter - If true, the table will be filtered after the row is added
 * @returns 
 */
async function addRowToExcelTableWithGraphAPI(row: any[], index: number, filePath: string, tableName: string, accessToken: string, filter: boolean = false) {
    const sessionId = await createFileCession(filePath, accessToken, true);//!persist must be = true because 
    if (!sessionId) return alert('There was an issue with the creation of the file cession. Check the console.log for more details');
    await clearFilterExcelTableGraphAPI(filePath, tableName, sessionId, accessToken);
    await addRow();
    if (filter) await filterTable();
    await sortExcelTableWithGraphAPI(filePath, tableName, [[3, true]], false, sessionId, accessToken);//We sort the table by the first column (the date column)
    const visible = await getVisibleCellsWithGraphAPI(filePath, tableName, sessionId, accessToken);
    await closeFileSession(sessionId, filePath, accessToken);
    return visible;

    async function addRow() {
        const url = `${GRAPH_API_BASE_URL}${filePath}:/workbook/tables/${tableName}/rows`;//The url to add a row to the table

        const body = {
            index: index,
            values: [row], 
        };
        const resp = await POSTRequestWithGraphAPI(url, accessToken, sessionId, JSON.stringify(body), "Error adding row", filePath);
        if(resp) console.log("Row added successfully!");

    }

    async function filterTable() {
        if (!filter) return;
        [0, 1].map(async index => {
            //!We use map because forEach doesn't await
            await filterExcelTableWithGraphAPI(workbookPath, tableName, tableTitles?.[index], [row[index]?.toString()], sessionId, accessToken);
        });
    };

}

function searchFiles() {
    (function showForm() {
        const form = document.getElementById('form') as HTMLDivElement;
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
        type folderItem = { name: string; id: string; folder: any; createdDateTime: string; lastModifiedDateTime: string };
        type fileItem = { name: string; id: string; file: any; createdDateTime: string; lastModifiedDateTime: string; "@microsoft.graph.downloadUrl": string };

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
            const path = (document.getElementById('folder') as HTMLInputElement)?.value;
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

        
        async function manageFilesDatabase(files: fileItem[], path: string, deleteRecord:boolean =false):Promise<fileItem[]> {
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
                        const data = {path:path, files:files };
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




