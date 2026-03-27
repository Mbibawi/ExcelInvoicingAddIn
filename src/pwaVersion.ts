class LawFirm {
    private stored;
    private tenantID;
    private settingsNames;

    constructor() {
        this.stored = getSavedSettings() || undefined;
        this.settingsNames = settingsNames;
        this.tenantID = "f45eef0e-ec91-44ae-b371-b160b4bbaa0c";
    }

    /**
     * 
     * @param {boolean} add - If false, the function will only show a form containing input fields for the user to provide the data for the new row to be added to the Excel Table. If true, the function will parse the values from the input fields in the form, and will add them as a new row to the Excel Table. Its default value is false.
     * @param {boolean} display - If provided, the function will show the visible rows in the UI after the new row has been added.
     * @param {any[]} row - If provided, the function will add the row directly to the Excel Table without needing to retrieve the data from the inputs.
     */
    async addNewEntry(add: boolean = false, row?: any[]) {
        spinner(true);//We show the spinner
        const { workbookPath, tableName } = this.getConsts(this.settingsNames.invoices);

        if ([this.stored, workbookPath, tableName].find(v => !v)) throwAndAlert('One of the constant values is not valid');

        const graph = new GraphAPI('', workbookPath);

        const TableRows = await graph.fetchExcelTable(tableName, true);
        if (!TableRows?.length) return alert('Failed to retrieve the Excel table')
        const tableTitles = TableRows[0];
        if (add) return addEntry(TableRows, tableTitles);
        
        await showAddNewForm(this);

        async function showAddNewForm(this$:LawFirm) {
            try {
                await createForm();
                spinner(false);//We hide the sinner
            } catch (error) {
                spinner(false);//We hide the sinner
                alert(error);
            }

            async function createForm() {
                const tableBody = TableRows!.slice(1, -1);
                const inputs: HTMLInputElement[] = [];
                const bound = (indexes: number[]) => inputs.filter(input => indexes.includes(getIndex(input))).map(input => [input, getIndex(input)]) as InputCol[];

                insertAddForm(tableTitles);

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
                        btnIssue.onclick = () => addEntry(TableRows!, tableTitles);
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

                                if (updateNext.includes(index)) input.onchange = () => this$.inputOnChange(index, bound(updateNext), tableBody, false);

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
        };

        async function addEntry(tableRows: any[][], tableTitles: string[]) {
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
                const date = getInputByIndex(inputs, colDate)?.valueAsDate as Date | null;
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
                if (!row) return throwAndAlert('The row is not valid');
                const visibleCells = await graph.addRowToExcelTable(row, tableRows!.length - 2, tableName!, tableTitles);

                if (!visibleCells?.length)
                    return throwAndAlert('There was an issue with the adding or the filtering, check the console.log for more details');

                alert('Row aded and the table was filtered');
                return visibleCells
            };
            function displayVisibleCells(visibleCells: string[][], display: boolean) {
                if (!display) return;
                const form = byID();
                const tableDiv = createDivContainer();
                if (!form) return throwAndAlert('The form element was not found');
                const table = document.createElement('table');
                table.classList.add('table');
                tableDiv.appendChild(table);

                const columns = [0, 1, 2, 3, 7, 8, 9, 10, 14];//The columns that will be displayed in the table;
                const rowClass = 'excelRow';
                (function insertTableHeader() {
                    if (!tableTitles) return throwAndAlert('No Table Titles');
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

                form.insertAdjacentElement('afterend', tableDiv);

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
        };

    };

    async issueInvoice() {
        const this$: LawFirm = this;
        
        await showInvoiceForm();

        async function showInvoiceForm() {
            spinner(true);//We show the spinner
            const { workbookPath, tableName, templatePath, saveTo } = this$.getConsts(this$.settingsNames.invoices);

            if ([this$.stored, workbookPath, tableName, templatePath, saveTo].find(v => !v)) throwAndAlert('One of the  constant values is not valid');

            const graph = new GraphAPI('', workbookPath);

            const sessionId = await graph.createFileSession() || '';
            if (!sessionId) return throwAndAlert('There was an issue with the creation of the file cession. Check the console.log for more details');

            const tableRows = await graph.fetchExcelTable(tableName, true);
            if (!tableRows?.length) return throwAndAlert('Failed to retrieve the Excel table');
            const tableTitles = tableRows[0];
            document.querySelector('table')?.remove();
            try {
                insertInvoiceForm(tableTitles);
                await graph.closeFileSession(sessionId);
                spinner(false);//We hide the spinner
            }
            catch (error) {
                throwAndAlert(`Error while showing the invoice user form: ${error}`)
                spinner(false);//We hide the spinner
            }

            function insertInvoiceForm(tableTitles: string[]) {
                const form = byID();
                if (!form) throw new Error('The form element was not found');
                const isNan = (index:number|string)=> isNaN(Number(index));
                form.innerHTML = '';
                const tableBody = tableRows!.slice(1, -1);
                const boundInputs: InputCol[] = [];

                (function insertInputs() {
                    insertInputsAndLables([0, 1, 2, 3, 3], 'input');//Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice
                    insertInputsAndLables(['Discount'], 'discount')[0].value = '0%'; //Inserting a discount percentage input and setting its default value to 0%
                    insertInputsAndLables(['Français', 'English'], 'lang', true); //Inserting languages checkboxes
                })();

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
                    btnIssue.onclick = () => createInvoice(tableName, tableTitles, templatePath!, saveTo, graph);
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
                        const input = document.createElement('input');
                        input.classList.add(css);
                        const isNaN = isNan(index);
                        !isNaN ? input.id = id + index.toString() : input.id = id;

                        (function inputType() {
                            if (checkBox) input.type = 'checkbox';
                            else if (isNaN || index as number < 3) input.type = 'text';
                            else input.type = 'date';
                        })();

                        (function notCheckBox() {
                            if (isNaN || checkBox) return;//If the index is not a number or the input is a checkBox, we return;
                            index = Number(index);
                            input.name = input.id;
                            input.dataset.index = index.toString();
                            if (index < 3)
                                boundInputs.push([input, index]);//Fields "Client"(0), "Affaire"(1), "Nature"(2) are the inputs that will need to get their dataList created or updated each time the previous input is changed.
                            if (index < 2)
                                input.onchange = () => this$.inputOnChange(index as number, boundInputs, tableBody, true);//We add onChange on "Client" (0) and "Affaire" (1) columns. We set combined = true in order to add to the dataList of the next column an option combining all the choices in the list
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
                        isNan(Number(index)) || checkBox ? label.innerText = index.toString() : label.innerText = tableTitles[Number(index)];
                        !isNan(Number(index)) ? label.htmlFor = id + index.toString() : label.htmlFor = id;
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
        };

        async function createInvoice(tableName: string, tableTitles: string[], templatePath: string, saveTo: string, graph: GraphAPI) {
            spinner(true);//We show the spinner
            try {
                await editInvoice(tableName, tableTitles, templatePath, saveTo, graph);
                spinner(false);//We hide the spinner
            } catch (error) {
                spinner(false);//We hide the sinner
                alert(error)
            }
        };

        async function editInvoice(tableName: string, tableTitles: string[], templatePath: string, saveTo: string, graph: GraphAPI) {
            const client = tableTitles[0], matter = tableTitles[1];//Those are the 'Client' and 'Matter' columns of the Excel table
            const sessionId = await graph.createFileSession(true) || '';//!persist must be = true. This means that if the session is closed, the changes made to the file will be saved.
            if (!sessionId) return throwAndAlert('There was an issue with the creation of the file cession. Check the console.log for more details');

            const inputs = Array.from(document.getElementsByTagName('input'));
            const criteria = inputs.filter(input => getIndex(input) >= 0);

            const discount = parseInt(inputs.find(input => input.id === 'discount')?.value || '0%');

            const lang = inputs.find(input => input.dataset.language && input.checked === true)?.dataset.language || 'FR';

            const date = new Date();//We need to generate the date at this level and pass it down to all the functions that need it
            const invoiceNumber = getInvoiceNumber(date);
            const data = await filterExcelData(criteria, discount, lang, invoiceNumber);

            if (!data) return throwAndAlert('Could not retrieve the filtered Excel table');

            const { wordRows, totalsLabels, clientName, matters, adresses } = data;

            const invoice = {
                number: invoiceNumber,
                clientName: clientName,
                matters: matters,
                adress: adresses,
                lang: lang
            }

            const contentControls = this$.getContentControlsValues(invoice, date);

            const fileName = getInvoiceFileName(clientName, matters, invoiceNumber);
            let saveToPath = `${saveTo}/${fileName}`;

            saveToPath = prompt(`The file will be saved in ${saveTo}, and will be named : ${fileName}.\nIf you want to change the path or the name, provide the full file path and name of your choice without any sepcial characters`, saveToPath) || saveTo;

            (async function editInvoiceFilterExcelClose() {
                await graph.createAndUploadDocumentFromTemplate(templatePath!, saveToPath, lang, [['Invoice', wordRows, 1]], contentControls, totalsLabels);
                await graph.clearFilterExcelTable(tableName!, sessionId);//We unfilter the table;
                await graph.filterExcelTable(tableName, client, [clientName], sessionId);//We filter the table by the matters that were invoiced
                await graph.filterExcelTable(tableName, matter, matters, sessionId);//We filter the table by the matters that were invoiced
                await graph.closeFileSession(sessionId);
            })();

            /**
             * Filters the Excel table according to the values of each inputs, then returns the values of the Word table rows that will be added to the Word table in the invoice template document 
             * @param {HTMLInputElement[]} inputs - the html inputs containing the values based on which the table will be filtered
             * @param {number} discount  - The discount percentage that will be applied to the amount of each invoiced row if any. It is a number between 0 and 100. If it is equal to 0, it means that no discount will be applied.
             * @param {string} lang - The language in which the invoice will be issued 
             * @returns {Promise<[string[][], string[], string[], string[]]>} - The values of the rows that will be added to the Word table in the invoice template
             */
            async function filterExcelData(inputs: HTMLInputElement[], discount: number, lang: string, invoiceNumber: string): Promise<{ wordRows: string[][], totalsLabels: string[], clientName: string, matters: string[], adresses: string[] } | void> {
                const clientCol = 0, matterCol = 1, dateCol = 3, addressCol = 15;//Indexes of the 'Matter' and 'Date' columns in the Excel table
                const clientNameInput = getInputByIndex(inputs, clientCol);
                const matterInput = getInputByIndex(inputs, matterCol);
                const clientName = clientNameInput!.value || '';
                const matters =
                    getArray(matterInput!.value) || []; //!The Matter input may include multiple entries separated by ', ' not only one entry.

                if (!clientName || !matters?.length) throwAndAlert('could not retrieve the client name or the matter/matters list from the inputs');

                const excelTable = await graph.fetchExcelTable(tableName, true);

                let tableRows = excelTable?.slice(1, -1) || undefined; //We exclude the first and the last rows of the table. Since we are calling the "range" endpoint, we get the whole table including the headers. The first row is the header, and the last row is the total row.

                if (!tableRows) return throwAndAlert('We could not retrieve the tableRows whie trying to issue the invoice');

                //tableRows = _filterTableByInputsValues([[clientNameInput!, clientCol], [matterInput!, matterCol]], excelTable!);
                tableRows = this$.filterTableByInputsValues([[clientNameInput!, clientCol], [matterInput!, matterCol]], excelTable!);
                tableRows = filterByDate(tableRows!, dateCol);

                const adresses = getUniqueValues(addressCol, tableRows) as string[];//!We must retrieve the adresses at this stage before filtering by "Matter" or any other column

                //const {wordRows, totalsLabels} = _getRowsData(tableRows, discount, lang, invoiceNumber);
                const { wordRows, totalsLabels } = this$.getRowsData(tableRows, discount, lang, invoiceNumber);

                return { wordRows, totalsLabels, clientName, matters, adresses };

                function filterByDate(visible: string[][], dateCol: number) {

                    const convert = (date: string | number) => dateFromExcel(Number(date)).getTime();

                    const [from, to] = inputs
                        .filter(input => getIndex(input) === dateCol)
                        .map(input => input.valueAsDate?.getTime());

                    if (from && to)
                        return visible.filter(row => convert(row[dateCol]) >= from && convert(row[dateCol]) <= to); //we filter by the date
                    else if (from)
                        return visible.filter(row => convert(row[dateCol]) >= from); //we filter by the date
                    else if (to)
                        return visible.filter(row => convert(row[dateCol]) <= to); //we filter by the date
                    else
                        return visible.filter(row => convert(row[dateCol]) <= new Date().getTime()); //we filter by the date
                }

            }
        }
    }

    async issueLetter() {
        showForm(this);

        function showForm(this$:LawFirm) {
            spinner(true);//We show the spinner
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
                btn.onclick = () => generate(this$);
            })();

            (function homeBtn() {
                showMainUI(true);
                spinner(false);//We hide the spinner
            })();
        };

        async function generate(this$:LawFirm) {
            try {
                await createLetter();
                spinner(false);//We hide the spinner
            } catch (error) {
                console.log(`There was an error: ${error}`)
                spinner(false);//We hide the spinner
            }

            async function createLetter() {
                spinner(true);
                const input = byID('textInput') as HTMLTextAreaElement;
                if (!input) return;
                const { templatePath, saveTo } = this$.getConsts(settingsNames.letter);

                const fileName = prompt('Provide the file name without special characthers');
                if (!fileName) return;
                const saveToPath = `${prompt('Provide the destination folder', saveTo || 'NO SAVE TO PATH PROVIDED')}/${fileName}.docx`;

                if (!saveToPath) return;
                const contentControls: [string, string][] = [['RTCoreText', input.value], ['RTReference', 'Référence'], ['RTClientName', 'Nom du Client'], ['RTEmail', 'Email du client']];

                await new GraphAPI('', saveToPath).createAndUploadDocumentFromTemplate(templatePath, saveToPath, 'FR', undefined, contentControls);
            }
        };

    }

    async issueLeaseLetter() {
        spinner(true);//We show the spinner
        const { workbookPath, tableName, templatePath, saveTo } = this.getConsts(this.settingsNames.leases);

        if ([this.stored, workbookPath, tableName, templatePath, saveTo].find(v => !v)) throwAndAlert('One of the  constant values is not valid');

        const graph = new GraphAPI('', workbookPath);
        const tableRows = await graph.fetchExcelTable(tableName, false);//We are calling the "/rows" endPoint, so we will get the tableBody without the headers

        const Ctrls: LeaseCtrls = {
            owner: { title: 'RTBailleur', col: 0, label: 'Nom du Bailleur', type: 'select', value: '' },
            adress: { title: 'RTAdresseDestinataire', label: 'Adresse du bien loué', col: 1, type: 'select', value: '' },
            tenant: { title: 'RTLocataire', label: 'Nom du Locataire', col: 2, type: 'select', value: '' },
            leaseDate: { title: 'RTDateBail', label: 'Date du Bail', col: 3, type: 'date', value: '' },
            leaseType: { title: 'RTNature', label: 'Nature du Bail', col: 4, type: 'text', value: '' },
            initialIndex: { title: 'RTIndiceInitial', label: 'Indice initial', col: 5, type: 'number', value: '' },
            indexQuarter: { title: 'RTTrimestre', label: 'Trimestre de l\'indice', col: 6, type: 'text', value: '' },
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
            startingMonth: { title: 'RTMoisRévision', value: '' },
        };

        const ctrls = Object.values(Ctrls);

        const findRT = (id: string) => ctrls.find(RT => RT.title === id);
        const fraction = (n: number) => Math.round(n * 100)/ 100;

        let row: any[] | void, rowIndex: number | null = null;
        await showForm(this);

        async function showForm(this$:LawFirm) {
            const inputs: InputCol[] = [];
            const findInput = (RT: RT) => inputs.find(([input, col]) => input.id === RT.title)?.[0];
            if (!tableRows) return;
            document.querySelector('table')?.remove();
            const form = byID();
            if (!form) return;
            form.innerHTML = '';
            const divs: HTMLDivElement[] = [];

            (function insertInputs() {
                const unvalid = (values: (string | undefined)[]) => values.find(value => !value || isNaN(Number(value)));
                ctrls
                    .filter(RT => !isNaN(RT.col!))
                    .map(RT => inputs.push([createInput(RT), RT.col!] as const));

                const owner = findInput(Ctrls.owner);
                if (owner) populateSelectElement(owner, getUniqueValues(Ctrls.owner.col!, tableRows), false);

                (function inputsOnChange() {
                    const filled = inputs.filter(([input, col]) => col <= Ctrls.tenant.col!);
                    filled.forEach(([input, col]) => input.onchange = () => [row, rowIndex] = this$.inputOnChange(col, inputs, tableRows, false) || [undefined, null]);

                    const index = findInput(Ctrls.index);
                    const currentLeaseInput = findInput(Ctrls.currentLease);
                    
                    index!.onchange = () => {
                        if (!row?.length)
                            return alert('No single lease having owner name, property adress and tenant name as in the inputs was found');
                        const initial = row[Ctrls.initialIndex.col!];//This is the value of the inital index
                        const base = row[Ctrls.index.col!] || initial;//!For the base index, we will retrieve the value of the "Indice de Révision" (column 10) from the Excel row. We will not retrieve this value from the input but from the row itself. If this is the first time we are indexing the lease, we will fall back to the intial index (i.e., the value indicated in the lease agreement)
                        const latestIndex = index!.valueAsNumber; //this is the latest index as provided by the user when the input.onChange() event was fired
                        const currentLease = row[Ctrls.currentLease.col!];//This is the value of the current lease
                        if (unvalid([base, latestIndex, currentLease])) return alert('Please make sure that the values of the current lease, the base indice and the new indice are all provided and valid numbers');

                        Ctrls.currentLease.value = currentLease;//!We immediately set the value of this control at this stage, because we will escape this Ctrl when we will update Ctrls values from the inputs, because the corresponding input will be showing the new lease value not the original value

                        (function newLease(){
                            const newLease = fraction(currentLease * (latestIndex / base));//we get a 2 digits fractions from the value
                            currentLeaseInput!.valueAsNumber = newLease;//!We only update the input value, NOT the value of the Excel row (row). We need to keep the initial value in case the user wants to correct  the value in the index input which means we will need to recalculate the newLease value based on the current lease value. We will hence keep the current lease value unchanged until the generate() function is called.
                            Ctrls.newLease.value = newLease;//We update the new lease RT
                            Ctrls.baseIndex.value = latestIndex;//We update  the value of the base index with the latest index
                        })();
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
                btn.onclick = () => generate(inputs, row);
            })();

            (function homeBtn() {
                showMainUI(true);
                spinner(false);//We hide the spinner
            })();

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
                    const col = RT.col!.toString();
                    input.dataset.index = col;
                    div.dataset.index = col;
                    append(input);
                    divs.push(div);
                    return input as HTMLInputElement
                };
            };
        };
        async function generate(inputs: InputCol[], row: any[] | void) {

            if (!inputs.length) return throwAndAlert('The inputs collection is missing');
            if (!row) return throwAndAlert('The values in the inputs did not point to a unique lease in the Excel table');  

            const date = new Date();
            const fileName = prompt('Provide the file name without special characthers') || '';
            const savePath = prompt('Provide the destination folder', `${saveTo}/${fileName}_${date.getFullYear()}${date.getMonth() + 1}${date.getDate()}@${date.getHours()}-${date.getMinutes()}.docx`);
            if (!savePath) return alert('The path for saving the file is not valid');

            inputs.map(([input, col]) => {
                const RT = findRT(input.id) as RT;
                if(RT=== Ctrls.currentLease) return; //! We DO NOT update the value of the current lease from the input because the value in the input is the new lease value after revision, not the original value. We need to keep the original value
                if (RT.type === 'date') RT.value = getISODate(input.valueAsDate);
                else if (RT.type === 'number') RT.value = fraction(input.valueAsNumber);
                else RT.value = input.value
            });

            (function setMissingValues() {
                const anniversary = (year: number, date: Date) => { date.setFullYear(year); return getDateString(date) };
                
                const leaseDate = dateFromExcel(row[Ctrls.leaseDate.col!]);
                const year = date.getFullYear();
                
                Ctrls.revisionDate.value = getISODate(date);//!This Ctrl is associated with a column in the table, that's why we are setting its value to ISO date in order to update the excel table later with a valid date format

                (function withNoColumn() {
                    Ctrls.initialYear.value = getIndexYear(row[Ctrls.initialIndexDate.col!]);
                    Ctrls.baseYear.value = getIndexYear(row[Ctrls.baseIndexDate.col!]);
                    Ctrls.revisionYear.value = year.toString();
                    Ctrls.anniversaryDate.value = anniversary(year, leaseDate);
                    Ctrls.nextRevision.value = anniversary(year + 1, leaseDate);
                    Ctrls.startingMonth.value = `${new Intl.DateTimeFormat('fr-FR', { month: 'long' }).format(date)} ${year.toString()}`;
                })();

                function getIndexYear(date: number) {
                    const newDate = dateFromExcel(date);
                    const month = newDate.getMonth();
                    if (month < 3) {
                        //if the date of publication of the index is within the 1st quarter of the year, it means the index is the index of Q4 of the previous year
                        return newDate.getFullYear() -1
                    } else {
                        //The year of the index is the same year as the year of its publication
                        return newDate.getFullYear();
                    }
                }
            })();

            const contentControls: [string, string][] = ctrls.map(RT =>{
                if (RT.type === 'date') return [RT.title, getDateString(new Date(RT.value) || null)]
                else return [RT.title, RT.value.toString()];
            });
            try {
                await graph.createAndUploadDocumentFromTemplate(templatePath, savePath, 'FR', undefined, contentControls);
                await updateExcelTable();
                spinner(false);//We hide the spinner
            } catch (error) {
                console.log(error);
                alert(error);
                spinner(false);//We hide the spinner
            }

            async function updateExcelTable() {
                Ctrls.currentLease.value = Ctrls.newLease.value; //!This must be done at this stage NOT EARLIER, otherwise, we will lose the value of the original lease when editing the Word template. Notice  that Ctrls.newLease is not associated with a column of the Excel table, (i.e., Ctrls.newLease.col property is undefined), which means the row[] will not updated from this Ctrl.
                    if (row && rowIndex) {
                        await graph.updateExcelTableRow(tableName, rowIndex, update(row));
                    } else {
                        row = Array(inputs.length);
                        await graph.addRowToExcelTable(update(row), rowIndex, tableName);
                    }
                    function update(row:any[]) {
                        ctrls.filter(ctrl=>ctrl.col).forEach(({value, col}) => row[col!] = value);
                        return row
                    }
            };
        }
    }

    async searchFiles() {
        spinner(true);//We show the spinner
        const graph = new GraphAPI('');
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

            try {
                await fetchAndFilter();
                spinner(false);//Hide the spinner
            } catch (error) {
                spinner(false);//Hide the spinner
                console.log(error);
                alert(error);
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
                        const body = { requests: requests };
                        const response = await graph.sendRequest(batchUrl, 'POST', body, undefined, "application/json", "Error fetching subfolders")
                        if (!response?.ok) return;
                        return await response?.json();
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
                const response = await graph.sendRequest(url, 'GET', undefined, undefined, undefined, 'Error fetching items from endpoint');
                if (!response?.ok) return;
                return await response.json();
            };

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

    private findSetting = (name: string, settings: settingInput[] | undefined) => settings?.find(setting => setting.name === name);

    private getConsts(setting: { workBook: string, tableName: string, wordTemplate: string, saveTo: string }) {
        const workbookPath = this.findSetting(setting.workBook, this.stored)?.value || prompt('Provide the Excel workbook path') || '';
        const tableName = this.findSetting(setting.tableName, this.stored)?.value || prompt('Provide the name of the Excel table containing the data') || '';
        const templatePath = this.findSetting(setting.wordTemplate, this.stored)?.value || prompt('Provide the path for the Word invoice template') || 'MISSING TEMPLATE PATH';
        const saveTo = this.findSetting(setting.saveTo, this.stored)?.value || prompt('Provide teh path for the folder where the invoice should be saved') || 'MISSING SAVETO PATH';
        return { workbookPath, tableName, templatePath, saveTo }
    }

    /**
     * Updates the data list or the value of bound inputs according to the value of the input that has been changed
     * @param {number} index - the dataset.index of the input that has been changed
     * @param {any[][]} table - The table that will be filtered to update the data list of the button. If undefined, it means that the data list will not be updated.
     * @param {boolean} combine - If true, it means that the dataList of the next bound input, will include an additional option combining all the options in the dataList
     * @returns 
     */
    private inputOnChange(index: number, inputs: InputCol[], table: any[][] | undefined, combine: boolean): [any[], number] | void {
        if (!table?.length) return;

        const filledInputs =
            inputs
                .filter(([input, col]) => input.value && col <= index)//Those are all the inputs that the user filled with data

        const filtered = this.filterTableByInputsValues(filledInputs, table);//We filter the table based on the filled inputs

        if (!filtered.length) return;

        const boundInputs = inputs.filter(([input, col]) => col > index)//Those are the inputs for which we want to create  or update their data lists

        for (const [input, col] of boundInputs) {
            input.value = ''; //We reset the value of all bound inputs.
            const list = getUniqueValues(col, filtered);
            const row = fillBound(list, input);
            if (row) return [row, table.indexOf(row)];//!If the function returns true, it means that we filled the value of all the bound inputs, so we break the loop. If it returns false, it means that there is more than one value in the list, so we need to create or update the data list of the input.
            //if (fillBound(list, input)) break;//!If the function returns true, it means that we filled the value of all the bound inputs, so we break the loop. If it returns false, it means that there is more than one value in the list, so we need to create or update the data list of the input.
            populateSelectElement(input, list, combine);
        }

        function fillBound(list: any[], input: HTMLInputElement): void | any[] {
            if (list.length > 1) return;
            const value = list[0], found = filtered.length < 2;
            if (!found) return setValue(input, value);//If the filtered array contains more than one row with the same unique value in the corresponding column, we will not fill the next inputs
            const row = filtered[0];//This is the unique row in the filtered list, we will use it to fill all the other inputs
            boundInputs.forEach(([input, col]) => setValue(input, row[col]));
            return row;
        }

        function setValue(input: HTMLInputElement, value: any) {
            if (input.type === "date")
                input.value = getISODate(dateFromExcel(value));//!We must convert the dates from Excel, and pass the ISO date to the input value (NOT to the input.valueAsDate) in order to avoid the timezone offset issue when using input.valueASDate
            else input.value = value?.toString() || '';
        };
    };

    /**
     * Filters the table according to the values of the inputs. The value of each input is compared to the value of the cell in the corresponding column in the table. If the value of the input is included in the cell value, it means that this row matches the criteria of this input. For a row to be included in the resulting filtered table, it must match the criteria of all the inputs.
     * @param {HTMLInputElement[]} inputs - the html inputs containing the values based on which the table will be filtered
     * @param {any[][]} table - The table that will be filtered
     * @returns {any[][]} - The resulting filtered table
     */
    private filterTableByInputsValues(inputs: InputCol[], table: any[][]): any[][] {
        const values = inputs.map(([input, col]) => [col, input.value.split(splitter)] as const);//!some inputs may contain multiple comma separated values if the user has selected more than one option in the data list. So we split the input value by ", " and we check if the cell value is included in the resulting array.
        return table.filter(row => values.every(([col, value]) => value.includes(row[col])));
    };

    /**
     * Returns a string[][] representing the rows to be inserted in the Word table containing the invoice details
     * @param {string[][]} tableRows - The filtered Excel rows from which the data will be extracted and put in the required format 
     * @param {string} lang - The language in which the invoice will issued
     * @returns {string[][]} - the rows to be added to the table. Each row has 4 elements
     */
    getRowsData(tableRows: any[][], discount: number = 0, lang: string, invoiceNumber: string): { wordRows: string[][], totalsLabels: string[] } {

        const labels: { [index: string]: lable } = {
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
        }
        const totalsLabels: string[] = [];
        const colDate = 3, colAmount = 9, colVAT = 10, colHours = 7, colRate = 8, colNature = 2, colDescr = 14;//Indexes of the Excel table columns from which we extract the date 

        const wordRows: string[][] = tableRows.map(row => {
            const date = dateFromExcel(Number(row[colDate]));
            const time = getTimeSpent(Number(row[colHours]));

            let description = `${String(row[colNature])} : ${String(row[colDescr])}`;//Column Nature + Column Description;

            //If the billable hours are > 0, we add to the description: time spent and hourly rate
            if (time)
                description += ` (${labels.hourlyBilled[lang as keyof lable]} ${time}, ${labels.hourlyRate[lang as keyof lable]} ${Math.abs(row[colRate]).toString()}\u00A0€).`;


            const rowValues: string[] = [
                getDateString(date),//Column Date
                description,
                getAmountString(row[colAmount] * -1), //Column "Amount": we inverse the +/- sign for all the values 
                getAmountString(Math.abs(row[colVAT])), //Column VAT: always a positive value
            ];
            return rowValues;
        });
        pushTotalsRows();
        return { wordRows, totalsLabels };

        function pushTotalsRows() {
            //Adding rows for the totals of the different categories and amounts
            const total = (lable: lable) => [colAmount, colVAT].map(col => sumColumn(col, lable.nature)) as values;//!It always returns the absolute values of the total amount and the total VAT
            const amount = (v: values) => v[0];
            const totalFees = total(labels.totalFees);
            const feesDiscount = totalFees.map(amount => amount * (discount / 100));//This is an additional discount applied when the invoice is issued. The Excel table may already include other discounts registered as "Remise"
            const feesDeductions = total(labels.totalDeduction).map((amount, index) => amount += feesDiscount[index]) as values;//This is the total of the deductions from the fees: the "Remise" deductions, and the additional discount added at the time the invoice is issued
            const netFees = totalFees.map((amount, index) => amount - feesDeductions[index]) as values;
            const totalPayments = total(labels.totalPayments);
            const totalExpenses = total(labels.totalExpenses);
            const totalTimeSpent: values = [sumColumn(colHours), NaN];//by omitting to pass the "natures" argument to sumColumn, we do not filter the "Total Time" column by any crieteria. We will get the sum of all the column. since the VAT = NaN, the VAT cell will end up empty.
            const totalDue = netFees.map((amount, index) => amount + totalExpenses[index] - totalPayments[index]) as values;
            const percentage = (amount(feesDeductions) / amount(totalFees)) * 100;

            ['EN', 'FR'].forEach((lang) => (labels.totalDeduction[lang as keyof lable] as string) += ` (${percentage}%)`);

            (function pushTotalsRows() {
                pushRow(labels.totalFees, totalFees);
                pushRow(labels.totalDeduction, feesDeductions, !amount(feesDeductions));
                pushRow(labels.netFees, netFees, !(amount(netFees) < amount(totalFees)));//We don't push this row if the there is no deduction applied on the fees or if the deduction is = 0
                pushRow(labels.totalTimeSpent, totalTimeSpent, !amount(totalTimeSpent));
                pushRow(labels.totalExpenses, totalExpenses, !amount(totalExpenses));
                pushRow(labels.totalPayments, totalPayments, !amount(totalPayments));
                amount(totalDue) < 0 ? pushRow(labels.totalReinbursement, totalDue) : pushRow(labels.totalDue, totalDue)
            })();


            (function addDiscountRowToExcel() {
                if (!discount) return;
                const newRow = tableRows
                    .find(row => labels.totalFees.nature.includes(row[colNature]));
                if (!newRow) return;
                const [amount, vat] = feesDiscount;//!The discount must be added as a positive number. This is like a payment made by the client
                const descr = prompt('Provide a description for the discount', `Remise sur les honoraires de la facture n° ${invoiceNumber}`) || '';
                const date = getISODate(new Date());
                const cells: [number, string | number][] = [
                    [colNature, 'Remise'],
                    [colAmount, amount],
                    [colVAT, vat],
                    [colDescr, descr],
                    [colDate, date],
                    [colDate + 1, date],
                ];

                cells.forEach(([col, value]) => newRow[col] = value);

                new LawFirm().addNewEntry(true, newRow);
            })();

            function pushRow(rowLable: lable, [amount, vat]: values, ignore: boolean = false) {
                if (ignore || !amount || isNaN(amount)) return;
                const lable = rowLable?.[lang as keyof lable] as string || '';
                if (lable) totalsLabels.push(lable);
                const value = rowLable === labels.totalTimeSpent ? getTimeSpent(amount) : getAmountString(amount);
                wordRows.push(
                    [
                        lable,
                        '',
                        value,
                        getAmountString(vat)//VAT is always a positive value
                    ]);
            }
            /**
             * 
             * @param {number} col - the index of the column to be summed 
             * @param {string[] | null} natures - the natures of the rows to be included in the sum. If null, we include all the rows regardless of their nature
             * @returns 
             */
            function sumColumn(col: number, natures: string[] = []): number {
                let rows = tableRows;
                if (natures.length) rows = tableRows.filter(row => natures.includes(row[colNature]));//If natures is specified, we filter the rows to include only the ones whose nature is included in the natures array 
                return Math.abs(sumArray(rows.map(row => Number(row[col]))));//!We return the absolute value of the total
            }
        }

        function sumArray(values: number[]) {
            let sum = 0;
            values.forEach(value => sum += value);
            return sum
        }

        function getAmountString(value: number): string {
            if (isNaN(value)) return '';

            const amount = value.toLocaleString(`${lang.toLowerCase()}-${lang.toUpperCase()}`, {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            });

            const versions = {
                FR: `${amount}\u00A0€`,
                EN: `€\u00A0${amount}`,
            }

            return versions[lang as keyof typeof versions];
        }

        /**
         * Convert the time as retrieved from an Excel cell into 'hh:mm' format
         * @param {number} time - The time as stored in an Excel cell
         * @returns {string} - The time as 'hh:mm' format
         */
        function getTimeSpent(time: number): string {
            if (!time || time <= 0) return '';
            time = time * (60 * 60 * 24)//84600 is the number in seconds per day. Excel stores the time as fraction number of days like "1.5" which is = 36 hours 0 minutes 0 seconds;
            const minutes = Math.floor(time / 60);
            const hours = Math.floor(minutes / 60);
            return [hours, minutes % 60, 0]
                .map(el => el.toString().padStart(2, '0'))
                .join(':');
        }

    }

    getContentControlsValues(invoice: { number: string, clientName: string, matters: string[], adress: string[], lang: string }, date: Date): [string, string][] {
        const fields: InvoiceCtrls = {
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
                value: 'DELETECONTENTECONTROL',//!by setting text = "DELETECONTENTECONTROL", the contentControl will be deleted
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
        return Object.values(fields).map(RT => [RT.title, RT.value as string]);
    }
};


class Marianne extends LawFirm {
    private report: setting = {};

    super() {
    }

    async issueReports() {
        await showReportsForm(this);

        async function showReportsForm(this$: any) {
            spinner(true);//We show the spinner
            const { workbookPath, tableName, templatePath, saveTo } = this$.getConsts(this$.settingsNames.invoices);

            if ([this$.stored, workbookPath, tableName, templatePath, saveTo].find(v => !v)) throwAndAlert('One of the  constant values is not valid');

            const graph = new GraphAPI('', workbookPath);

            const sessionId = await graph.createFileSession() || '';
            if (!sessionId) return throwAndAlert('There was an issue with the creation of the file cession. Check the console.log for more details');

            const tableRows = await graph.fetchExcelTable(tableName, true);
            if (!tableRows?.length) return throwAndAlert('Failed to retrieve the Excel table');
            const tableTitles = tableRows[0];
            document.querySelector('table')?.remove();
            try {
                insertInvoiceForm(tableTitles);
                await graph.closeFileSession(sessionId);
                spinner(false);//We hide the spinner
            }
            catch (error) {
                throwAndAlert(`Error while showing the invoice user form: ${error}`)
                spinner(false);//We hide the spinner
            }

            function insertInvoiceForm(tableTitles: string[]) {
                const form = byID();
                if (!form) throw new Error('The form element was not found');
                form.innerHTML = '';
                const tableBody = tableRows!.slice(1, -1);
                const boundInputs: InputCol[] = [];

                (function insertInputs() {
                    insertInputsAndLables([0, 1, 2, 3, 3], 'input');//Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice
                    insertInputsAndLables(['Discount'], 'discount')[0].value = '0%'; //Inserting a discount percentage input and setting its default value to 0%
                    insertInputsAndLables(['Français', 'English'], 'lang', true); //Inserting languages checkboxes
                })();

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
                    btnIssue.onclick = () => issueReport(this$, tableName, tableTitles, templatePath!, saveTo, graph);
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
                                input.onchange = () => this$.inputOnChange(index as number, boundInputs, tableBody, true);//We add onChange on "Client" (0) and "Affaire" (1) columns
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
        };

        function monthlyReport(this$: any) {

            this$.issueReport()
        }

        function annualReport(this$: any) {

        }

        function returnedReport(this$: any) {

        }

        async function issueReport(this$: any, tableName: string, tableTitles: string[], templatePath: string, saveTo: string, graph: GraphAPI) {
            spinner(true);//We show the spinner
            try {
                await editInvoice(this$, tableName, tableTitles, templatePath, saveTo, graph);
                spinner(false);//We hide the spinner
            } catch (error) {
                spinner(false);//We hide the sinner
                alert(error)
            }
        };

        async function editInvoice(this$: any, tableName: string, tableTitles: string[], templatePath: string, saveTo: string, graph: GraphAPI) {
            const client = tableTitles[0], matter = tableTitles[1];//Those are the 'Client' and 'Matter' columns of the Excel table
            const sessionId = await graph.createFileSession(true) || '';//!persist must be = true. This means that if the session is closed, the changes made to the file will be saved.
            if (!sessionId) return throwAndAlert('There was an issue with the creation of the file cession. Check the console.log for more details');

            const inputs = Array.from(document.getElementsByTagName('input'));
            const criteria = inputs.filter(input => getIndex(input) >= 0);

            const discount = parseInt(inputs.find(input => input.id === 'discount')?.value || '0%');

            const lang = inputs.find(input => input.dataset.language && input.checked === true)?.dataset.language || 'FR';

            const date = new Date();//We need to generate the date at this level and pass it down to all the functions that need it
            const invoiceNumber = getInvoiceNumber(date);
            const data = await filterExcelData(criteria, discount, lang, invoiceNumber);

            if (!data) return throwAndAlert('Could not retrieve the filtered Excel table');

            const { wordRows, totalsLabels } = data;

            const report = {
                number: invoiceNumber,
                clientName: '',
                matters: 'matters',
                adress: '',
                lang: lang
            }

            const contentControls = this$.getContentControlsValues(report, date);

            const fileName = getInvoiceFileName('', [''], invoiceNumber);
            let saveToPath = `${saveTo}/${fileName}`;

            saveToPath = prompt(`The file will be saved in ${saveTo}, and will be named : ${fileName}.\nIf you want to change the path or the name, provide the full file path and name of your choice without any sepcial characters`, saveTo) || saveTo;

            (async function editInvoiceFilterExcelClose() {
                await graph.createAndUploadDocumentFromTemplate(templatePath!, saveToPath, lang, [['Invoice', wordRows, 1]], contentControls, totalsLabels);
                await graph.clearFilterExcelTable(tableName!, sessionId);//We unfilter the table;
                //await graph.filterExcelTable(tableName, client, [clientName], sessionId);//We filter the table by the matters that were invoiced
                //await graph.filterExcelTable(tableName, matter, matters, sessionId);//We filter the table by the matters that were invoiced
                await graph.closeFileSession(sessionId);
            })();

            /**
             * Filters the Excel table according to the values of each inputs, then returns the values of the Word table rows that will be added to the Word table in the invoice template document 
             * @param {HTMLInputElement[]} inputs - the html inputs containing the values based on which the table will be filtered
             * @param {number} discount  - The discount percentage that will be applied to the amount of each invoiced row if any. It is a number between 0 and 100. If it is equal to 0, it means that no discount will be applied.
             * @param {string} lang - The language in which the invoice will be issued 
             * @returns {Promise<[string[][], string[], string[], string[]]>} - The values of the rows that will be added to the Word table in the invoice template
             */
            async function filterExcelData(inputs: HTMLInputElement[], discount: number, lang: string, invoiceNumber: string): Promise<{ wordRows: string[][], totalsLabels: string[] } | void> {
                const clientCol = 0, matterCol = 1, dateCol = 3, addressCol = 15;//Indexes of the 'Matter' and 'Date' columns in the Excel table
                const clientNameInput = getInputByIndex(inputs, clientCol);
                const matterInput = getInputByIndex(inputs, matterCol);
                const clientName = clientNameInput!.value || '';
                const matters =
                    getArray(matterInput!.value) || []; //!The Matter input may include multiple entries separated by ', ' not only one entry.

                if (!clientName || !matters?.length) throwAndAlert('could not retrieve the client name or the matter/matters list from the inputs');
                let tableRows = this$.getExcelTable(tableName);
                tableRows = this$.filterTableByInputsValues([[clientNameInput!, clientCol], [matterInput!, matterCol]], tableRows);

                tableRows = filterByDate(tableRows!, dateCol);


                const { wordRows, totalsLabels } = this$.getRowsData(tableRows, discount, lang, invoiceNumber);

                return { wordRows, totalsLabels };

                function filterByDate(visible: string[][], dateCol: number) {

                    const convert = (date: string | number) => dateFromExcel(Number(date)).getTime();

                    const [from, to] = inputs
                        .filter(input => getIndex(input) === dateCol)
                        .map(input => input.valueAsDate?.getTime());

                    if (from && to)
                        return visible.filter(row => convert(row[dateCol]) >= from && convert(row[dateCol]) <= to); //we filter by the date
                    else if (from)
                        return visible.filter(row => convert(row[dateCol]) >= from); //we filter by the date
                    else if (to)
                        return visible.filter(row => convert(row[dateCol]) <= to); //we filter by the date
                    else
                        return visible.filter(row => convert(row[dateCol]) <= new Date().getTime()); //we filter by the date
                }

            }
        }

        async function getExcelTable(tableName: string, graph: GraphAPI) {
            const excelTable = await graph.fetchExcelTable(tableName, true);
            let tableRows = excelTable?.slice(1, -1) || undefined; //We exclude the first and the last rows of the table. Since we are calling the "range" endpoint, we get the whole table including the headers. The first row is the header, and the last row is the total row.
            if (!tableRows) return throwAndAlert('We could not retrieve the tableRows whie trying to issue the invoice');
            return tableRows
        }
    }

    test() {
        const m = this.report;
        m.saveTo = {
            label: '',
            name: '',
            value: ''
        };
    }

    /**
     * This function isn't used in Marianne Class
     * @returns 
     */
    async issueLetter(): Promise<void> {
        console.log('this is not a valid method in Marianne Class');
        return
    }
    /**
     * This function isn't used in Marianne Class
     * @returns 
     */
    async issueLeaseLetter(): Promise<void> {
        return console.log('this is not a valid method in Marianne Class')
    }

    /**
     * This function isn't used in Marianne Class
     * @returns 
     */
    async searchFiles(): Promise<void> {
        return console.log('this is not a valid method in Marianne Class')
    }
}

(function startApp() {
    showMainUI();//!This must come after the classes have been declared
})();

/**
 * Convert the date in an Excel row into a javascript date (in milliseconds)
 * @param {number} excelDate - The date retrieved from an Excel cell
 * @returns {Date} - a javascript format of the date
 */
function dateFromExcel(excelDate: number): Date {
    const day = 86400000;//this is the milliseconds in a day
    const dateMS = Math.round((excelDate - 25569) * day);//This gives the days converted from milliseconds. 
    //!We have to do this in order to avoid the timezone conversion issues
    const date = new Date(dateMS);
    return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
}

function showMainUI(homeBtn?: boolean) {
    const container = byID('btns');
    if (!container) return;
    container.innerHTML = "";
    if (homeBtn) return appendBtn('home', 'Back to Main', ()=>showMainUI());
    const lf = new LawFirm();
    appendBtn('entry', 'Add Entry', () => lf.addNewEntry());
    appendBtn('invoice', 'Invoice', () => lf.issueInvoice());
    appendBtn('letter', 'Letter', () => lf.issueLetter());
    appendBtn('lease', 'Leases', () => lf.issueLeaseLetter());
    appendBtn('search', 'Search Files', () => lf.searchFiles());
    appendBtn('settings', 'Settings', () => saveSettings());

    function appendBtn(id: string, text: string, onClick: onClick) {
        const btn = document.createElement('button');
        btn.id = id;
        btn.classList.add("ms-Button");
        btn.innerText = text;
        btn.onclick = onClick;
        container?.appendChild(btn);
        return btn
    }
};







