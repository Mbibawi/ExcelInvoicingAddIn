
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
/**
 * 
 * @param {boolean} add - If false, the function will only show a form containing input fields for the user to provide the data for the new row to be added to the Excel Table. If true, the function will parse the values from the input fields in the form, and will add them as a new row to the Excel Table. Its default value is false.
 * @param {any[]} row - If provided, the function will add the row directly to the Excel Table without needing to retrieve the data from the inputs.
 */
async function addNewEntry(add: boolean = false, row?: any[]) {
    accessToken = await getAccessToken() || '';

    (async function showForm() {
        if (add) return;
        if (!workbookPath || !tableName) return alert('The Excel Workbook Path or the name of the Excel table are not valid');

        TableRows = await fetchExcelTableWithGraphAPI(accessToken, workbookPath, tableName) as string[][];

        if (!TableRows) return;

        insertAddForm(TableRows[0]);
    })();

    (async function addEntry() {
        if (!add) return;
        if (row) return await addRow(row);//If a row is already passed, we will add them directly

        await addRow(parseInputs() || undefined, true)

        function parseInputs() {
            const stop = (missing: string) => alert(`${missing} missing. You must at least provide the client, matter, nature, date and the amount. If you provided a time start, you must provide the end time and the hourly rate. Please review your iputs`);

            const inputs = Array.from(document.getElementsByTagName('input')) as HTMLInputElement[];//all inputs

            const nature = getInputByIndex(inputs, 2)?.value;
            if (!nature) return stop('The matter is')
            const date = getInputByIndex(inputs, 3)?.valueAsDate
            if (!date) return stop('The invoice date is');
            const amount = getInputByIndex(inputs, 9);
            const rate = getInputByIndex(inputs, 8)?.valueAsNumber;

            const debit = ['Honoraire', 'Débours/Dépens', 'Débours/Dépens non facturables', 'Rétrocession d\'honoraires'].includes(nature);//We check if we need to change the value sign 

            const row = inputs.map(input => {
                const index = getIndex(input);
                if ([3, 4].includes(index))
                    return getISODate(date);//Those are the 2 date columns
                else if ([5, 6].includes(index))
                    return getTime([input]);//time start and time end columns
                else if (index === 7) {
                    //!This is a hidden input
                    const totalTime = getTime([getInputByIndex(inputs, 5), getInputByIndex(inputs, 6)]);//Total time column

                    if (totalTime > 0 && rate && amount && !amount.valueAsNumber) amount.valueAsNumber = totalTime * 24 * rate// making the amount equal the rate * totalTime
                    return totalTime
                }
                else if (debit && index === 9)
                    return input.valueAsNumber * -1 || 0;//This is the amount if negative
                else if ([8, 9, 10].includes(index))
                    return input.valueAsNumber || 0;//Hourly Rate, Amount, VAT
                else return input.value;
            });

            if (missing()) return stop('Some of the required fields are');

            return row

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
            if (!row) return;
            await addRowToExcelTableWithGraphAPI([row], TableRows.length - 2, workbookPath, tableName, accessToken);

            if (!filter) return;

            [0, 1].map(async index => {
                //!We use map because forEach doesn't await
                await filterExcelTable(workbookPath, tableName, TableRows[0]?.[index], [row[index]?.toString()] || [], accessToken);
            });

            alert('Row aded and table was filtered');

        }

    })()


    function insertAddForm(title: string[]) {
        const form = document.getElementById('form');
        if (!form) return;
        form.innerHTML = '';

        const divs = title.map((title, index) => {
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

// Update Word Document
async function invoice(issue: boolean = false) {
    accessToken = await getAccessToken() || '';

    (async function show() {
        if (issue) return;
        if (!workbookPath || !tableName) return alert('The Excel Workbook path and/or the name of the Excel Table are missing or invalid');

        TableRows = await fetchExcelTableWithGraphAPI(accessToken, workbookPath, tableName) as string[][];

        if (!TableRows) return;

        insertInvoiceForm(TableRows);

    })();

    (async function issueInvoice() {
        if (!issue) return;
        if (!templatePath || !destinationFolder) return alert('The full path of the Word Invoice Template and/or the destination folder where the new invoice will be saved, are either missing or not valid');

        const inputs = Array.from(document.getElementsByTagName('input'));

        const criteria = inputs.filter(input => Number(input.dataset.index) >= 0);

        const discount = parseInt(inputs.find(input => input.id === 'discount')?.value || '0%');

        const lang = inputs.find(input => input.dataset.language && input.checked === true)?.dataset.language || 'FR';

        TableRows = await fetchExcelTableWithGraphAPI(accessToken, workbookPath, tableName) as string[][];//We fetch the table again in case there where changes made since it was fetched the first time when the userform was inserted

        const [wordRows, totalsRows, filtered] = filterExcelData(TableRows, criteria, discount, lang);

        const date = new Date();

        const invoice = {
            number: getInvoiceNumber(date),
            clientName: getInputValue(0, criteria),
            matters: getArray(getInputValue(1, criteria)),
            adress: getUniqueValues(15, filtered) as string[],
            lang: lang
        }
        const contentControls = getContentControlsValues(invoice, date);

        const fileName = getInvoiceFileName(invoice.clientName, invoice.matters, invoice.number);
        let filePath = `${destinationFolder}/${fileName}`;

        filePath = prompt(`The file will be saved in ${destinationFolder}, and will be named : ${fileName}./nIf you want to change the path or the name, provide the full file path and name of your choice without any sepcial characters`, filePath) || filePath;

        await createAndUploadXmlDocument(wordRows, contentControls, accessToken, templatePath, filePath, totalsRows);

        (async function filterTable() {
            await clearFilterExcelTableGraphAPI(workbookPath, tableName, accessToken); //We start by clearing the filter of the table, otherwise the insertion will fail
            [0, 1].map(async index => {
                await filterExcelTable(workbookPath, tableName, TableRows[0][index], getUniqueValues(index, filtered) as string[], accessToken)
            });
        })();

        /**
         * Filters the Excel table according to the values of each inputs, then returns the values of the Word table rows that will be added to the Word table in the invoice template document 
         * @param {any[][]} data - The Excel table rows that will be filtered
         * @param {HTMLInputElement[]} criteria - the html inputs containing the values based on which the table will be filtered
         * @param {string} lang - The language in which the invoice will be issued 
         * @returns {string[][]} - The values of the rows that will be added to the Word table in the invoice template
         */
        function filterExcelData(data: any[][], criteria: HTMLInputElement[], discount: number, lang: string): [string[][], string[], any[][]] {

            //Filtering by Client (criteria[0])
            data = data.filter(row => row[getIndex(criteria[0])] === criteria[0].value);
            const adress = getUniqueValues(15, data);//!We must retrieve the adresses at this stage before filtering by "Matter" or any other column

            [1, 2].forEach(index => {
                //!Matter and Nature inputs (from columns 2 & 3 of the Excel table) may include multiple entries separated by ', ' not only one entry.
                const list = criteria[index].value.split(',').map(el => el.trimStart().trimEnd());//We generate a string[] from the input.value
                data = data.filter(row => list.includes(row[index]));//We filter the data
            });
            //We finaly filter by date
            data = filterByDate(data);

            return [...getRowsData(data, discount, lang), data];

            function filterByDate(data: string[][]) {

                const convertDate = (date: string | number) => dateFromExcel(Number(date)).getTime();

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

    function insertInvoiceForm(excelTable: string[][]) {
        const form = document.getElementById('form');
        if (!form) return;
        form.innerHTML = '';
        const title = excelTable[0];

        insertInputsAndLables([0, 1, 2, 3, 3], 'input');//Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice

        insertInputsAndLables(['Discount'], 'discount', false)[0].value = '0%'; //Inserting a discount percentage input and setting its default value to 0%

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
                appendLable(index);
                return appendInput(index);
            });

            function appendInput(index: number | string) {
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
                        input.onchange = () => inputOnChange(Number(input.dataset.index), excelTable.slice(1, -1), true);

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

                form?.appendChild(input);

                return input;
            }

            function appendLable(index: number | string) {
                const label = document.createElement('label');
                isNaN(Number(index)) || checkBox ? label.innerText = index.toString() : label.innerText = title[Number(index)];
                !isNaN(Number(index)) ? label.htmlFor = id + index.toString() : label.htmlFor = id;
                form?.appendChild(label);
            }
        };

    }

}

async function issueLetter(create: boolean = false) {
    accessToken = await getAccessToken() || '';
    const templatePath = '';
    (function showForm() {
        if (create) return;
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

        createAndUploadXmlDocument(undefined, contentControls, accessToken, templatePath, filePath);
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


async function createAndUploadXmlDocument(rows: string[][] | undefined, contentControls: string[][] | undefined, accessToken: string, templatePath: string, filePath: string, totals: string[] = []) {

    if (!accessToken || !templatePath || !filePath) return;

    const blob = await fetchFileFromOneDriveWithGraphAPI(accessToken, templatePath);

    if (!blob) return;

    const [doc, zip] = await convertBlobIntoXML(blob);

    if (!doc) return;

    (function editTable() {
        if (!rows) return;
        const table = getXMLElements(doc, "w:tbl", 1) as Element;
        if (!table) return;

        rows.forEach((row, index) => {
            const newXmlRow = insertRowToXMLTable();
            if (!newXmlRow) return;
            const isTotal = totals.includes(row[0]);
            const isLast = index === rows.length - 1;
            row.forEach((text, index) => {
                addCellToXMLTableRow(newXmlRow, getStyle(index, isTotal && !isLast), [isTotal, isLast].includes(true), text)
            })
        });

        function insertRowToXMLTable(after: number = -1) {
            if (!table) return;
            const row = createXMLElement("w:tr");
            after >= 0 ? (getXMLElements(table, 'w:tr', after) as Element)?.insertAdjacentElement('afterend', row) :
                table.appendChild(row);
            return row;
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

        function setStyle(targetElement: Element, style: string, backGroundColor: string = ''): void {
            // Create or find the run properties element
            //const styleProps = createAndAppend(runElement, "w:rPr", false);
            const tag = targetElement.tagName.toLocaleLowerCase();
            (function cell() {
                if (tag !== 'w:tc') return;
                const cellProp = createAndAppend(targetElement, 'w:tcPr', false);
                createAndAppend(cellProp, 'w:vAlign').setAttribute('w:val', "center");
                //createAndAppend(cellProp, 'w:tcStyle').setAttribute('w:val', 'InvoiceCellCentered');
                if (!backGroundColor) return;
                const background = createAndAppend(cellProp, 'w:shd');//Adding background color to cell
                background.setAttribute('w:val', "clear");
                background.setAttribute('w:fill', backGroundColor);
            })();

            (function parag() {
                if (tag !== 'w:p') return;
                if (!style) return;
                const props = createAndAppend(targetElement, "w:pPr", false);
                createAndAppend(props, "w:pStyle").setAttribute("w:val", style);
            })();

            function createAndAppend(parent: Element, tag: string, append: boolean = true) {
                const newElement = createXMLElement(tag);
                if (append) parent.appendChild(newElement)
                else parent.insertBefore(newElement, parent.firstChild);
                return newElement
            }
        }

        function addCellToXMLTableRow(row: Element, style: string, isTotal: boolean, text?: string) {
            if (!row) return;
            const cell = createXMLElement("w:tc");//new table cell
            row.appendChild(cell);
            if (isTotal)
                setStyle(cell, style, 'D9D9D9');//We set the background color of the cell
            else setStyle(cell, style, '');
            const parag = createXMLElement("w:p");//new table paragraph
            cell.appendChild(parag)
            setStyle(parag, style, '');
            const newRun = createXMLElement("w:r");// new run
            parag.appendChild(newRun);

            if (!text) return;

            const newText = createXMLElement("w:t");
            newText.textContent = text;

            newRun.appendChild(newText);

        }
    })();

    (function editContentControls() {
        if (!contentControls) return;
        const ctrls = getXMLElements(doc, "w:sdt") as Element[];
        contentControls
            .forEach(([title, text]) => {
                const control = findXMLContentControlByTitle(ctrls, title);
                if (!control) return;
                editXMLContentControl(control, text);
            });

        function findXMLContentControlByTitle(ctrls: Element[], title: string) {
            return ctrls.find(control => control.getElementsByTagName("w:alias")[0]?.getAttribute("w:val") === title);
        }

        function editXMLContentControl(control: Element, text: string) {
            if (!text) return control.remove();
            const paragraphs = text.split('/n');
            if (paragraphs.length > 1)
                return paragraphs.forEach(parag => addNewParagraph(parag));


            const textElement = control.getElementsByTagName("w:t")[0];

            if (!textElement) return addNewParagraph(text);

            textElement.textContent = text;

            function addNewParagraph(text: string) {
                const sdtContent = control.getElementsByTagName("sdtContent")[0];
                if (!sdtContent) alert("No sdtContent");
                const paragElement = createXMLElement("w:p"); // Create a new paragraph
                const runElement = createXMLElement("w:r"); // Create a run
                const textElement = createXMLElement("w:t"); // Create text element
                textElement.textContent = text; // Set the paragraph text
                runElement.appendChild(textElement);
                paragElement.appendChild(runElement);
                sdtContent.appendChild(paragElement); // Add paragraph to the content control
            }
        }

    })();

    await convertXMLToBlobAndUpload(doc, zip, filePath, accessToken);

    function createXMLElement(tag: string) {
        return doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', tag);
    }

    function getXMLElements(xmlDoc: XMLDocument | Element, tag: string, index?: number): Element[] | Element {
        const elements = xmlDoc.getElementsByTagName(tag);
        if (index) return elements[index];
        return Array.from(elements)
    }
};

/**
 * Converts the blob of a Word document into an XML
 * @param blob - the blob of the file to be converted
 * @returns {[XMLDocument, JSZip]} - The xml document, and the zip containing all the xml files
 */
//@ts-expect-error
async function convertBlobIntoXML(blob: Blob): [XMLDocument, JSZip] {
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

async function addRowToExcelTableWithGraphAPI(row: any[][], index: number, filePath: string, tableName: string, accessToken: string) {

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
    } else {
        alert(`Error adding row: ${await response.text()}`);
    }

}

async function filterExcelTable(filePath: string, tableName: string, columnName: string, values: string[], accessToken: string) {
    if (!accessToken) return;

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
    } else {
        alert(`Error applying filter: ${await filterResponse.text()}`);
    }
}



