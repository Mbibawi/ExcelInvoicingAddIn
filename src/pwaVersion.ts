// Authentication
//const accessToken = getAccessToken();


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

async function addNewEntry(add: boolean = false) {
    accessToken = await getAccessToken() || '';

    (async function show() {
        if (add) return;
        TableRows = await fetchExcelTableWithGraphAPI(accessToken, excelPath, tableName);

        if (!TableRows) return;

        insertAddForm(TableRows[0]);
    })();

    (async function addEntry() {
        if (!add) return;
        const inputs = Array.from(document.getElementsByTagName('input')) as HTMLInputElement[];//all inputs
        const nature = getInputByIndex(inputs, 2)?.value || '';
        const date = getInputByIndex(inputs, 3)?.valueAsDate || undefined;
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

        const stop = 'You must at least provide the client, matter, nature, date and the amount. If you provided a time start, you must provide a time end, and an hourly rate. Please review your fields';

        if (missing()) return alert(stop);

        function missing() {
            if (row[5] === row[6]) return false;//If the total time = 0 we do not need to alert if the hourly rate is missing
            else if (row.filter((el, i) => (i < 4 || i === 9) && !el).length > 0) return true;//if client name, matter, nature, date or amount are missing
            //else if (row[9]) return [5, 6,7,8].map(index => row[index] = 0).length < 1;//This means the amount has been provided and does not  depend on the time spent or the hourly rate. We set the values of the startTime and endTime to 0, and return false (length<1 must return false)
            else if (row[5] && (!row[6] || !row[8]))
                return true//if startTime is provided but without endTime or without hourly rate
            else if (row[6] && (!row[5] || !row[8]))
                return true//if endTime is provided but without startTime or without hourly rate
        };


        await addRowToExcelTableWithGraphAPI([row], TableRows.length - 2, excelFilePath, tableName, accessToken);

        [0, 1].map(async index => {
            //!We use map because forEach doesn't await
            //@ts-ignore
            await filterExcelTable(excelFilePath, tableName, TableRows[0][index], row[index].toString(), accessToken);
        });

        alert('Row aded and table was filtered');


    })()


    function insertAddForm(title: string[]) {
        const form = document.getElementById('form');
        if (!form) return;
        form.innerHTML = '';

        title.forEach((title, index) => {
            if (![4, 7, 15].includes(index)) form.appendChild(createLable(title, index));//We exclued the labels for "Total Time" and for "Year"
            form.appendChild(createInput(index));
        });

        for (const t of title) {//!We could not use for(let i=0; i<title.length; i++) because the await does not work properly inside this loop
        };

        (function addBtn() {
            const btnIssue = document.createElement('button');
            btnIssue.innerText = 'Add Entry';
            btnIssue.classList.add('button');
            btnIssue.onclick = () => addNewEntry(true);
            form.appendChild(btnIssue);
        })();

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
                else if ([4, 7, 15].includes(index)) input.style.display = 'none';//We hide those 3 columns: 'Total Time' and the 'Year' and 'Link to a File'
                else if (index < 3 || index > 10) {
                    //We add a dataList for those fields
                    input.setAttribute('list', input.id + 's');
                    input.onchange = () => inputOnChange(index, TableRows.slice(1, -1), false);
                    if (![1, 16].includes(index))
                        createDataList(input.id, getUniqueValues(index, TableRows.slice(1, -1), tableName));//We don't create the data list for columns 'Matter' (1) and 'Adress' (16) because it will be created when the 'Client' field is updated
                }

                if (index > 4 && index < 11)
                    //Those are the "Start Time", "End Time", "Total Time", "Hourly Rate", "Amount", "VAT" columns . The "Hourly Rate" input is hidden, so it can't be changed by the user. We will add the onChange event to it by simplicity
                    input.onchange = () => inputOnChange(index, undefined, false);//!We are passing the table[][] argument as undefined, and the invoice argument as false 
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

        TableRows = await fetchExcelTableWithGraphAPI(accessToken, excelPath, tableName);

        if (!TableRows) return;

        insertInvoiceForm(TableRows);

    })();

    (async function issueInvoice() {
        if (!issue) return;
        const inputs = Array.from(document.getElementsByTagName('input'));

        const criteria = inputs.filter(input => Number(input.dataset.index) >= 0);

        const lang = inputs.find(input => input.dataset.language && input.checked === true)?.dataset.language || 'FR';

        const filtered = filterExcelData(TableRows, criteria, lang);

        const date = new Date();

        const invoice = {
            number: getInvoiceNumber(date),
            clientName: getInputValue(0, criteria),
            matters: getArray(getInputValue(1, criteria)),
            adress: getArray(getInputValue(16, criteria)),
            lang: lang
        }
        const contentControls = getContentControlsValues(invoice, date);

        const filePath = `${destinationFolder}/${getInvoiceFileName(invoice.clientName, invoice.matters, invoice.number)}`;

        await createAndUploadXmlDocument(filtered, contentControls, accessToken, filePath);

        function filterExcelData(data: any[][], criteria: HTMLInputElement[], lang: string) {

            //Filtering by Client (criteria[0])
            data = data.filter(row => row[getIndex(criteria[0])] === criteria[0].value);


            [1, 2].forEach(index => {
                //!Matter and Nature inputs (from columns 2 & 3 of the Excel table) may include multiple entries separated by ', ' not only one entry.
                const list = criteria[index].value.replaceAll(' ', '').split(',');//We generate a string[] from the input.value
                data = data.filter(row => list.includes(row[index]));//We filter the data
            });
            //We finaly filter by date
            data = filterByDate(data);

            return getRowsData(data, lang);

            function filterByDate(data: string[][]) {
                const [from, to] = criteria
                    .filter(input => getIndex(input) === 3)
                    .map(input => new Date(input.value).getTime());

                const convertDate = (date: string | number) => dateFromExcel(Number(date)).getTime();

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
        insertInputsAndLables([0, 1, 2, 3, 3]);//Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice
        insertInputsAndLables(['Français', 'English'], true); //Inserting langauges checkboxes

        (function addBtn() {
            const btnIssue = document.createElement('button');
            btnIssue.innerText = 'Generate Invoice';
            btnIssue.classList.add('button');
            btnIssue.onclick = () => invoice(true);
            form.appendChild(btnIssue);
        })();

        function insertInputsAndLables(indexes: (number | string)[], checkBox: boolean = false): HTMLInputElement[] {
            let id = 'input';
            let css = 'field';
            if (checkBox) css = 'checkBox';

            return indexes.map(index => {
                checkBox ? id = id : id+= index.toString();
                appendLable(index);
                return appendInput(Number(index));
            });

            function appendInput(index:number) {
                const input = document.createElement('input');
                input.classList.add(css);
                input.id = id;

                (function setType() {               
                    if (checkBox) input.type = 'checkbox';
                    else if (index < 3) input.type = 'text';
                    else input.type = 'date';
                })();

                (function notCheckBox() { 
                    if (checkBox) return;
                    input.name = input.id;
                    input.dataset.index = index.toString();
                    input.setAttribute('list', input.id + 's');
                    input.autocomplete = "on";

                    if (index < 2)
                        input.onchange = () => inputOnChange(Number(input.dataset.index), excelTable.slice(1, -1), true);

                    if (index < 1)
                        createDataList(id, Array.from(new Set(TableRows.slice(1, -1).map(row => row[0]))));//We create a unique values dataList for the 'Client' input
                })();

                (function isCheckBox() {
                    if (!checkBox) return;
                    input.dataset.language = index.toString().slice(0, 2).toUpperCase();
                })();

                form?.appendChild(input);

                return input;
            }

            function appendLable(index:number|string) {
                const label = document.createElement('label');
                checkBox ? label.innerText = index.toString() : label.innerText = title[Number(index)];
                label.htmlFor = id;
                form?.appendChild(label);
            }
        };

    }

}

/**
 * Updates the data list or the value of bound inputs according to the value of the input that has been changed
 * @param {number} index - the dataset.index of the input that has been changed
 * @param {any[][]} table - the table that will be filtered. If undefined, it means that no data list will be updated.
 * @param {boolean} invoice - If true, it means that we called the function in order to generate an invoice. If false, we called it in order to add a new entry in the table
 * @returns 
 */
function inputOnChange(index: number, table: any[][] | undefined, invoice: boolean) {
    let inputs = Array.from(document.getElementsByTagName('input') as HTMLCollectionOf<HTMLInputElement>);

    if (!table && !invoice) {
        const boundInputs = [5, 6, 7, 9, 10];//Those are "Start Time" (5), "End Time" (6), "Total Time" (7, although it is hidden), "Amount" (9), "VAT" (10) columns. We exclude the "Hourly Rate" column (8). We let the user rest it if he wants
        boundInputs
            .forEach(i => i > index ? reset(i) : i = i);

        if (index === 9)
            boundInputs
                .forEach(i => i < index ? reset(i) : i = i);


        function reset(i: number) {
            const input = getInputByIndex(inputs, i);
            if (!input) return;
            input.valueAsNumber = 0;
            input.value = '';
        }
    }

    if (!table) return;

    if (invoice)
        inputs = inputs.filter(input => input.dataset.index && Number(input.dataset.index) < 3); //Those are all the inputs that serve to filter the table (first 3 columns only)
    else
        inputs = inputs.filter(input => [0, 1, 16].includes(getIndex(input))); //Those are all the inputs that have data lists associated with them

    const filledInputs =
        inputs
            .filter(input => input.value && getIndex(input) <= index)
            .map(input => getIndex(input));//Those are all the inputs that the user filled with data


    const boundInputs = inputs.filter(input => getIndex(input) > index);//Those are the inputs for which we want to create  or update their data lists


    if (filledInputs.length < 1 || boundInputs.length < 1) return;

    boundInputs.forEach(input => input.value = '');

    const filtered = filterOnInput(inputs, filledInputs, table);//We filter the table based on the filled inputs

    if (filtered.length < 1) return;

    boundInputs.map(input => createDataList(input?.id, getUniqueValues(getIndex(input), filtered, tableName), invoice));

    if (invoice) {
        const nature = getInputByIndex(inputs, 2);//We get the nature input in order to fill automaticaly its values by a ', ' separated string
        if (!nature) return;
        nature.value = Array.from(document.getElementById(nature?.id + 's')?.children as HTMLCollectionOf<HTMLOptionElement>)?.map((option) => option.value).join(', ');
    }

    function filterOnInput(inputs: HTMLInputElement[], filled: number[], table: any[][]) {
        let filtered: any[][] = table;
        for (let i = 0; i < filled.length; i++) {
            filtered = filtered.filter(row => row[filled[i]].toString() === getInputByIndex(inputs, filled[i])?.value)
        }
        return filtered
    }
};


async function createAndUploadXmlDocument(rows: string[][], contentControls: string[][], accessToken: string, filePath: string) {

    if (!accessToken) return;
    const schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

    return await createAndEditNewXmlDoc();

    async function createAndEditNewXmlDoc() {
        const blob = await fetchFileFromOneDriveWithGraphAPI(accessToken, filePath);
        if (!blob) return;
        const zip = await convertBlobIntoXML(blob);
        const doc = zip.xmlDoc;
        if (!doc) return;
        const table = getXMLElement(doc, "w:tbl", 0);

        rows.forEach((row, x) => {
            const newXmlRow = insertRowToXMLTable(doc, table);
            if (!newXmlRow) return;
            const isTotal = row[0].startsWith('Total');
            const isLast = x === rows.length - 1;
            row.forEach((text, index) => {
                addCellToXMLTableRow(doc, newXmlRow, getStyle(index, isTotal), [isTotal, isLast].includes(true), text)
            })
        });

        contentControls
            .forEach(([title, text]) => {
                const control = findXMLContentControlByTitle(doc, title);
                if (!control) return;
                editXMLContentControl(control, text);
            });

        console.log('doc = ', doc.children[0]);

        const newBlob = await convertXMLIntoBlob(doc, zip.zip);

        await uploadFileToOneDriveWithGraphAPI(newBlob, filePath, accessToken);

        function getStyle(cell: number, isTotal: boolean) {
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
    }

    //await editDocumentWordJSAPI(await copyTemplate()?.id, accessToken, data, getContentControlsValues(invoice.lang))



    async function fetchBlobFromFile(templatePath: string, accessToken: string) {
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/${templatePath}:/content`, {
            method: 'GET',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
            },
        });

        if (!response.ok) throw new Error('Failed to fetch the Word file from OneDrive');

        return await response.blob();

    }

    async function convertBlobIntoXML(blob: Blob) {
        //@ts-ignore
        const zip = new JSZip();

        const arrayBuffer = await blob.arrayBuffer();

        await zip.loadAsync(arrayBuffer);

        const documentXml = await zip.file("word/document.xml").async("string");

        const parser = new DOMParser();

        const xmlDoc = parser.parseFromString(documentXml, "application/xml");

        return { xmlDoc, zip }
    }

    //@ts-expect-error
    async function convertXMLIntoBlob(editedXml: XMLDocument, zip: JSZip) {

        const serializer = new XMLSerializer();
        let modifiedDocumentXml = serializer.serializeToString(editedXml);

        zip.file("word/document.xml", modifiedDocumentXml);

        return await zip.generateAsync({ type: "blob" });
    }

    function getXMLElement(xmlDoc: XMLDocument | Element, tag: string, index: number) {
        const elements = xmlDoc.getElementsByTagName(tag);
        return elements[index];
    }

    function insertRowToXMLTable(xmlDoc: XMLDocument, table: Element, after: number = -1) {
        if (!table) return;
        const row = createTableElement(xmlDoc, "w:tr");
        after >= 0 ? getXMLElement(table, 'w:tr', after)?.insertAdjacentElement('afterend', row) :
            table.appendChild(row);
        return row;
    }

    function setStyle(targetElement: Element, style: string, backGroundColor: string = '', doc: Document): void {
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
            const newElement = createTableElement(doc, tag);
            if (append) parent.appendChild(newElement)
            else parent.insertBefore(newElement, parent.firstChild);
            return newElement
        }
    }

    function addCellToXMLTableRow(xmlDoc: XMLDocument, row: Element, style: string, isTotal: boolean, text?: string) {
        if (!xmlDoc || !row) return;
        const cell = createTableElement(xmlDoc, "w:tc");//new table cell
        row.appendChild(cell);
        if (isTotal)
            setStyle(cell, style, 'D9D9D9', xmlDoc);//We set the background color of the cell
        else setStyle(cell, style, '', xmlDoc);
        const parag = createTableElement(xmlDoc, "w:p");//new table paragraph
        cell.appendChild(parag)
        setStyle(parag, style, '', xmlDoc);
        const newRun = createTableElement(xmlDoc, "w:r");// new run
        parag.appendChild(newRun);

        if (!text) return;

        const newText = createTableElement(xmlDoc, "w:t");
        newText.textContent = text;

        newRun.appendChild(newText);

    }

    function createTableElement(xmlDoc: XMLDocument, tag: string) {
        return xmlDoc.createElement(tag);
    }

    function findXMLContentControlByTitle(xmlDoc: XMLDocument, title: string) {
        const contentControls = Array.from(xmlDoc.getElementsByTagName("w:sdt"));
        return contentControls.find(control => control.getElementsByTagName("w:alias")[0]?.getAttribute("w:val") === title);
    }

    function editXMLContentControl(control: Element, text: string) {
        if (!control) return;
        if (!text) return control.remove();
        const textElement = control.getElementsByTagName("w:t")[0];
        if (!textElement) return;
        textElement.textContent = text;
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
    } else {
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

async function filterExcelTable(filePath: string, tableName: string, columnName: string, filterValue: string, accessToken: string) {

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
    } else {
        alert(`Error applying filter: ${await filterResponse.text()}`);
    }
}



