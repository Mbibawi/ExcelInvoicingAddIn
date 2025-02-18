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
// Fetch OneDrive File by Path
async function fetchOneDriveFileByPath(filePathAndName: string, accessToken: string) {
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
    } catch (error) {
        console.error("Error fetching OneDrive file:", error);
    }
}

async function addNewEntry(add: boolean = false) {
    accessToken = await getAccessToken() || '';

    (async function show() {
        if (add) return;
        excelData = await fetchExcelTable(accessToken, excelPath, 'LivreJournal');

        if (!excelData) return;

        showForm(excelData[0]);
    })();

    (async function addEntry() {
        if (!add) return;
        const inputs = Array.from(document.getElementsByTagName('input')) as HTMLInputElement[];//all inputs

        const row = inputs.map(input => {
            const index = getIndex(input);

            if (index === 3)
                return getISODate(input.valueAsDate);//The date
            else if (index === 4)
                return getISODate(getInputByIndex(inputs, 3)?.valueAsDate);//the Year - we return the full date of the date input
            else if ([5, 6].includes(index))
                return getTime([input]) || 0;//time start and time end
            else if (index === 7)
                return getTime([getInputByIndex(inputs, 5), getInputByIndex(inputs, 6)], true);//Total time
            else if ([8, 9, 10].includes(index))
                return input.valueAsNumber;//Hourly Rate, Amount, VAT
            else if (index === 15)
                return ''//'Link to a file' column
            else return input.value;
        });
        
        await addRowToExcelTable([row], excelData.length - 1, excelFilePath, 'LivreJournal', accessToken);

        function getISODate(date: Date | undefined | null) {
            if (!date) return;
            return [date.getFullYear(), date.getMonth() + 1, date.getDate()]
                .map(d => d.toString().padStart(2, '0'))
                .join('-');//This returns the date in the iso format : "YYYY-mm-dd"
            //return date?.toISOString().split('T')[0]
        }

        function getTime(inputs: (HTMLInputElement | undefined)[], total: boolean = false) {
            const day = (1000 * 60 * 60 * 24);
            if (!total && inputs[0]) return inputs[0].valueAsNumber / day;

            const from = inputs[0]?.valueAsNumber;
            const to = inputs[1]?.valueAsNumber;

            if (!from || !to) return;

            return (to - from) / day;
        }

    })()


    function showForm(title: string[]) {
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
            input.classList.add(css);
            input.id = id;
            input.name = id;
            input.autocomplete = "on";
            input.dataset.index = index.toString();
            input.type = 'text';
            if ([8, 9, 10].includes(index))
                input.type = 'number';
            else if (index === 3)
                input.type = 'date';
            else if ([5, 6].includes(index))
                input.type = 'time';
            else if ([4, 7, 15].includes(index)) input.style.display = 'none';//We hide those 3 columns: 'Total Time' and the 'Year' and 'Link to a File'
            else if ([0, 1, 2, 11, 12, 13, 16].includes(index)) {
                //We add a dataList for those fields
                input.setAttribute('list', input.id + 's');
                input.onchange = () => inputOnChange(index, excelData.slice(1, -1), false);
                if (![1,16].includes(index))
                    createDataList(input.id, getUniqueValues(index, excelData.slice(1, -1)));//We don't create the data list for columns 'Matter' (1) and 'Adress' (16) because it will be created when the 'Client' field is updated
            }

            return input
        }
    }
}

// Update Word Document
async function invoice(issue: boolean = false) {
    accessToken = await getAccessToken() || '';

    (async function show() {
        if (issue) return;

        excelData = await fetchExcelTable(accessToken, excelPath, 'LivreJournal');

        if (!excelData) return;

        insertInvoiceForm(excelData);

    })();

    (async function issueInvoice() {
        if (!issue) return;
        const inputs = Array.from(document.getElementsByTagName('input'));

        const criteria = inputs.filter(input => Number(input.dataset.index) >= 0);

        const lang = inputs.find(input => input.type === 'checkbox' && input.checked === true)?.dataset.language || 'FR';

        const filtered = filterExcelData(excelData, criteria, lang);
        console.log('filtered table = ', filtered);

        const date = new Date();

        const invoice = {
            number: getInvoiceNumber(date),
            clientName: getInputValue(0, criteria),
            matters: getArray(getInputValue(1, criteria)),
            adress: Array.from(new Set(filtered.map(row => row[16]))),
            lang: lang
        }
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
function inputOnChange(index: number, table: any[][], invoice: boolean) {
    let inputs = Array.from(document.getElementsByTagName('input') as HTMLCollectionOf<HTMLInputElement>);

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

    boundInputs.map(input => createDataList(input?.id, getUniqueValues(getIndex(input), filtered), invoice));

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

async function extractInvoiceData(lang: string) {
    const inputs = Array.from(document.getElementsByTagName('input'));
    const criteria = inputs.filter(input => input.dataset.index);

    //@ts-expect-error
    const tableData = filterRows(0, await fetchExcelData());


    return getRowsData(tableData, lang);


    function filterRows(i: number, tableData: string[][]) {
        while (i < criteria.length) {
            const input = criteria[i];
            tableData = tableData.filter(row => row[Number(input.dataset.index)].toString() === input.value);
            i++
        }
        return tableData
    }


}

function dateFromExcel(excelDate: number) {
    const date = new Date((excelDate - 25569) * (60 * 60 * 24) * 1000);//This gives the days converted from milliseconds. 
    const dateOffset = date.getTimezoneOffset() * 60 * 1000;//Getting the difference in milleseconds
    return new Date(date.getTime() + dateOffset);
}

function getContentControlsValues(invoice: { number: string, clientName: string, matters: string[], adress: string[], lang: string }, date: Date): string[][] {
    const fields: { [index: string]: { title: string, text: string } } = {
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

function getMSGraphClient(accessToken: string) {
    //@ts-expect-error
    return MicrosoftGraph.Client.init({
        //@ts-expect-error
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}
// Save Word Document to Another Location
async function saveWordDocumentToNewLocation(invoice: { number: string, clientName: string, matters: string[] }, accessToken: string, originalFilePath: string = templatePath, newFilePath: string = destinationFolder) {
    const date = new Date();
    const fileName = newWordFileName(invoice.clientName, invoice.matters, invoice.number);
    newFilePath = `${destinationFolder}/${fileName}`
    try {

        const fileContent = await fetchOneDriveFileByPath(originalFilePath, accessToken);
        const resp = await getMSGraphClient(accessToken).api(`/me/drive/root:${newFilePath}:/content`).put(fileContent);
        console.log('put response = ', resp);
        return newFilePath

    } catch (error) {
        console.error("Error saving Word document:", error);
    }
}

function getNewExcelRow(inputs: HTMLInputElement[]) {
    return inputs.map(input => {
        input.value

    })

}

async function addRowToExcelTable(row: any[][], index: number, filePath: string, tableName: string = 'LivreJournal', accessToken: string) {
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/rows/add`;

    const body = {
        index: index, // Example row
        values: row, // Example row
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
        console.error("Error adding row:", await response.text());
    }
}


