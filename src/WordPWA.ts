/// <reference types="office-js" />

async function fetchExcelTable(accessToken: string, filePath: string, tableName = 'LivreJournal'): Promise<string[][]> {

    const fileUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook/tables/${tableName}/range`;

    const response = await fetch(fileUrl, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!response.ok) throw new Error("Failed to fetch Excel data");

    const data = await response.json();
    //@ts-ignore
    return data.values; // Returns data as string[][]
}

async function fetchWordTemplate(accessToken: string, filePath: string): Promise<Blob> {
    const fileUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/content`;

    const response = await fetch(fileUrl, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!response.ok) throw new Error("Failed to fetch Word template");

    return await response.blob(); // Returns the Word template as a Blob
}

async function saveWordDocument(accessToken: string, filePath: string, blob: Blob): Promise<void> {
    const fileUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/content`;

    const response = await fetch(fileUrl, {
        method: "PUT",
        headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        },
        body: blob
    });

    if (!response.ok) throw new Error("Failed to save Word document");
}

async function createDocumentFromTemplate(accessToken: string, templatePath: string, newPath: string, excelData: string[][], contentControlData: string[][]): Promise<void> {
    // Fetch the Word template
    const templateBlob = await fetchWordTemplate(accessToken, templatePath);

    // Load template into Word
    await Word.run(async (context) => {
        const doc = context.document;
        doc.body.insertFileFromBase64(await blobToBase64(templateBlob), Word.InsertLocation.replace);
        await context.sync();

        // Get the first table and add rows from Excel data
        const tables = doc.body.tables;
        tables.load("items");
        await context.sync();

        if (tables.items.length > 0) {
            const firstTable = tables.items[0];

            for (const row of excelData) {
                //@ts-expect-error
                firstTable.addRow(-1, row);
            }
        }

        // Update content controls by title
        const contentControls = doc.contentControls;
        contentControls.load("items, title");
        await context.sync();

        /*
        contentControls.items.forEach(([title, text]) => {
            if (title) {
                control.insertText(contentControlData[control.title], Word.InsertLocation.replace);
            }
        });
        */

        await context.sync();

        // Save the modified document
        //@ts-expect-error
        const base64Doc = doc.body.getBase64();
        await context.sync();

        // Convert base64 to Blob and save to OneDrive
        const finalBlob = base64ToBlob(await base64Doc.value);
        await saveWordDocument(accessToken, newPath, finalBlob);
    });
}

async function copyWordTemplate() {

}

// Utility function: Convert Blob to Base64
async function blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result!.toString().split(",")[1]);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

// Utility function: Convert Base64 to Blob
function base64ToBlob(base64: string): Blob {
    const byteCharacters = atob(base64);
    const byteNumbers = new Array(byteCharacters.length).fill(0).map((_, i) => byteCharacters.charCodeAt(i));
    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

// Usage Example
async function mainWithWordgraphApi() {
    const accessToken = await getAccessToken() || ''; // Ensure you obtain this via MSAL.js
    if (!accessToken) return

    const excelPath = "Legal/Mon Cabinet d'Avocat/Comptabilité/Comptabilité de Mon Cabinet_15 10 2023.xlsm";

    // Fetch Excel data

    const excelData = await fetchExcelTable(accessToken, excelPath, 'LivreJournal');

    if (!excelData) return;

    insertInvoiceForm(excelData, Array.from(new Set(excelData.map(row => row[0]))));

    const inputs = Array.from(document.getElementsByTagName('input'));

    const criteria = inputs.filter(input => Number(input.dataset.index) >= 0);

    //!For testing only
    criteria[0].value = 'SCI SHAMS';
    criteria[1].value = 'Adjudication studio rue Théodore Deck';
    criteria[2].value = ['CARPA', 'Honoraire', 'Débours/Dépens', 'Provision/Règlement'].join(', ');
    criteria[3].value = '2015-01-01';
    criteria[4].value = '2025-01-01';

    inputs.filter(input => input.type === 'checkbox')[1].checked = true;

    const lang = inputs.find(input => input.type === 'checkbox' && input.checked === true)?.dataset.language || 'FR';
    console.log('language = ', lang)

    const filtered = filterExcelData(excelData, criteria, lang);
    const invoice = {
        number: getInvoiceNumber(new Date()),
        clientName: getInputValue(0, criteria),
        matters: getArray(getInputValue(1, criteria)),
        lang: lang,
        adress: Array.from(new Set(filtered.map(row => row[16])))
    }

    const path = "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/";
    const templatePath = path + 'FactureTEMPLATE [NE PAS MODIFIDER].dotm';
    const fileName: string = newWordFileName(invoice.clientName, invoice.matters, invoice.number);

    // Define content control replacements
    const contentControls = getContentControlsValues(invoice);

    await editWordWithGraphApi(filtered, contentControls, templatePath, fileName, accessToken);
    return


    async function editWithAny() {
        // Generate Word document from template
        await createDocumentFromTemplate(accessToken, templatePath, `${path}Client/${fileName}`, excelData, contentControls);

    }

}

function getInputValue(index: number, inputs: HTMLInputElement[]) {
    return inputs.find(input => Number(input.dataset.index) === index)?.value || ''
}

function insertInvoiceForm(excelTable: string[][], clientUniqueValues: string[]) {
    const form = document.getElementById('form');
    if (!form) return;
    const title = excelTable[0];
    const inputs = insertInputsAndLables([0, 1, 2, 3, 3]);//Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice

    inputs.forEach(input => input?.addEventListener('focusout', async () => await inputOnChange(input), { passive: true }));

    insertInputsAndLables(['Français', 'English'], true); //Inserting langauges checkboxes
    form.innerHTML += `<button onclick="generateInvoice()"> Generate Invoice</button>`; //Inserting the button that generates the invoice

    function insertInputsAndLables(indexes: (number | string)[], checkBox: boolean = false): HTMLInputElement[] {
        const id = 'input';
        return indexes.map(index => {
            const input = document.createElement('input');
            if (checkBox) input.type = 'checkbox';
            else if (Number(index) < 3) input.type = 'text';
            else input.type = 'date';
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

            form?.appendChild(label);
            form?.appendChild(input);
            if (Number(index) === 0) createDataList(input?.id, clientUniqueValues);//We create a unique values dataList for the 'Client' input
            return input
        });
    };

    async function inputOnChange(input: HTMLInputElement, unfilter: boolean = false) {
        return console.log('filter table on input change was called')
        const index = Number(input.dataset.index);

        if (index === 0) unfilter = true;//If this is the 'Client' column, we remove any filter from the table;

        //We filter the table accordin to the input's value and return the visible cells
        const visibleCells = await filterTable(undefined, [{ column: index, value: getArray(input.value) }], unfilter);

        if (visibleCells.length < 1) return alert('There are no visible cells in the filtered table');

        //We create (or update) the unique values dataList for the next input 
        const nextInput = getNextInput(input);
        if (!nextInput) return;
        createDataList(nextInput?.id || '', await getUniqueValues(Number(nextInput.dataset.index), visibleCells));


        function getNextInput(input: HTMLInputElement) {
            let nextInput: Element | null = input.nextElementSibling;
            while (nextInput?.tagName !== 'INPUT' && nextInput?.nextElementSibling) {
                nextInput = nextInput.nextElementSibling
            };

            return nextInput as HTMLInputElement
        }

        if (index === 1) {
            //!Need to figuer out how to create a multiple choice input for nature
            const nature = new Set((await filterTable(undefined, undefined, false)).map(row => row[index]));
            //nature.forEach(el => form?.appendChild(createCheckBox(undefined, el)));
        }

    };

}

function filterExcelData(data: string[][], criteria: HTMLInputElement[], lang: string, i: number = 0) {
    while (i < 2) {
        //!We exclued the nature input (column = 2) because it is a string[] of values not only one value.
        data = data.filter(row => row[Number(criteria[i].dataset.index)] === criteria[i].value);
        i++
    }

    const nature = criteria[2].value.replaceAll(' ', '').split(',');
    data = data.filter(row => nature.includes(row[2]));

    data = filterByDate(data);

    return getData(data);

    function filterByDate(data: string[][]) {
        const [from, to] = getDateCriteria();

        if (Number(from) && Number(to))
            return data.filter(row => convertDate(row[3]) >= new Date(from).getTime() && convertDate(row[3]) <= new Date(to).getTime()); //we filter by the date
        else if (Number(from))
            return data.filter(row => convertDate(row[3]) >= new Date(from).getTime()); //we filter by the date
        else if (Number(to))
            return data.filter(row => convertDate(row[3]) <= new Date(to).getTime()); //we filter by the date
        else
            return data.filter(row => convertDate(row[3]) <= new Date().getTime()); //we filter by the date


        function getDateCriteria() {
            const [from, to] = criteria.filter(input => Number(input.dataset.index) === 3);
            return [new Date(from.value) || undefined, new Date(to.value) || undefined];
        }

        function convertDate(date: string){
            return dateFromExcel(Number(date)).getTime();
        }
    }

    function getData(filteredTable: string[][]): string[][] {
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
            totalHourlyBilled: {
                FR: 'Total du temps facturé au taux horaire (hors prestations au forfait)',
                EN: 'Total billable hours (other than lump-sum billed services)'
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
            decimalSign: { FR: ',', EN: '.' }[lang] || '.',
        }
        const amount = 9, vat = 10, hours = 7, rate = 8, nature = 2, descr = 14;

        const data: string[][] = filteredTable.map(row => {
            const date = dateFromExcel(Number(row[3]));
            const time = getTimeSpent(Number(row[hours]));

            let description = `${String(row[nature])} : ${String(row[descr])}`;//Column Nature + Column Description;

            //If the billable hours are > 0
            if (time)
                //@ts-ignore
                description += `(${lables.hourlyBilled[lang]} ${time} ${lables.hourlyRate[lang]} ${Math.abs(row[rate]).toString()} €)`;


            const rowValues: string[] = [
                [date.getDate(), date.getMonth() + 1, date.getFullYear()].join('/'),//Column Date
                description,
                getAmountString(Number(row[amount]) * -1), //Column "Amount": we inverse the +/- sign for all the values 
                getAmountString(Math.abs(Number(row[vat]))), //Column VAT: always a positive value
            ];
            return rowValues;
        });

        pushTotalsRows();
        return data

        function getAmountString(value: number): string {
            //@ts-expect
            return value.toFixed(2).replace('.', lables.decimalSign) + ' €' || ''
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
                pushSumRow(lables.totalFees, totalFee, totalFeeVAT)
            if (totalExpenses > 0)
                pushSumRow(lables.totalExpenses, totalExpenses, totalExpensesVAT);
            if (totalPayments > 0)
                pushSumRow(lables.totalPayments, totalPayments, totalPaymentsVAT);
            if (totalTimeSpent > 0)
                pushSumRow(lables.totalHourlyBilled, totalTimeSpent)//!We don't pass the vat argument in order to get the corresponding cell of the Word table empty

            pushSumRow(lables.totalDue, totalDue, totalDueVAT);

            function pushSumRow(label: { FR: string, EN: string }, amount: number, vat?: number) {
                if (!amount) return;
                amount = Math.abs(amount);
                data.push(
                    [
                        //@ts-ignore
                        label[lang],
                        '',
                        label === lables.totalHourlyBilled ? getTimeSpent(amount) || '' : getAmountString(amount) || '',//The total amount can be a negative number, that's why we use Math.abs() in order to get the absolute number without the negative sign
                        //@ts-ignore
                        Number(vat) >= 0 ? getAmountString(Math.abs(vat)) : '' //!We must check not only that vat is a number, but that it is >=0 in order to avoid getting '' each time the vat is = 0, because we need to show 0 vat values
                    ]);
            }


            function getTotals(index: number, nature: string | null) {
                const total =
                    filteredTable.filter(row => nature ? row[2] === nature : row[2] === row[2])
                        .map(row => Number(row[index]));
                let sum = 0;
                for (let i = 0; i < total.length; i++) {
                    sum += total[i]
                }
                if (index === 7)
                    console.log('this is the hourly rate') //!need to something to adjust the time spent format
                return sum;
            }

        }

        function getTimeSpent(time: number) {
            if (!time || time <= 0) return undefined;
            time = time * (60 * 60 * 24)//84600 is the number in seconds per day. Excel stores the time as fraction number of days like "1.5" which is = 36 hours 0 minutes 0 seconds;
            const minutes = Math.floor(time / 60);
            const hours = Math.floor(minutes / 60);
            return [hours, minutes % 60, 0]
                .map(el => el.toString().padStart(2, '0'))
                .join(':');
        }

        function dateFromExcel(excelDate: number) {
            const date = new Date((excelDate - 25569) * (60 * 60 * 24) * 1000);//This gives the days converted from milliseconds. 
            const dateOffset = date.getTimezoneOffset() * 60 * 1000;//Getting the difference in milleseconds
            return new Date(date.getTime() + dateOffset);
        }

    }
}

//main().catch(console.error);

async function editWordWithGraphApi(excelData: string[][], contentControlData: string[][], templatePath: string, fileName: string, accessToken: string) {
    // Function to authenticate and get access token

    const fileData = await copyTemplate(accessToken, templatePath, fileName);
    if (!fileData) return;
    const fileId: string = fileData.id;
    await addRowsToTable(fileId, excelData);
    await updateContentControls(fileId, contentControlData);

    console.log('Document creation and updates completed successfully');

    // Function to copy a Word template to a new location
    async function copyTemplate(accessToken: string, templatePath: string, fileName: string) {

        const fileData = await getFileDataByPath(accessToken, templatePath);

        if (!fileData || !fileData.id) return;

        const copyTo = `https://graph.microsoft.com/v1.0/me/drive/items/${fileData.id}/copy`;

        const response = await fetch(copyTo, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                parentReference: {
                    driveId: fileData.parentReference.driveId,
                    id: fileData.id
                },
                name: fileName,
            }),
        });

        if (!response.ok) {
            throw new Error(`Failed to copy template: ${response.statusText}`);
        }

        // Wait for the copy operation to complete
        const location = response.headers.get('Location');
        if (!location) {
            throw new Error('Copy operation did not return a location header');
        }

        // Poll the status URL until the copy operation is complete
        let statusResponse;
        do {
            statusResponse = await fetch(location, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                },
            });
            if (!statusResponse.ok) {
                throw new Error(`Failed to check copy status: ${statusResponse.statusText}`);
            }
            const status = await statusResponse.json();
            if (status.status === 'completed') {
                return await response.json(); // Return the path of the new file
            }
            await new Promise((resolve) => setTimeout(resolve, 1000)); // Wait 1 second before polling again
        } while (true);
    }

    async function getFileDataByPath(accessToken: string, filePath: string) {
        const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(filePath)}`;

        const fileResponse = await fetch(endpoint, {
            method: 'GET',
            headers: { Authorization: `Bearer ${accessToken}` },
        });

        return await fileResponse.json();
    }

    // Function to add rows to the first table in a Word document
    async function addRowsToTable(fileId: string, newRows: string[][]) {
        // JSON patch to add rows to the first table
        const patchData = newRows.map((row) => ({
            op: 'add',
            path: `/tables/0/rows/-`, // The "0" refers to the first table in the document
            value: row,
        }));

        if (await patch(patchData, fileId) === true)
            console.log('Rows added successfully');
    }

    // Function to update content controls by their titles
    async function updateContentControls(filePath: string, contentControls: any[][]) {

        // JSON patch to update content controls
        const patchData = contentControls.map(([title, text]) => ({
            op: 'replace',
            path: `/contentControls[title='${title}']/text`,
            value: text,
        }));

        if (await patch(patchData, filePath) === true)
            console.log('Content controls updated successfully');
    }

    async function patch(patchData: { op: string; path: string; value: string | string[] }[], fileId: string) {
        const endpoint = `https://graph.microsoft.com/v1.0/me/drive/root:/${fileId}:/content`;

        const response = await fetch(endpoint, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(patchData),
        });

        if (!response.ok)
            throw new Error(`Failed to update content controls: ${response.statusText}`);
        else return true;
    }
}
