/// <reference types="office-js" />

async function fetchExcelTable(accessToken: string, filePath: string, tableName: string): Promise<string[][]> {

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

    const excelData = await fetchExcelTable(accessToken, excelPath, tableName);

    if (!excelData) return;

    insertInvoiceForm(excelData);

    const inputs = Array.from(document.getElementsByTagName('input'));

    const criteria = inputs.filter(input => Number(input.dataset.index) >= 0);

    //!For testing only
    criteria[0].value = 'SCI SHAMS';
    criteria[1].value = 'Adjudication studio rue Théodore Deck';
    criteria[2].value = ['CARPA', 'Honoraire', 'Débours/Dépens', 'Provision/Règlement'].join(', ');
    criteria[3].value = '2015-01-01';
    criteria[4].value = '2025-01-01';

    inputs.filter(input => input.type === 'checkbox')[1].checked = true;

    const lang = inputs.find(input => input.dataset.language && input.checked === true)?.dataset.language || 'FR';
    console.log('language = ', lang)
    const date = new Date();
    const filtered = filterExcelData(excelData, criteria, lang);
    const invoice = {
        number: getInvoiceNumber(date),
        clientName: getInputValue(0, criteria),
        matters: getArray(getInputValue(1, criteria)),
        lang: lang,
        adress: Array.from(new Set(filtered.map(row => row[16])))
    }

    const path = "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/";
    const templatePath = path + 'FactureTEMPLATE [NE PAS MODIFIDER].dotm';
    const fileName: string = newWordFileName(invoice.clientName, invoice.matters, invoice.number);

    // Define content control replacements
    const contentControls = getContentControlsValues(invoice, date);

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

function insertInvoiceForm(excelTable: string[][]) {
    const form = document.getElementById('form');
    if (!form) return;
    form.innerHTML = '';
    const title = excelTable[0];
    const inputs = insertInputsAndLables([0, 1, 2, 3, 3]);//Inserting the fields inputs (Client, Matter, Nature, Date). We insert the date twice


    insertInputsAndLables(['Français', 'English'], true); //Inserting langauges checkboxes
    (function addBtn() {
        const btnIssue = document.createElement('button');
        btnIssue.innerText = 'Generate Invoice';
        btnIssue.classList.add('button');
        btnIssue.onclick = () => invoice(true);
        form.appendChild(btnIssue);
    })();

    function insertInputsAndLables(indexes: (number | string)[], checkBox: boolean = false): HTMLInputElement[] {
        const id = 'input';
        let css = 'field';
        if (checkBox) css = 'checkBox'
        return indexes.map(index => {
            const input = document.createElement('input');
            input.classList.add(css);
            if (checkBox) input.type = 'checkbox';
            else if (Number(index) < 3) input.type = 'text';
            else input.type = 'date';
            checkBox ? input.id = id : input.id = id + index.toString();
            if (!checkBox) {
                input.name = input.id;
                input.dataset.index = index.toString();
                input.setAttribute('list', input.id + 's');
                input.autocomplete = "on";
                if (Number(index) < 2)
                    input.onchange = () => inputOnChange(Number(input.dataset.index), excelTable.slice(1), true);
            }
            else if (checkBox) input.dataset.language = index.toString().slice(0, 2).toUpperCase();


            const label = document.createElement('label');
            checkBox ? label.innerText = index.toString() : label.innerText = title[Number(index)];
            label.htmlFor = input.id;

            form?.appendChild(label);
            form?.appendChild(input);
            if (Number(index) < 1) createDataList(input?.id, Array.from(new Set(TableRows.slice(1, -1).map(row => row[0]))));//We create a unique values dataList for the 'Client' input
            return input
        });
    };

}

function getInputByIndex(inputs: HTMLInputElement[], index: number) {
    return inputs.find(input => Number(input.dataset.index) === index)
}

function getIndex(input: HTMLInputElement) {
    return Number(input.dataset.index)
}

function filterExcelData(data: string[][], criteria: HTMLInputElement[], lang: string, i: number = 0) {

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
