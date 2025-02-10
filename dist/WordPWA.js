"use strict";
/// <reference types="office-js" />
async function fetchExcelTable(accessToken, filePath, tableName = 'LivreJournal') {
    if (!)
        const fileUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/workbook:/workbook/tables/${tableName}/rows`;
    const response = await fetch(fileUrl, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!response.ok)
        throw new Error("Failed to fetch Excel data");
    const data = await response.json();
    return data.values; // Returns data as string[][]
}
async function fetchWordTemplate(accessToken, filePath) {
    const fileUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}:/content`;
    const response = await fetch(fileUrl, {
        method: "GET",
        headers: { Authorization: `Bearer ${accessToken}` }
    });
    if (!response.ok)
        throw new Error("Failed to fetch Word template");
    return await response.blob(); // Returns the Word template as a Blob
}
async function saveWordDocument(accessToken, filePath, fileName, blob) {
    const fileUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${filePath}/${fileName}:/content`;
    const response = await fetch(fileUrl, {
        method: "PUT",
        headers: {
            Authorization: `Bearer ${accessToken}`,
            "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        },
        body: blob
    });
    if (!response.ok)
        throw new Error("Failed to save Word document");
}
async function createDocumentFromTemplate(accessToken, templatePath, newPath, excelData, contentControlData) {
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
                firstTable.addRow(-1, row);
            }
        }
        // Update content controls by title
        const contentControls = doc.contentControls;
        contentControls.load("items, title");
        await context.sync();
        contentControls.items.forEach((control) => {
            if (contentControlData[control.title]) {
                control.insertText(contentControlData[control.title], Word.InsertLocation.replace);
            }
        });
        await context.sync();
        // Save the modified document
        const base64Doc = doc.body.getBase64();
        await context.sync();
        // Convert base64 to Blob and save to OneDrive
        const finalBlob = base64ToBlob(await base64Doc.value);
        await saveWordDocument(accessToken, newPath, newName, finalBlob);
    });
}
// Utility function: Convert Blob to Base64
async function blobToBase64(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result.toString().split(",")[1]);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}
// Utility function: Convert Base64 to Blob
function base64ToBlob(base64) {
    const byteCharacters = atob(base64);
    const byteNumbers = new Array(byteCharacters.length).fill(0).map((_, i) => byteCharacters.charCodeAt(i));
    const byteArray = new Uint8Array(byteNumbers);
    return new Blob([byteArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}
// Usage Example
async function main() {
    const accessToken = await getAccessToken() || ''; // Ensure you obtain this via MSAL.js
    const excelPath = "Legal/Mon Cabinet d'Avocat/Comptabilité/Comptabilité de Mon Cabinet_15 10 2023.xlsm";
    // Fetch Excel data
    const excelData = filterExcelData(await fetchExcelTable(accessToken, excelPath, 'LivreJournal'));
    let inputs = Array.from(document.getElementsByTagName('input'));
    const criteria = inputs.filter(input => Number(input.dataset.index));
    const lang = inputs.find(input => input.type === 'checkbox' && input.checked === true)?.dataset.language || 'FR';
    const invoice = { clientName: getInputValue(0), matters: getArray(getInputValue(1)), lang: lang, adress: excelData.map(row => row[15]) };
    const path = "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/";
    const templatePath = path + 'FactureTEMPLATE [NE PAS MODIFIDER].dotm';
    const newPath = path + 'Clients/' + newWordFileName(new Date(), invoice.clientName, invoice.matters);
    // Define content control replacements
    const contentControls = getContentControlsValues(invoice);
    function getInputValue(index) {
        return inputs.find(input => Number(input.dataset.index) === index)?.value || '';
    }
    // Generate Word document from template
    await createDocumentFromTemplate(accessToken, templatePath, newPath, excelData, contentControls);
    function filterExcelData(data, i = 0) {
        while (i < criteria.length) {
            data = data.filter(row => row[Number(criteria[i].dataset.index)] === criteria[i].value);
            i++;
        }
        return getData(data);
        function getData(tableData) {
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
                totalTimeSpent: {
                    FR: 'Total des heures facturables (hors prestations facturées au forfait) ',
                    EN: 'Total billable hours (other than lump-sum billed services)'
                },
                decimal: {
                    nature: '',
                    FR: ',',
                    EN: '.'
                },
            };
            const amount = 9, vat = 10, hours = 7, rate = 8, nature = 2, descr = 14;
            const data = tableData.map(row => {
                const date = dateFromExcel(Number(row[3]));
                const time = getTimeSpent(Number(row[hours]));
                let description = `${String(row[nature])} : ${String(row[descr])}`; //Column Nature + Column Description;
                //If the billable hours are > 0
                if (time)
                    //@ts-ignore
                    description += `(${lables.hourlyBilled[lang]} ${time} ${lables.hourlyRate[lang]} ${Math.abs(row[rate]).toString()} €)`;
                const rowValues = [
                    [date.getDate(), date.getMonth() + 1, date.getFullYear()].join('/'), //Column Date
                    description,
                    getAmountString(Number(row[amount]) * -1), //Column "Amount": we inverse the +/- sign for all the values 
                    getAmountString(Math.abs(Number(row[vat]))), //Column VAT: always a positive value
                ];
                return rowValues;
            });
            pushTotalsRows();
            return data;
            function getAmountString(value) {
                //@ts-ignore
                return value.toFixed(2).replace('.', lables.decimal[lang] || '.') + ' €' || '';
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
                    pushSumRow(lables.totalFees, totalFee, totalFeeVAT);
                if (totalExpenses > 0)
                    pushSumRow(lables.totalExpenses, totalExpenses, totalExpensesVAT);
                if (totalPayments > 0)
                    pushSumRow(lables.totalPayments, totalPayments, totalPaymentsVAT);
                if (totalTimeSpent > 0)
                    pushSumRow(lables.totalTimeSpent, totalTimeSpent); //!We don't pass the vat argument in order to get the corresponding cell of the Word table empty
                pushSumRow(lables.totalDue, totalDue, totalDueVAT);
                function pushSumRow(label, amount, vat) {
                    if (!amount)
                        return;
                    amount = Math.abs(amount);
                    data.push([
                        //@ts-ignore
                        label[lang],
                        '',
                        label === lables.totalTimeSpent ? getTimeSpent(amount) || '' : getAmountString(amount) || '', //The total amount can be a negative number, that's why we use Math.abs() in order to get the absolute number without the negative sign
                        //@ts-ignore
                        Number(vat) >= 0 ? getAmountString(Math.abs(vat)) : '' //!We must check not only that vat is a number, but that it is >=0 in order to avoid getting '' each time the vat is = 0, because we need to show 0 vat values
                    ]);
                }
                function getTotals(index, nature) {
                    const total = tableData.filter(row => nature ? row[2] === nature : row[2] === row[2])
                        .map(row => Number(row[index]));
                    let sum = 0;
                    for (let i = 0; i < total.length; i++) {
                        sum += total[i];
                    }
                    if (index === 7)
                        console.log('this is the hourly rate'); //!need to something to adjust the time spent format
                    return sum;
                }
            }
            function getTimeSpent(time) {
                if (!time || time <= 0)
                    return undefined;
                time = time * (60 * 60 * 24); //84600 is the number in seconds per day. Excel stores the time as fraction number of days like "1.5" which is = 36 hours 0 minutes 0 seconds;
                const minutes = Math.floor(time / 60);
                const hours = Math.floor(minutes / 60);
                return [hours, minutes % 60, 0]
                    .map(el => el.toString().padStart(2, '0'))
                    .join(':');
            }
            function dateFromExcel(excelDate) {
                const date = new Date((excelDate - 25569) * (60 * 60 * 24) * 1000); //This gives the days converted from milliseconds. 
                const dateOffset = date.getTimezoneOffset() * 60 * 1000; //Getting the difference in milleseconds
                return new Date(date.getTime() + dateOffset);
            }
        }
    }
}
main().catch(console.error);
//# sourceMappingURL=WordPWA.js.map