"use strict";
// Authentication
//const accessToken = getAccessToken();
const accessToken = getAccessToken();
const path = "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/";
const templatePath = path + 'FactureTEMPLATE [NE PAS MODIFIDER].dotm';
const destinationFolder = path + 'Clients';
const excelFilePath = "Legal/Mon Cabinet d'Avocat/Comptabilité/Comptabilité de Mon Cabinet_15 10 2023.xlsm";
function getAccessToken() {
    const clientId = "157dd297-447d-4592-b2d3-76b643b97132";
    const redirectUri = "https://mbibawi.github.io/ExcelInvoicingAddIn"; //!must be the same domain as the app
    const msalConfig = {
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
    return 'EwCIA8l6BAAUBKgm8k1UswUNwklmy2v7U/S+1fEAAXeDtX1LNbITceYdWIcNn8CsVuiQ5WtwZAjhF+ZLRK48+MhGO/clQlmj4HL1AJzCNIthvolYwVgD3ju+py/fhYw7tn93YGGcvGFb9jyFyZkTz06nVYDll84hoZgyZTvBmzoGiRiCVqVsS7qCuZcP0STI6KrhrJzNEVcAKdlUneM/0rlFSnTuyvQM3TH54shooZYwZqsZpSd0jmyaUqM17uE4M4A5DZ0XpeeymwwR7TdQJq4DjnCV+GgKpqoRKTbefmQVXTLGtOtNVdOwWoL1MaEXf7gkqckd9MydJLqNlkre5cqEmM51i1Rtkg5isz3ZUcaJSZer920ZAvRQdSlxIAcQZgAAECV4s2wimu9YxrmDRo4SQ/VQAqNYYePkZvNSQli/YqwoNbz0e7Mpi/Wv+5kjdsRvbxJgcli2GHhSsBs8R8ZTXNm15KuwGQuji0kJFcUhxz4dzXPHdR2P0OfCp92OtyXm8yjs8PJEzoTf2totObuVb3HnfRyUBd7QYLPFJBd1hcJP2FZes8R6J7UeXlQHqabJbVXka6yEwQ+rTUH2DZjihUo4EHk0Sgc0PR+U6kOKP5akk4fot9JGv8GDRNFjveKe2ec33pBiZ33ZlDa9XPIdDZHc6RPAHs+SzhJ/LJZYtjBf7hOgB+cb5fVgtwKkzT420q1oRXs8KS/mqc8MHZEN20ktuOev+TYUAsIQnpwa2mdAvOZ7z7C4hCOaJSNCYZefm4/6fn7bntwGz1JPizPNB2D5lKrGHpdF/CbcWjUoBVofxfoReD9NKR8xnTNISrVVvRi178Ra2ZuIpR6TukM3NhgFXh5jKtcgNcc+WOZ3tqnFdJHMWDJgEbOk55sqZ/6uckrVS2e282RFOw5Pq1tQ9i/CcbwmXEsn/K7lO5BgiDTko6+KO1U3deGz84btblRi9TUQrSGS8g18NnkLF1bpfJquGuY8Ph0p1Fn42DEoc4sAkVAgP6L5mSQZ5FGV6F4C/eoXxqwtSQFGbfbjkF69XvqgB/2XaQqZxblhhDGhGA02FuNgJg2uyyREhoEFswLXHVpbdddR1T/In/rqNLW9p9QzLpFC8IzhVLojGjVLjK22S07KLppxOV0b52j0g2jTObflUiPXoqk4x6JiuYzOpqn/Pwp8VUm94TE7LMbJfMenYYl7Ag==';
    return getTokenWithMSAL(clientId, redirectUri, msalConfig);
}
// Fetch OneDrive File by Path
async function fetchOneDriveFileByPath(filePathAndName) {
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
    }
    catch (error) {
        console.error("Error fetching OneDrive file:", error);
    }
}
// Fetch Excel Data
async function fetchExcelData(filePath = excelFilePath) {
    try {
        //@ts-expect-error
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(await fetchOneDriveFileByPath(filePath));
        const worksheet = workbook.getWorksheet(1);
        const data = worksheet.getSheetValues();
        return data;
    }
    catch (error) {
        console.error("Error fetching Excel data:", error);
    }
}
// Update Word Document
async function invoice() {
    const tableData = await extractInvoiceData();
    const lang = Array.from(document.getElementsByTagName('input')).find(input => input.type === 'checkbox' && input.checked === true)?.value;
    const invoiceDetails = {
        clientName: tableData.map(row => String(row[0]))[0] || 'CLIENT',
        matters: (await getUniqueValues(1, tableData)).map(el => String(el)),
        adress: (await getUniqueValues(15, tableData)).map(el => String(el)),
        lang: lang || 'FR'
    };
    const richText = getContentControlsValues(invoiceDetails);
    if (!tableData)
        return;
    try {
        const document = new Word.Document();
        document.load(await fetchOneDriveFileByPath(templatePath));
        // Update Table
        //@ts-expect-error
        const table = document.body.tables[0];
        tableData.forEach(rowData => {
            const row = table.addRow();
            rowData.forEach(cellData => {
                row.addCell(cellData);
            });
        });
        // Update Rich Text Content Controls
        const contentControls = document.contentControls;
        richText.forEach(([title, text]) => {
            const control = contentControls.getByTitle(title);
            //@ts-expect-error
            control.text = text;
        });
        await document.save();
    }
    catch (error) {
        console.error("Error updating Word document:", error);
    }
}
async function extractInvoiceData() {
    const inputs = Array.from(document.getElementsByTagName('input'));
    const criteria = inputs.filter(input => input.dataset.index);
    const tableData = filterRows(0, await fetchExcelData());
    return getData(tableData);
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
    function filterRows(i, tableData) {
        while (i < criteria.length) {
            const input = criteria[i];
            tableData = tableData.filter(row => row[Number(input.dataset.index)].toString() === input.value);
            i++;
        }
        return tableData;
    }
}
function getContentControlsValues(invoice) {
    const date = new Date();
    const fields = {
        dateLabel: {
            title: 'LabelParisLe',
            text: { FR: 'Paris le ', EN: 'Paris on ' }[invoice.lang] || '',
        },
        date: {
            title: 'RTInvoiceDate',
            text: [date.getDate(), date.getMonth() + 1, date.getFullYear()].join('/'),
        },
        numberLabel: {
            title: 'LabelInvoiceNumber',
            text: { FR: 'Facturen n° : ', EN: 'Invoice No.:' }[invoice.lang] || '',
        },
        number: {
            title: 'RTInvoiceNumber',
            text: [date.getDate(), date.getMonth() + 1, date.getFullYear() - 2000].join('') + '/' + [date.getHours(), date.getMinutes()].join(''),
        },
        subjectLable: {
            title: 'LabelSubject',
            text: { FR: 'Objet : ', EN: 'Subject: ' }[invoice.lang] || '',
        },
        subject: {
            title: 'RTMatter',
            text: invoice.matters.join(' & '),
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
            text: invoice.adress.join(' & '),
        },
    };
    return Object.keys(fields).map(key => [fields[key].title, fields[key].text]);
}
function getMSGraphClient() {
    //@ts-expect-error
    return MicrosoftGraph.Client.init({
        //@ts-expect-error
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}
// Save Word Document to Another Location
async function saveWordDocumentToNewLocation(originalFilePath, newFilePath = destinationFolder, invoice) {
    const date = new Date();
    const fileName = `Test_Facture_${invoice.clientName}_${Array.from(invoice.matters).join('&')}_${[date.getFullYear(), date.getMonth() + 1, date.getDate()].join('')}@${[date.getHours(), date.getMinutes()].join(':')}.docx`;
    newFilePath = `${destinationFolder} / ${fileName}`;
    try {
        const fileContent = await fetchOneDriveFileByPath(originalFilePath);
        await getMSGraphClient().api(`/me/drive/root:${newFilePath}:/content`).put(fileContent);
    }
    catch (error) {
        console.error("Error saving Word document:", error);
    }
}
//# sourceMappingURL=pwaVersion.js.map