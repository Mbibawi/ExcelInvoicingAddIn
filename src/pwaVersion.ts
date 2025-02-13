// Authentication
//const accessToken = getAccessToken();

const accessToken = getAccessToken();


const path = "Legal/Mon Cabinet d'Avocat/Comptabilité/Factures/"
const templatePath = path + 'FactureTEMPLATE [NE PAS MODIFIDER].dotm';
const destinationFolder = path + 'Clients';
const excelFilePath = "Legal/Mon Cabinet d'Avocat/Comptabilité/Comptabilité de Mon Cabinet_15 10 2023.xlsm"


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
async function fetchOneDriveFileByPath(filePathAndName: string) {
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

// Fetch Excel Data
async function fetchExcelData(filePath: string = excelFilePath) {
    try {
        //@ts-expect-error
        const workbook = new ExcelJS.Workbook();
        const data = await fetchOneDriveFileByPath(filePath);
        if (!data) return;
        
        await workbook.xlsx.load(data);

    
        const worksheet = workbook.getWorksheet(1);
        //const data = worksheet.getSheetValues();

        return data;
    } catch (error) {
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
    const richText: string[][] = getContentControlsValues(invoiceDetails);

    if (!tableData) return;
    try {
        const document = new Word.Document();
        //@ts-expect-error
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
    } catch (error) {
        console.error("Error updating Word document:", error);
    }
}

async function extractInvoiceData() {
    const inputs = Array.from(document.getElementsByTagName('input'));
    const criteria = inputs.filter(input => input.dataset.index);

    //@ts-expect-error
    const tableData = filterRows(0, await fetchExcelData());


    return getData(tableData);

    function getData(tableData: string[][]) {

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
        }
        const amount = 9, vat = 10, hours = 7, rate = 8, nature = 2, descr = 14;

        const data: string[][] = tableData.map(row => {
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
            //@ts-ignore
            return value.toFixed(2).replace('.', lables.decimal[lang] || '.') + ' €' || ''
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
                pushSumRow(lables.totalTimeSpent, totalTimeSpent)//!We don't pass the vat argument in order to get the corresponding cell of the Word table empty

            pushSumRow(lables.totalDue, totalDue, totalDueVAT);

            function pushSumRow(label: { FR: string, EN: string }, amount: number, vat?: number) {
                if (!amount) return;
                amount = Math.abs(amount);
                data.push(
                    [
                        //@ts-ignore
                        label[lang],
                        '',
                        label === lables.totalTimeSpent ? getTimeSpent(amount) || '' : getAmountString(amount) || '',//The total amount can be a negative number, that's why we use Math.abs() in order to get the absolute number without the negative sign
                        //@ts-ignore
                        Number(vat) >= 0 ? getAmountString(Math.abs(vat)) : '' //!We must check not only that vat is a number, but that it is >=0 in order to avoid getting '' each time the vat is = 0, because we need to show 0 vat values
                    ]);
            }


            function getTotals(index: number, nature: string | null) {
                const total =
                    tableData.filter(row => nature ? row[2] === nature : row[2] === row[2])
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
    function filterRows(i: number, tableData: string[][]) {
        while (i < criteria.length) {
            const input = criteria[i];
            tableData = tableData.filter(row => row[Number(input.dataset.index)].toString() === input.value);
            i++
        }
        return tableData
    }


}

function getContentControlsValues(invoice: { clientName: string, matters: string[], adress: string[], lang: string }): string[][] {
    const date = new Date();
    const fields: { [index: string]: { title: string, text: string } } = {
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
async function saveWordDocumentToNewLocation(originalFilePath: string, newFilePath: string = destinationFolder, invoice: { clientName: string, matters: string }) {
    const date = new Date();
    const fileName = `Test_Facture_${invoice.clientName}_${Array.from(invoice.matters).join('&')}_${[date.getFullYear(), date.getMonth() + 1, date.getDate()].join('')}@${[date.getHours(), date.getMinutes()].join(':')}.docx`;
    newFilePath = `${destinationFolder} / ${fileName}`
    try {

        const fileContent = await fetchOneDriveFileByPath(originalFilePath);
        await getMSGraphClient().api(`/me/drive/root:${newFilePath}:/content`).put(fileContent);
    } catch (error) {
        console.error("Error saving Word document:", error);
    }
}


