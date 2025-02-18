"use strict";
// Authentication
//const accessToken = getAccessToken();
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
    return getTokenWithMSAL(clientId, redirectUri, msalConfig);
}
// Fetch OneDrive File by Path
async function fetchOneDriveFileByPath(filePathAndName, accessToken) {
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
// Update Word Document
async function invoice(issue = false) {
    accessToken = await getAccessToken() || '';
    (async function show() {
        if (issue)
            return;
        excelData = await fetchExcelTable(accessToken, excelPath, 'LivreJournal');
        if (!excelData)
            return;
        insertInvoiceForm(excelData);
    })();
    (async function issueInvoice() {
        if (!issue)
            return;
        const inputs = Array.from(document.getElementsByTagName('input'));
        const criteria = inputs.filter(input => Number(input.dataset.index) >= 0);
        (function fillInputs() {
            return;
            //!For testing only
            criteria[0].value = 'SARL MARTHA';
            criteria[1].value = 'Redressement Judiciaire';
            criteria[2].value = 'CARPA, Honoraire, Débours/Dépens, Provision/Règlement';
            criteria[3].value = '2015-01-01';
            criteria[4].value = '2025-01-01';
            inputs.filter(input => input.type === 'checkbox')[1].checked = true;
        })();
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
        };
        const contentControls = getContentControlsValues(invoice, date);
        const filePath = `${destinationFolder}/${newWordFileName(invoice.clientName, invoice.matters, invoice.number)}`;
        await createAndUploadXmlDocument(filtered, contentControls, accessToken, filePath, lang);
    })();
}
async function extractInvoiceData(lang) {
    const inputs = Array.from(document.getElementsByTagName('input'));
    const criteria = inputs.filter(input => input.dataset.index);
    //@ts-expect-error
    const tableData = filterRows(0, await fetchExcelData());
    return getRowsData(tableData, lang);
    function filterRows(i, tableData) {
        while (i < criteria.length) {
            const input = criteria[i];
            tableData = tableData.filter(row => row[Number(input.dataset.index)].toString() === input.value);
            i++;
        }
        return tableData;
    }
}
function dateFromExcel(excelDate) {
    const date = new Date((excelDate - 25569) * (60 * 60 * 24) * 1000); //This gives the days converted from milliseconds. 
    const dateOffset = date.getTimezoneOffset() * 60 * 1000; //Getting the difference in milleseconds
    return new Date(date.getTime() + dateOffset);
}
function getContentControlsValues(invoice, date) {
    const fields = {
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
function getMSGraphClient(accessToken) {
    //@ts-expect-error
    return MicrosoftGraph.Client.init({
        //@ts-expect-error
        authProvider: (done) => {
            done(null, accessToken);
        }
    });
}
// Save Word Document to Another Location
async function saveWordDocumentToNewLocation(invoice, accessToken, originalFilePath = templatePath, newFilePath = destinationFolder) {
    const date = new Date();
    const fileName = newWordFileName(invoice.clientName, invoice.matters, invoice.number);
    newFilePath = `${destinationFolder}/${fileName}`;
    try {
        const fileContent = await fetchOneDriveFileByPath(originalFilePath, accessToken);
        const resp = await getMSGraphClient(accessToken).api(`/me/drive/root:${newFilePath}:/content`).put(fileContent);
        console.log('put response = ', resp);
        return newFilePath;
    }
    catch (error) {
        console.error("Error saving Word document:", error);
    }
}
//# sourceMappingURL=pwaVersion.js.map