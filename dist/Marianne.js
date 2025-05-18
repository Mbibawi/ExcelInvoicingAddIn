"use strict";
const workbookName = '';
const tablename = '';
const Columns = {
    ReceiptNumber: 0,
    ArrivalDate: 1,
    RegisterNumber: 2,
    FileNumber: 3,
    CaseNumber: 4,
    CaseYear: 5,
    JudiciaryCaseYear: 6,
    CaseCourt: 7,
    CaseType: 8,
    ClaimantName: 9,
    DefendantName: 10,
    AssignmentDate: 11,
    ReceptionDate: 12,
    FirstMeetingDate: 13,
    CurrrentStatus: 14,
    LastMeetingDate: 15,
    EndOfInstructionDate: 16,
    AchievementMonth: 17,
    ProductionType: 18,
    returnedExpertName: 19,
    returnedExpertRegistration: 20,
    returnedReportDate: 21,
    returnedProductionType: 22,
    PartyName: 23,
    PartyAddress: 24,
    Observations: 25,
};
const forms = {
    addRecord: {
        label: 'أضف دعوى جديدة'
    },
    modifyRecord: {
        label: 'عدل بيانات دعوى موجودة أو أضف طرف جديد'
    },
    deleteRecord: {
        label: 'احذف بيان'
    },
    notices: {
        label: 'اصدار الإخطارات والتقارير',
        table: 'REPnotice',
        template: 'Noticies'
    },
    monthly: {
        label: 'اصدار تقرير الإنجاز الشهري',
        table: 'REPmonth',
        template: 'Monthly',
    },
    followUp: {
        label: 'اصدار تقرير المتابعة',
        table: 'REPFollowup',
        template: 'FollowUp'
    },
    returned: {
        label: 'اصدار تقرير القضايا المرتدة',
        table: 'REPreturned',
        template: 'Returned',
    },
    receptionDate: {
        label: 'اصدار تقرر القضايا المتداولة بناء على تاريخ الاستلام',
        table: 'REPreturned',
        template: 'PendingCases'
    },
    arrivalDate: {
        label: 'اصدار تقرير القضايا المتداولة بناء على تاريخ الورود للمكتب',
        table: 'REPreturned',
        template: 'PendingCases'
    },
};
const optionsLists = {
    status: {
        postponed: 'لم يبدأ',
        ongoing: 'جاري',
        issued: 'منتهي وصدر',
        notIssued: 'منتهي ولم يصدر',
    },
    achivementType: {
        report: 'تقرير',
        memorandum: 'مذكرة'
    },
    caseType: {
        civilAppealed: 'مدني مستأنف',
        administratif: 'قضاء إداري',
        civil: 'مدني',
        labourAppealed: 'عمال متسأنف',
        labour: 'عمال',
        tax: 'ضرائب',
        family: 'أحوال شخصية',
        criminal: 'جنح ونيابات',
        publicFunds: 'أموال عامة',
    }
};
async function startReporting() {
    const workbookPath = localStorage.reportingPath || prompt('Please provide the path for the workbook');
    if (!localStorage.reportingPath)
        localStorage.reportingPath = workbookPath;
    const tableName = localStorage.reportingTable || prompt('Please provide the name of the table');
    if (!localStorage.reportingTable)
        localStorage.reportingTable = tableName;
    const accessToken = await getAccessToken() || '';
    const sessionId = await createFileCession(workbookPath, accessToken, true);
    let table = await fetchExcelTableWithGraphAPI(sessionId, accessToken, workbookPath, tableName);
    const titles = table[0];
    const form = document.getElementById('form');
    const btnsContainer = document.getElementById('btns');
    const btns = {
        report: 'اصدار التقارير',
        crud: 'إضافة أو حذف أو تعديل البيانات',
    };
    initalBtns();
    function initalBtns() {
        Object.values(btns).forEach((lable) => {
            const btn = createButton(lable);
            btn.onclick = () => showOperationsBtns(lable);
        });
    }
    function showOperationsBtns(label) {
        btnsContainer.innerHTML = "";
        (function showReportingBtns() {
            if (label !== btns.report)
                return;
            const reporting = [
                forms.monthly,
                forms.followUp,
                forms.returned,
                forms.receptionDate,
                forms.arrivalDate
            ];
            const container = document.createElement('div');
            form.appendChild(container);
            reporting.map(form => {
                const button = createButton(form.label);
                button.onclick = () => issueReport(form);
            });
        })();
        (function showCRUDBtns() {
            if (label !== btns.crud)
                return;
            const crud = [forms.addRecord, forms.deleteRecord, forms.modifyRecord];
            crud.map(form => {
                const button = createButton(form.label);
                button.onclick = () => showCRUDForm(form);
            });
        })();
        (function insertBackBtn() {
            const btn = createButton('العودة للقائمة السابقة');
            btn.onclick = initalBtns;
        })();
    }
    ;
    function createButton(label) {
        const btn = document.createElement('button');
        btn.classList.add('button');
        btn.innerText = label;
        btnsContainer.appendChild(btn);
        return btn;
    }
    function issueReport(form) {
        const templatePath = `baseUrl${form.template}`;
        (async function monthlyReport() {
            if (form !== forms.monthly)
                return;
            const [criteria1, criteria2] = await getDateCriterion();
            if (!criteria1 || !criteria2)
                return alert('We could not build the dates for filtering the table');
            await clearFilterExcelTableGraphAPI(workbookPath, tableName, sessionId, accessToken);
            await filterExcelTableWithGraphAPI(workbookPath, tableName, titles[1], [criteria1, criteria2], sessionId, accessToken, false);
            const visibleCells = await getVisibleCellsWithGraphAPI(workbookPath, tableName, sessionId, accessToken);
            const rows = visibleCells.slice(1);
            await createAndUploadXmlDocument(accessToken, templatePath, workbookPath, 'AR', form.table, rows);
            async function getDateCriterion() {
                const date = await promptForUserInput('date', 'اختر أي يوم في خلال الشهر الذي تريد أن يشمله التقرير. التقرير سوف يشمل كافة القضايا من يوم واحد في الشهر حتى آخر الشهر.');
                if (!date)
                    return;
                const [year, month, day] = date.split('-');
                const nextMonth = (Number(month) + 1).toString().padStart(2, '0');
                return [
                    `>=${year}-${month}-01`,
                    `<=${year}-${nextMonth}-01`
                ];
            }
        })();
        (async function followUpReport() {
            if (form !== forms.followUp)
                return;
            const values = [optionsLists.status.ongoing];
            //pendingCasesReport(Columns.CurrrentStatus, form)
        })();
        (async function pendingByDateReceivedReport() {
            if (form !== forms.receptionDate)
                return;
            await pendingCasesReport([Columns.ReceptionDate, Columns.JudiciaryCaseYear, Columns.CaseYear, Columns.CaseNumber], form);
        })();
        (async function pendingByDateArrivedReport() {
            if (form !== forms.arrivalDate)
                return;
            await pendingCasesReport([Columns.ArrivalDate, Columns.ReceptionDate, Columns.ReceiptNumber, Columns.RegisterNumber], form);
        })();
        (async function returnedReport() {
            if (form !== forms.returned)
                return;
        })();
    }
    /**
     * This function issues a report for all the ongoing cases. This reports is issued in 2 versions: in 1 versions, the cases are sorted by the date they were assigned to the expert, and in the other version, they are sorted by the date the case was assigned to the bureau
     * @param {string[]} values - This is the list ofe "current status" values based on which the Excel table will be filtered, these values should exclude the "finished and issued report" case
     * @param {number} column - This is index of the column based on which the Excel table will be sorted: either this is the "date of assignment to the Expert" (column AssignmentDate) or the "date of arrival to the burreau" (column ArrivalDate)
     */
    async function pendingCasesReport(sorting, form) {
        const values = Object.values(optionsLists.status)
            .filter(value => value !== optionsLists.status.issued);
        const until = await promptForUserInput('date', 'حدد حتى أي تاريخ تريد للتقرير أن يصدر.');
        if (!until)
            return alert('You must choose a valid date');
        await clearFilterExcelTableGraphAPI(workbookPath, tableName, sessionId, accessToken);
        await filterExcelTableWithGraphAPI(workbookPath, tableName, titles[Columns.CurrrentStatus], values, sessionId, accessToken); //Filtering all the cases where the report has not been issued
        await filterExcelTableWithGraphAPI(workbookPath, tableName, titles[Columns.ReceptionDate], [`<=${until}`], sessionId, accessToken, false); //Filtering the table by the date of reception before the provided date
        const fields = sorting.map(column => [titles[column], true]);
        if (sorting)
            await sortExcelTableWithGraphAPI(workbookPath, tableName, fields, true, sessionId, accessToken); //Sorting the table in ascending order
        const visibleCells = await getVisibleCellsWithGraphAPI(workbookPath, tableName, sessionId, accessToken);
        const checker = (row) => `${row[Columns.ReceiptNumber]}${row[Columns.ReceptionDate]}${row[Columns.RegisterNumber]}`;
        let visibleRows = visibleCells.slice(1);
        (function getUniqueRowForEachCase() {
            //!Each case may have more than one record (row) in the table depending on the Parties. We need to exclude the repetions
            const checkersSet = new Set();
            visibleCells.forEach(row => checkersSet.add(checker(row)));
            visibleRows = Array.from(checkersSet)
                .map(check => visibleRows.find(row => checker(row) === check) || []);
        })();
        const groups = Object.values(optionsLists.caseType)
            .map(type => visibleRows.filter(row => row[Columns.CaseType] === type)); //!The cases must be arranged in the following order: 1) Civil Appeald, 2) Administrative, 3) Civil, 4) Labour Appealed, 5) Labour, 6) Tax, Person, Criminal Public Funds
        let rows = groups.flat()
            .map((row, index) => getWordTableRow(row, index)); //!The columns need to be checked
        (function splitInto2Columns() {
            return;
            const count = Math.floor(rows.length / 2) + (rows.length % 2); //eg: 6 if length = 11
            rows = rows
                .filter((row, index) => index < count) //eg: from rows[0] to rows[5]
                .map((row, index) => [...row, ...rows[count + index] || row.map(el => '')]);
        })();
        const date = new Date();
        const filePath = `baseURL/${form.template} Report until${until}@${getISODate(date).replace('-', '.')}`;
        await editReport(form, accessToken, filePath, rows, []);
        function getWordTableRow(row, index) {
            return [
                String(index + 1),
                getInternalNumber(row, form),
                getCaseNumber(row),
                `${row[Columns.ClaimantName]} x ${row[Columns.DefendantName]}`,
                row[Columns.AssignmentDate],
                row[Columns.ReceptionDate],
                row[Columns.FirstMeetingDate],
                getCurrentStauts(row, optionsLists.status.postponed),
                getCurrentStauts(row, optionsLists.status.ongoing),
                getCurrentStauts(row, optionsLists.status.notIssued),
                row[Columns.ProductionType],
                row[Columns.returnedExpertName],
                row[Columns.returnedExpertRegistration],
                row[Columns.Observations]
            ];
        }
        function getCurrentStauts(row, title) {
            if (row[Columns.CurrrentStatus] !== title)
                return '';
            if (title === optionsLists.status.postponed)
                return '✔';
            else if (title === optionsLists.status.ongoing)
                return getDate(row[Columns.LastMeetingDate]);
            else if (title === optionsLists.status.notIssued)
                return getDate(row[Columns.EndOfInstructionDate]);
            else
                return '';
        }
        function getInternalNumber(row, form) {
            let date;
            if (form.label === forms.receptionDate.label) //!Check if it is the Columns.ReceptionDate or Columns.ArraivalDate
                date = getDate(row[Columns.ReceptionDate]);
            else
                date = `/${getDate(row[Columns.ReceptionDate], true)}`; //We return the year as 'yyyy'
            return `${row[Columns.ReceiptNumber] || row[Columns.RegisterNumber]}\n${date}`;
        }
        function getCaseNumber(row) {
            //myArray(ColumnCaseNumber.index, col) & " " & ColumnCaseYear.name & " " & CaseYear & " " & ColumnCaseCourt.name & " " & myArray(ColumnCaseCourt.index, col)
            let caseYear = `${row[Columns.JudiciaryCaseYear] || row[Columns.CaseYear]}`;
            return `${row[Columns.CaseNumber]} ${titles[Columns.CaseYear]} ${caseYear} ${titles[Columns.CaseCourt]} ${row[Columns.CaseCourt]}`;
        }
        function getDate(dateString, year = false) {
            const date = new Date(dateString);
            if (year)
                return date.getFullYear().toString();
            return [date.getDate(), (date.getMonth() + 1), date.getFullYear].join('/');
        }
    }
    function showCRUDForm(crud) {
        const args = [titles, table.slice(1), form];
        const api = [workbookPath, tableName, accessToken, sessionId, true];
        if (crud === forms.addRecord)
            addRecord(args, ...api);
        else if (crud === forms.modifyRecord)
            modifyRecord(args, ...api);
        else if (crud === forms.deleteRecord)
            deleteRecord(args, ...api);
    }
}
async function promptForUserInput(type, message) {
    return new Promise((resolve, reject) => {
        // Create the prompt container
        const div = document.createElement('div');
        div.id = 'datePrompt';
        div.classList.add('prompt');
        const p = document.createElement('p');
        p.textContent = message;
        div.appendChild(p);
        // Create the date input field
        const input = document.createElement('input');
        input.type = type;
        div.appendChild(input);
        // Create the OK button
        const ok = document.createElement('button');
        ok.textContent = "OK";
        ok.style.marginTop = "10px";
        div.appendChild(ok);
        // Create the Cancel button
        const cancel = document.createElement('button');
        cancel.textContent = "Cancel";
        cancel.style.marginLeft = "10px";
        div.appendChild(cancel);
        // Append the prompt to the document
        document.body.appendChild(div);
        // OK Button Logic
        ok.onclick = () => {
            if (type === 'date' && input.valueAsDate) {
                const date = getISODate(input.valueAsDate);
                div.remove(); // Clean up the prompt
                resolve(date); // Resolve the Promise with the date
            }
            else if (!input.value) {
                alert(`Please select a valid ${type} before clicking OK.`);
            }
        };
        // Cancel Button Logic
        cancel.onclick = () => {
            div.remove(); // Clean up the prompt
            reject("User canceled the date input."); // Reject the Promise
        };
    });
}
function addRecord([titles, body, form], workbookPath, tableName, accessToken, sessionId, filter) {
    const main = [
        Columns.ClaimantName,
        Columns.DefendantName,
        Columns.PartyName,
        Columns.PartyAddress,
        Columns.Observations
    ];
    const details = [
        Columns.ReceiptNumber,
        Columns.ArrivalDate,
        Columns.AssignmentDate,
        Columns.ReceptionDate,
        Columns.RegisterNumber,
        Columns.FileNumber,
        Columns.CaseNumber,
        Columns.CaseYear,
        Columns.JudiciaryCaseYear,
        Columns.CaseCourt,
        Columns.CaseType,
        Columns.FirstMeetingDate,
        Columns.LastMeetingDate
    ];
    const returned = [
        Columns.returnedExpertName,
        Columns.returnedExpertRegistration,
        Columns.returnedReportDate,
        Columns.returnedProductionType
    ];
    const status = [
        Columns.CurrrentStatus,
        Columns.EndOfInstructionDate,
        Columns.AchievementMonth,
        Columns.ProductionType
    ];
    createBlock(main, 'main', form);
    createBlock(details, 'details', form);
    createBlock(returned, 'returned', form);
    createBlock(status, 'status', form);
    (function insertAddBtn() {
        const btn = document.createElement('button');
        btn.classList.add('button');
        btn.onclick = addRecord;
        form.appendChild(btn);
    })();
    function addRecord() {
        const inputs = Array.from(document.getElementsByTagName('input'));
        const row = Object.keys(Columns).map((key, index) => getValues(index, inputs));
        addRow(row, filter);
        async function addRow(row, filter = false) {
            if (!row)
                throw new Error('The row is not valid');
            const visibleCells = await addRowToExcelTableWithGraphAPI(row, body.length - 2, workbookPath, tableName, accessToken, filter);
            if (!visibleCells)
                return alert('There was an issue with the adding or the filtering, check the console.log for more details');
            alert('Row aded and the table was filtered');
            displayVisibleCells(visibleCells, form);
            spinner(false); //We hide the spinner
        }
        ;
    }
    function createBlock(columns, css, form) {
        const dates = [
            Columns.ArrivalDate,
            Columns.AssignmentDate,
            Columns.ReceptionDate,
            Columns.FirstMeetingDate,
            Columns.LastMeetingDate,
            Columns.EndOfInstructionDate,
            Columns.returnedReportDate
        ];
        const div = document.createElement('div');
        form.appendChild(div);
        columns.forEach(column => {
            const container = document.createElement('div');
            div.appendChild(container);
            (function insertLabel() {
                const label = document.createElement('label');
                label.innerText = titles[column];
                container.appendChild(label);
            })();
            (function insertInput() {
                const input = document.createElement('input');
                input.dataset.index = column.toString();
                input.type = 'text';
                if (dates.includes(column))
                    input.type = 'date';
                input.classList.add(css);
                if (column = Columns.CaseType)
                    addDataList(input, column, Object.values(optionsLists.caseType));
                container.appendChild(input);
            })();
        });
        function addDataList(input, index, unique) {
            const list = `col${index}`;
            input.setAttribute('list', list);
            if (!unique)
                unique = getUniqueValues(index, body);
            createDataList(list, unique);
        }
    }
}
;
function modifyRecord([titles, body, form], workbookPath, tableName, accessToken, sessionId, filter) {
    (function insertModifyBtn() {
        const btn = document.createElement('button');
        btn.classList.add('button');
        btn.onclick = modifyRecord;
        form.appendChild(btn);
    })();
    function modifyRecord() {
        const inputs = Array.from(document.getElementsByTagName('input'));
        const row = Object.keys(Columns).map((key, index) => getValues(index, inputs));
        modifyRow(row, filter);
        async function modifyRow(row, filter = false) {
            if (!row)
                throw new Error('The row is not valid');
            const visibleCells = await addRowToExcelTableWithGraphAPI(row, body.length - 2, workbookPath, tableName, accessToken, filter);
            if (!visibleCells)
                return alert('There was an issue with the adding or the filtering, check the console.log for more details');
            alert('Row aded and the table was filtered');
            displayVisibleCells(visibleCells, form);
            spinner(false); //We hide the spinner
        }
        ;
    }
}
function deleteRecord([titles, body, form], workbookPath, tableName, accessToken, sessionId, filter) {
    (function insertDeleteBtn() {
        const btn = document.createElement('button');
        btn.classList.add('button');
        btn.onclick = deleteRecord;
        form.appendChild(btn);
    })();
    function deleteRecord() {
        const inputs = Array.from(document.getElementsByTagName('input'));
        const row = Object.keys(Columns).map((key, index) => getValues(index, inputs));
        deleteRow(row, filter);
        async function deleteRow(row, filter = false) {
            if (!row)
                throw new Error('The row is not valid');
            const visibleCells = await addRowToExcelTableWithGraphAPI(row, body.length - 2, workbookPath, tableName, accessToken, filter);
            if (!visibleCells)
                return alert('There was an issue with the adding or the filtering, check the console.log for more details');
            alert('Row aded and the table was filtered');
            displayVisibleCells(visibleCells, form);
            spinner(false); //We hide the spinner
        }
        ;
    }
}
function getValues(index, inputs) {
    const input = getInputByIndex(inputs, index);
    if ([
        Columns.ArrivalDate,
        Columns.AssignmentDate,
        Columns.ReceptionDate,
        Columns.FirstMeetingDate,
        Columns.LastMeetingDate,
        Columns.EndOfInstructionDate,
        Columns.returnedReportDate
    ].includes(index))
        return getISODate(input.valueAsDate || undefined); //Those are the 2 date columns
    else
        return input.value;
}
function displayVisibleCells(visibleCells, form) {
    const tableDiv = createDivContainer();
    const table = document.createElement('table');
    table.classList.add('table');
    tableDiv.appendChild(table);
    const columns = Object.values(Columns);
    const rowClass = 'excelRow';
    (function insertTableHeader() {
        if (!tableTitles)
            throw new Error('No Table Titles');
        const headerRow = document.createElement('tr');
        headerRow.classList.add(rowClass);
        const thead = document.createElement('thead');
        table.appendChild(thead);
        thead.appendChild(headerRow);
        tableTitles.forEach((cell, index) => {
            if (!columns.includes(index))
                return;
            addTableCell(headerRow, cell, 'th');
        });
    })();
    (function insertTableRows() {
        const tbody = document.createElement('tbody');
        table.appendChild(tbody);
        visibleCells.forEach((row, index) => {
            if (index < 1)
                return; //We exclude the header row
            if (!row)
                return;
            const tr = document.createElement('tr');
            tr.classList.add(rowClass);
            tbody.appendChild(tr);
            row.forEach((cell, index) => {
                if (!columns.includes(index))
                    return;
                addTableCell(tr, cell, 'td');
            });
        });
    })();
    if (!form)
        throw new Error('The form element was not found');
    if (form) {
        form?.insertAdjacentElement('afterend', tableDiv);
    }
    function createDivContainer() {
        const id = 'retrieved';
        let tableDiv = document.getElementById(id);
        if (tableDiv) {
            tableDiv.innerHTML = '';
            return tableDiv;
        }
        ;
        tableDiv = document.createElement('div');
        tableDiv.classList.add('table-div');
        tableDiv.id = id;
        return tableDiv;
    }
    function addTableCell(parent, text, tag) {
        const cell = document.createElement(tag);
        //   cell.classList.add(css);
        cell.textContent = text;
        parent.appendChild(cell);
    }
}
;
async function editReport(form, accessToken, filePath, rows, contentControls) {
    if (!accessToken || !filePath)
        return;
    const templatePath = `baseURL${form.template}`;
    const blob = await fetchFileFromOneDriveWithGraphAPI(accessToken, templatePath);
    if (!blob)
        return;
    const [doc, zip] = await convertBlobIntoXML(blob);
    if (!doc)
        return;
    const lang = 'AR';
    const schema = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    (function editTable() {
        if (!rows)
            return;
        const tables = getXMLElements(doc, "tbl");
        const table = getXMLTableByTitle(tables, form.table);
        if (!table)
            return;
        const firstRow = getXMLElements(table, 'tr', 1);
        rows.forEach((row, index) => {
            const newXmlRow = insertRowToXMLTable(NaN, true) || table.appendChild(createXMLElement('tr'));
            if (!newXmlRow)
                return;
            // const isTotal = totalsLabels?.includes(row[0]);
            // const isLast = index === rows.length - 1;
            //return editCells(newXmlRow, row, isLast, isTotal);
            return editCells(newXmlRow, row);
        });
        firstRow.remove(); //We remove the first row when we finish
        function editCells(tableRow, values) {
            const cells = getXMLElements(tableRow, 'tc') || values.map(v => tableRow.appendChild(createXMLElement('tc'))); //getting all the cells in the row element
            cells.forEach((cell, index) => {
                const textElement = getXMLElements(cell, 't', 0) || appendParagraph(cell);
                if (!textElement)
                    return console.log('No text element was found !');
                const pPr = setTextLanguage(cell, lang); //We call this here in order to set the language for all the cells. It returns the pPr element if any.
                textElement.textContent = values[index];
            });
        }
        function insertRowToXMLTable(after = -1, clone = false) {
            if (clone)
                return cloneFirstRow();
            else
                return create();
            function create() {
                if (!table)
                    return;
                const row = createXMLElement("tr");
                after >= 0 ? getXMLElements(table, 'tr', after)?.insertAdjacentElement('afterend', row) :
                    table.appendChild(row);
                return row;
            }
            function cloneFirstRow() {
                const row = firstRow.cloneNode(true);
                table?.appendChild(row);
                return row;
            }
            ;
        }
        function getXMLTableByTitle(tables, title, property = 'tblCaption') {
            if (!title)
                return;
            return tables
                .filter(table => tblCaption(table))
                .find(table => tblCaption(table).getAttribute('w:val') === title);
            function tblCaption(table) {
                return getXMLElements(table, property, 0);
            }
        }
    })();
    (function editContentControls() {
        if (!contentControls)
            return;
        const ctrls = getXMLElements(doc, "sdt");
        contentControls
            .forEach(([title, text]) => {
            const control = findXMLContentControlByTitle(ctrls, title);
            if (!control)
                return;
            editXMLContentControl(control, text);
        });
        function findXMLContentControlByTitle(ctrls, title) {
            return ctrls.find(control => getXMLElements(control, "alias", 0)?.getAttributeNS(schema, 'val') === title);
        }
        function editXMLContentControl(control, text) {
            if (!text)
                return control.remove();
            const sdtContent = getXMLElements(control, "sdtContent", 0);
            if (!sdtContent)
                return;
            const paragTemplate = getParagraphOrRun(sdtContent); //This will set the language for the paragraph or the run
            if (!paragTemplate)
                return console.log('No template paragraph or run were found !');
            setTextLanguage(paragTemplate, lang); //We amend the language element to the "w:pPr" or "r:pPr" child elements of paragTemplate
            text.split('\n')
                .forEach((parag, index) => editParagraph(parag, index));
            function editParagraph(parag, index) {
                let textElement;
                if (index < 1)
                    textElement = getXMLElements(paragTemplate, 't', index);
                else
                    textElement = appendParagraph(paragTemplate, sdtContent); //We pass sdtContent as parent argument
                if (!textElement)
                    return console.log('No textElement was found !');
                textElement.textContent = parag;
            }
        }
    })();
    await convertXMLToBlobAndUpload(doc, zip, filePath, accessToken);
    /**
     * Adds a new paragraph XML element or appends a cloned paragraph, and in both cases, it returns the textElement of the paragraph
     * @param {Element} element - The element to which the new paragraph will be appended if the parent argument is not provided. If the parent argument is provided, the element will be cloned assuming that this is a pargraph element
     * @param {Elemenet} parent - If provided, element will be cloned and appended to parent.
     * @returns {Element} the textElemenet attached to the paragraph
     */
    function appendParagraph(element, parent) {
        if (parent)
            return clone();
        else
            return create();
        function clone() {
            const parag = element?.cloneNode(true);
            parent?.appendChild(parag);
            return getXMLElements(parag, 't', 0);
        }
        function create() {
            const parag = element.appendChild(createXMLElement('p'));
            parag.appendChild(createXMLElement('pPr'));
            const run = parag.appendChild(createXMLElement('r'));
            return run.appendChild(createXMLElement('t'));
        }
    }
    function createXMLElement(tag, parent) {
        return doc.createElementNS(schema, tag);
    }
    function getXMLElements(xmlDoc, tag, index = NaN) {
        const elements = xmlDoc.getElementsByTagNameNS(schema, tag);
        if (!isNaN(index))
            return elements?.[index];
        return Array.from(elements);
    }
    /**
     * Looks for a child "w:p" (paragraph) element, if it doesn't find any, it looks for a "w:r" (run) element.
     * @param {Element} parent - the parent XML of the paragraph or run element we want to retrieve.
     * @returns {Element | undefined} - an XML element representing a "w:p" (paragraph) or, if not found, a "w:r" (run), or undefined
     */
    function getParagraphOrRun(parent) {
        return getXMLElements(parent, 'p', 0) || getXMLElements(parent, 'r', 0);
    }
    /**
     * Finds a "w:pPr" XML element (property element) which is a child of the XML parent element passed as argument. If does not find it, it looks for a "w:rPr" XML element. When it finds either a "w:pPr" or a "w:rPr" element, it appends a "w:lang" element to it, and sets its "w:val" attribute to the language passed as "lang"
     * @param {Element} parent - the XML element containing the paragraph or the run for which we want to set the language.
     * @returns {Element | undefined} - the "w:pPr" or "w:rPr" property XML element child of the parent element passed as argument
     */
    function setTextLanguage(parent, lang) {
        const pPr = getXMLElements(parent, 'pPr', 0) ||
            getXMLElements(parent, 'rPr', 0);
        if (!pPr)
            return;
        pPr
            .appendChild(createXMLElement('lang')) //appending a "w:lang" element
            .setAttributeNS(schema, 'val', `${lang.toLowerCase()}-${lang.toUpperCase()}`); //setting the "w:val" attribute of "w:lang" to the appropriate language like "fr-FR"
        return pPr;
    }
}
//# sourceMappingURL=Marianne.js.map