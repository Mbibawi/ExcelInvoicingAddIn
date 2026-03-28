import { LawFirm } from "./pwaVersion.js";
import { saveSettings } from "./index.js";
export const splitter = "; OR "; //This is the splitter that will be used to separate multiple values in the input fields. We need to use a splitter that is not likely to be included in the values themselves.
class LawFirmUI {
    lf;
    constructor() {
        this.lf = new LawFirm(true);
    }
    appendUIBtns(homeBtn = false) {
        const container = byID('btns');
        if (!container)
            return;
        container.innerHTML = "";
        if (homeBtn)
            return this.appendBtn(container, 'home', 'Back to Main', () => this.appendUIBtns());
        this.appendBtn(container, 'entry', 'Add Entry', () => this.lf.addNewEntry());
        this.appendBtn(container, 'invoice', 'Invoice', () => this.lf.issueInvoice());
        this.appendBtn(container, 'letter', 'Letter', () => this.lf.issueLetter());
        this.appendBtn(container, 'lease', 'Leases', () => this.lf.issueLeaseLetter());
        this.appendBtn(container, 'search', 'Search Files', () => this.lf.searchFiles());
        this.appendBtn(container, 'settings', 'Settings', () => saveSettings());
    }
    appendBtn(container, id, text, onClick) {
        const btn = document.createElement('button');
        btn.id = id;
        btn.classList.add("ms-Button");
        btn.innerText = text;
        btn.onclick = onClick;
        container?.appendChild(btn);
        return btn;
    }
}
const LFUI = new LawFirmUI();
export function showLawFirmUI(homeBtn) {
    LFUI.appendUIBtns(homeBtn);
}
;
export function byID(id = "form") { return document.getElementById(id); }
;
/**
 *
 * @param select
 * @param uniqueValues
 * @param  {boolean} combine - determines whether we will add to the list an element containing all the options. Its defalult value is "false"
 */
export function populateSelectElement(select, uniqueValues, combine = false) {
    const list = createDataList(select.id, uniqueValues, combine);
    if (!list)
        return;
    select.setAttribute('list', list.id);
    select.autocomplete = "on";
    return list;
}
/**
 *
 * @param id
 * @param uniqueValues
 * @param combine
 * @returns
 */
function createDataList(id, uniqueValues, combine = false) {
    //const uniqueValues = Array.from(new Set(visible.map(row => row[i])));
    if (!id || uniqueValues?.length < 2)
        return;
    id += 's';
    // Create a new datalist element
    let dataList = Array.from(document.getElementsByTagName('datalist')).find(list => list.id === id);
    if (dataList)
        dataList.remove();
    dataList = document.createElement('datalist');
    dataList.id = id;
    // Append options to the datalist
    uniqueValues.forEach(option => addOption(option));
    if (combine)
        addOption(uniqueValues.join(splitter));
    // Attach the datalist to the body or a specific element
    document.body.appendChild(dataList);
    function addOption(option) {
        const optionElement = document.createElement('option');
        optionElement.value = option;
        dataList?.appendChild(optionElement);
    }
    return dataList;
}
;
//# sourceMappingURL=ui.js.map