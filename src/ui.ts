import { LawFirm, Marianne, saveSettings} from "./pwaVersion.js";

export const splitter = "; OR ";//This is the splitter that will be used to separate multiple values in the input fields. We need to use a splitter that is not likely to be included in the values themselves.

export interface IUserInterface<>{
    appendUIBtns(homeBtn?: boolean): HTMLButtonElement | undefined;
}


class LawFirmUI implements IUserInterface {
    private lf;
    constructor(lawfirm: new () => LawFirm) {
        this.lf = new lawfirm();
    }

    appendUIBtns(homeBtn: boolean = false) {
        const container = byID('btns');
        if (!container) return;
        container!.innerHTML = "";
        if (homeBtn) return appendUIBtn(container, 'home', 'Back to Main', () => this.appendUIBtns());
        appendUIBtn(container, 'entry', 'Add Entry', () => this.lf.addNewEntry());
        appendUIBtn(container, 'invoice', 'Invoice', () => this.lf.issueInvoice());
        appendUIBtn(container, 'letter', 'Letter', () => this.lf.issueLetter());
        appendUIBtn(container, 'lease', 'Leases', () => this.lf.issueLeaseLetter());
        appendUIBtn(container, 'search', 'Search Files', () => this.lf.searchFiles());
        appendUIBtn(container, 'settings', 'Settings', () => saveSettings(this));
    }
}

class MarianneUI implements IUserInterface {
    private mr;
    constructor(marianne: new () => Marianne) {
        this.mr = new marianne();
    }
    appendUIBtns(homeBtn?: boolean): HTMLButtonElement | undefined {
        //overriding the function
        return undefined
    }
}

export const lfUI = new LawFirmUI(LawFirm);
export const mrUI = new MarianneUI(Marianne);

export function showUI(ui:IUserInterface, homeBtn:boolean = false) {
    ui.appendUIBtns(homeBtn);
};


export function byID(id: string = "form") { return document.getElementById(id) };

export function appendUIBtn(container: HTMLElement, id: string, text: string, onClick: onClick) {
    const btn = document.createElement('button');
    btn.id = id;
    btn.classList.add("ms-Button");
    btn.innerText = text;
    btn.onclick = onClick;
    container?.appendChild(btn);
    return btn
}

/**
 *
 * @param select 
 * @param uniqueValues 
 * @param  {boolean} combine - determines whether we will add to the list an element containing all the options. Its defalult value is "false"
 */
export function populateSelectElement(select: HTMLInputElement, uniqueValues: string[], combine: boolean = false) {
    const list = createDataList(select.id, uniqueValues, combine);
    if (!list) return;
    select.setAttribute('list', list.id);
    select.autocomplete = "on";
    return list
}

/**
 * 
 * @param id 
 * @param uniqueValues 
 * @param combine 
 * @returns 
 */
function createDataList(id: string, uniqueValues: string[], combine: boolean = false) {
    //const uniqueValues = Array.from(new Set(visible.map(row => row[i])));
    if (!id || uniqueValues?.length < 2) return;
    id += 's';

    // Create a new datalist element
    let dataList = Array.from(document.getElementsByTagName('datalist')).find(list => list.id === id);
    if (dataList) dataList.remove();
    dataList = document.createElement('datalist');
    dataList.id = id;
    // Append options to the datalist
    uniqueValues.forEach(option => addOption(option));

    if (combine)
        addOption(uniqueValues.join(splitter));

    // Attach the datalist to the body or a specific element
    document.body.appendChild(dataList);
    function addOption(option: string) {
        const optionElement = document.createElement('option');
        optionElement.value = option;
        dataList?.appendChild(optionElement);
    }
    return dataList
};
