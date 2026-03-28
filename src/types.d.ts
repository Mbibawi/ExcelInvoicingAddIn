type header = {
    "Authorization": string;
    "Content-Type": string;
    "workbook-session-id"?: string;
};
type lable = {
    nature: string[];
    FR: string;
    EN: string;
};
type RT = {
    title: string;
    value: string |number,
    label?: string;
    col?: number;
    type?: string;
};

type LeaseCtrls = {
    owner: RT;
    adress: RT;
    tenant: RT;
    leaseDate: RT;
    leaseType: RT;
    initialIndex: RT;
    indexQuarter: RT;
    initialIndexDate: RT;
    baseIndex: RT;
    baseIndexDate: RT;
    index: RT;
    indexDate: RT;
    currentLease: RT;
    revisionDate: RT;
    initialYear: RT;
    revisionYear: RT;
    baseYear: RT;
    anniversaryDate: RT;
    newLease: RT;
    nextRevision: RT;
    startingMonth: RT;
};

type InvoiceCtrls = {
    dateLabel: RT;
    date: RT;
    numberLabel: RT;
    number: RT;
    subjectLable: RT;
    subject: RT;
    fee: RT;
    amount: RT;
    vat: RT;
    disclaimer: RT;
    clientName: RT;
    adress: RT;
};

type settingInput = {
    label: string;//The lable of the input that will be created in the settings pannel
    name: string; //the name of the setting that will be saved in the local storage
    value: string;//The value of the setting
};

type setting = {
    workBook?: settingInput;
    wordTemplate?: settingInput;
    saveTo?: settingInput;
    tableName?: settingInput
}

type settings = {
    issueInvoice: setting,
    leases?: setting,
    Letter?: setting,
    reports?: setting
}

type values = [number, number];
type folderItem = { name: string; id: string; folder: any; createdDateTime: string; lastModifiedDateTime: string };
type fileItem = { name: string; id: string; file: any; createdDateTime: string; lastModifiedDateTime: string; "@microsoft.graph.downloadUrl": string };
type InputCol = [HTMLInputElement, number];
type onClick = (ev: MouseEvent) => any;
declare class JSZip {
    constructor();
    loadAsync(data: any): Promise<JSZip>;
    file(path: string, doc?: string): any; // or create a more detailed interface for the file
    generateAsync(options: { type: string }): Promise<any>;
    files: { [key: string]: any };
}