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
    value: string
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

type values = [number, number];
type folderItem = { name: string; id: string; folder: any; createdDateTime: string; lastModifiedDateTime: string };
type fileItem = { name: string; id: string; file: any; createdDateTime: string; lastModifiedDateTime: string; "@microsoft.graph.downloadUrl": string };
type InputCol = [HTMLInputElement, number];
