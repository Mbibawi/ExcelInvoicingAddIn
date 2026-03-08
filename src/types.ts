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
    tag: string;
    col: number | undefined;
    value: string
};
type values = [number, number];
type folderItem = { name: string; id: string; folder: any; createdDateTime: string; lastModifiedDateTime: string };
type fileItem = { name: string; id: string; file: any; createdDateTime: string; lastModifiedDateTime: string; "@microsoft.graph.downloadUrl": string };
