export interface IPayload {
    tenant: any;
    SPOUrl: string;
    ItemID?: number;
}

export interface IPayloadSaveOneDrive {
    tenant: any;
    SPOUrl: string;
    FileName: string;
    sourceFileName: string;
    FolderName: string;
    userGuidId: string;
}
