import { IFileFolder } from "./IFileFolder";

export interface IBreadCrumb {
    Location: string;
    InCurrentView: boolean;
    LocationDisplay: string;
    Order: number;
    IsSkipped: boolean;
    Contents: IFileFolder[];
}
