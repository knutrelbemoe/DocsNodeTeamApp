import { ILocationPayload } from '../interfaces/ILocationPayload';
import { IDefaultLocation } from '../Core/interfaces/IDefaultLocation';

export interface IMyTemplateTabAPI {
    GetLocationDetails(payload: ILocationPayload): Promise<any>;
    GetDefaultLocation(payload: IDefaultLocation): Promise<string>;
    SetDefaultLocation(payload: IDefaultLocation): Promise<any>;
}
