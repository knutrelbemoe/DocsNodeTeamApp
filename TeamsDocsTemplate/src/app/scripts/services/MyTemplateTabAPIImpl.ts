import { ApiAddress as ApiURL, ApiAddress } from './../Core/constants';
import { IMyTemplateTabAPI } from './MyTemplateTabAPI';
import { ILocationPayload } from '../interfaces/ILocationPayload';
import { IDefaultLocation } from '../Core/interfaces/IDefaultLocation';

export class MyTemplateTabAPIImpl implements IMyTemplateTabAPI {

    public GetLocationDetails(payload: ILocationPayload): Promise<any> {
        return new Promise((resolve, reject) => {
            fetch(ApiURL.GET_LOCATION_DETAILS_URL, {
                method: 'post',
                body: JSON.stringify(payload),
            }).then(response => {
                return response.json();
            }).then(responseJson => {
                resolve(responseJson);
            }).catch(error => {
                console.error('Error while converting json- ', error);
                reject(error);
            }).catch(error => {
                console.error('Error while getting location details- ', error);
                reject(error);
            });
        });

    }

    public GetDefaultLocation(payload: IDefaultLocation): Promise<string> {

        return new Promise((resolve, reject) => {
            fetch(ApiURL.GET_DEFAULT_LOCATION_URL, {
                method: 'post',
                body: JSON.stringify(payload),
            }).then(response => {
                return response.json();
            }).then(responseJson => {

                let resultArray: any[] = responseJson.d.results;
                if (resultArray.length > 0 && resultArray[0].Active && resultArray[0].FolderPath) {
                    resolve(resultArray[0].FolderPath);
                }
                else if (resultArray.length > 0 && resultArray[0].Active && !resultArray[0].FolderPath) {
                    resolve("Home");
                }
                else {
                    resolve("");
                }

            }).catch(error => {
                console.error('Error while converting json- ', error);
                reject(error);
            }).catch(error => {
                console.error('Error GetDefaultLocation - ', error);
                reject(error);
            });
        });

    }

    public SetDefaultLocation(payload: IDefaultLocation): Promise<any> {
        return new Promise((resolve, reject) => {
            fetch(ApiURL.SET_DEFAULT_LOCATION_URL, {
                method: 'post',
                body: JSON.stringify(payload),
            }).then(response => {
                if (response.status === 200) {
                    resolve(true);
                }
            }).catch(error => {
                console.error('Error SetDefaultLocation - ', error);
                reject(error);
            });
        });
    }

}
