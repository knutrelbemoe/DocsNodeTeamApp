import { ApiAddress as ApiURL } from './../Core/constants';
import { IPinnedLocations, IChannelFolder } from '../Core/interfaces';
import { ITeams, ITeamChannel } from './../Core/interfaces';
import { IPayload } from '../interfaces/IPayload';

export class DeleteMyLibraryApi {
    public async deletePinnedLocation(payload: any): Promise<any> {
      return await this._deletePinnedLocation(payload).then((response: boolean) => {
        return response;
       });
     }

    private async _deletePinnedLocation(payload: IPayload): Promise<any> {
         return await fetch(ApiURL.DELETE_PINNED_LOCATION, {
            method: 'post',
            body: JSON.stringify(payload),
          })
           .then(response => {
               return response.ok;
        })
           .catch(error => {
                console.log('Error while deleting pinned location - ', error);
            });
    }
}
