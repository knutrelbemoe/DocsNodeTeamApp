import { ApiAddress as ApiURL } from './../Core/constants';
import { IPinnedLocations, IChannelFolder, IOneDrive } from '../Core/interfaces';
import { ITeams, ITeamChannel } from './../Core/interfaces';
import { each } from 'bluebird';

export class SelectMyLibraryApi {

    public async getMyPinnedLocation(payload: any): Promise<any> {

        return await this._getPinnedLocation(payload).then((response: any[]) => {
            return response;
        });


    }
    public async getMyMsTeams(payload: any): Promise<any> {
        console.log(payload);
        return await this._getMsTeams(payload).then((response: any[]) => {
            return response;
        });


    }
    public async getMyMsTeamChannels(payload: any): Promise<any> {

        return await this._getMsTeamChannels(payload).then((response: any[]) => {
            return response;
        });


    }
    public async getMyOneDriveChilds(payload: any): Promise<any> {

        return await this._getOneDriveChilds(payload).then((response: any[]) => {
            return response;
        });


    }
    public async getTeamsChannelSubFolders(payload: any, level: number): Promise<any> {

        return await this._getSPOLibSubFolders(payload, level).then((response: any[]) => {
            return response;
        });


    }
    public async getTeamsChannelTabs(payload: any, level: number): Promise<any> {

        return await this._getTeamChannelTabSubFolders(payload, level).then((response: any[]) => {
            return response;
        });


    }
    public async getTeamsWebUrl(payload: any): Promise<any> {

        return await this._getTeamWebUrl(payload).then((response: any[]) => {
            return response;
        });


    }
    public async getCurrentUserDisplayName(payload: any): Promise<any> {

        return await this._getCurrentUserDisplayName(payload).then((response: any[]) => {
            return response;
        });


    }

    public async getTeamLibraryName(payload: any): Promise<any> {

        return await this._getTeamLibraryNames(payload).then((response: any[]) => {
            return response;
        });


    }
    public async getMyOneDrive(payload: any): Promise<any> {
        console.log(payload);
        return await this._getOneDrive(payload).then((response: any[]) => {
            return response;
        });


    }

    private async _getMsTeams(payload: any): Promise<any> {
        return await fetch(ApiURL.GET_USER_MSTEAMS, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            const allMsTeams: ITeams[] = [];

            data.value.map(team => {
                const eachTeams: ITeams = {};
                // eachTemplate.isSelected = false;

                if (team) {
                    eachTeams.TeamName = team.displayName;
                    eachTeams.TeamId = team.id;
                    eachTeams.TeamUrl = team["team@odata.navigationLink"];
                }

                allMsTeams.push(eachTeams);
            });

            return allMsTeams;
        }).catch(error => {
            console.log("Error while fetching templates - ", error);

        });
    }
    private async _getMsTeamChannels(payload: any): Promise<any> {
        return await fetch(ApiURL.GET_MSTEAM_CHANNELS, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            const allMsTeamChannels: any[] = [];

            data.map(channel => {
                const eachTeamChannel: ITeamChannel = {};
                if (channel) {
                    eachTeamChannel.ChannelDisplayName = channel.displayName;
                    eachTeamChannel.ChannelUrl = channel.webUrl;
                    eachTeamChannel.ChannelId = channel.id;
                    eachTeamChannel.FolderType = channel["@odata.type"];

                    if (channel.membershipType && channel.membershipType === "private") {
                        eachTeamChannel.FolderType = eachTeamChannel.FolderType + ".private";
                    }

                }

                allMsTeamChannels.push(eachTeamChannel);
            });

            return allMsTeamChannels;
        }).catch(error => {
            console.log("Error while fetching templates - ", error);

        });
    }
    private async _getSPOLibSubFolders(payload: any, level: number): Promise<any> {
        return await fetch(ApiURL.GET_CHANNEL_FOLDER, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            const allMsTeamsChannelFolder: any[] = [];

            data.d.results.map(folder => {
                const eachFolder: IChannelFolder = {};
                // eachTemplate.isSelected = false;

                eachFolder.FolderName = folder.Name;
                eachFolder.FolderUrl = folder.__metadata.uri;
                eachFolder.FolderRelativeUrl = folder.ServerRelativeUrl;
                eachFolder.FolderLevel = level;
                eachFolder.ChildCount = folder.ItemCount;
                eachFolder.FolderType = payload.FolderType;

                if (payload.FolderType === "tab.sharepoint.folder" && folder.Name === "Forms") {
                    console.log(eachFolder);
                }
                else if (payload.FolderType && payload.FolderType.includes("#microsoft.graph.channel")) {
                    let siteUrl = folder.__metadata.uri && folder.__metadata.uri.substr(0, folder.__metadata.uri.toLowerCase().indexOf("/_api"));
                    if (payload.SiteUrl && siteUrl.toLowerCase() !== payload.SiteUrl.toLowerCase()) {
                        eachFolder.FolderType = "#microsoft.graph.channel.private";
                        allMsTeamsChannelFolder.push(eachFolder);
                    }
                    else {
                        allMsTeamsChannelFolder.push(eachFolder);
                    }
                }
                else {
                    allMsTeamsChannelFolder.push(eachFolder);
                }
            });

            return allMsTeamsChannelFolder;
        }).catch(error => {
            console.log("Error while fetching templates - ", error);

        });
    }
    private async _getTeamChannelTabSubFolders(payload: any, level: number): Promise<any> {
        return await fetch(ApiURL.GET_TEAM_CHANNEL_TAB_URL, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            const allMsTeamsChannelFolder: any[] = [];

            data.value.map(folder => {
                const eachFolder: IChannelFolder = {};
                // eachTemplate.isSelected = false;

                eachFolder.FolderName = decodeURIComponent(folder.displayName);
                eachFolder.FolderUrl = folder.configuration.contentUrl;
                eachFolder.FolderRelativeUrl = String(folder.configuration.contentUrl).substr(String(folder.configuration.contentUrl).toLowerCase().indexOf("/sites"), (String(folder.configuration.contentUrl).length - 1));
                eachFolder.FolderLevel = level;
                eachFolder.FolderType = "tab.sharepoint.folder";
                allMsTeamsChannelFolder.push(eachFolder);
            });

            return allMsTeamsChannelFolder;
        }).catch(error => {
            console.log("Error while fetching templates - ", error);

        });
    }
    private async _getPinnedLocation(payload: any): Promise<any> {
        return await fetch(ApiURL.GET_PIN_LOCATION, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            const allPinnedLocation: IPinnedLocations[] = [];

            data.d.results.map(pinlocation => {
                const eachPinnedLocation: IPinnedLocations = {};

                // eachTemplate.isSelected = false;
                if (pinlocation.DocumentLibrary) {
                    eachPinnedLocation.LocationName = decodeURI(pinlocation.DocumentLibrary);
                }
                if (pinlocation.DocumentLibraryURL) {
                    // eachPinnedLocation.LocationName = decodeURI(pinlocation.DocumentLibraryURL.Description);
                    eachPinnedLocation.LocationUrl = decodeURI(pinlocation.DocumentLibraryURL.Url);
                    eachPinnedLocation.IsPinned = true;
                    eachPinnedLocation.PinnedType = pinlocation.PinnedType;
                    eachPinnedLocation.SiteUrl = pinlocation.SiteURL.Url;
                    eachPinnedLocation.key = decodeURI(pinlocation.DocumentLibraryURL.Description);
                    eachPinnedLocation.Id = pinlocation.Id;
                }

                allPinnedLocation.push(eachPinnedLocation);
            });

            return allPinnedLocation;
        }).catch(error => {
            console.log("Error while fetching templates - ", error);

        });
    }
    private async _getTeamWebUrl(payload: any): Promise<any> {
        return await fetch(ApiURL.GET_TEAMS_WEBURL, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            return data.value;
        }).catch(error => {
            console.log("Error while fetching templates - ", error);

        });
    }
    private async _getCurrentUserDisplayName(payload: any): Promise<any> {
        return await fetch(ApiURL.GET_CURRENT_USER, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {

            return data.displayName;
        }).catch(error => {
            console.log("Error while fetching templates - ", error);

        });
    }

    private async _getTeamLibraryNames(payload: any): Promise<any> {
        return await fetch(ApiURL.GET_TEAM_LIBRARY_NAMES, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            const teamLibInternalNames: any[] = [];

            data.d.results.map(libNames => {
                const eachLibInternalNames: any = {};

                // eachTemplate.isSelected = false;
                if (libNames.Title && (libNames.Title.toLowerCase() === "documents" || libNames.Title.toLowerCase() === "dokumenter")) {
                    eachLibInternalNames.EntityTypeName = libNames.EntityTypeName.toString().replace("_x0020_", " ");
                    eachLibInternalNames.Title = decodeURI(libNames.Title);
                    teamLibInternalNames.push(eachLibInternalNames);
                }


            });

            return teamLibInternalNames;
        }).catch(error => {
            console.log("Error while fetching templates - ", error);

        });
    }
    private async _getOneDrive(payload: any): Promise<any> {
        return await fetch(ApiURL.GET_USER_ONEDRIVE, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            const allOneDrive: IOneDrive[] = [];

            data.value.map(drive => {
                const eachDriveItem: IOneDrive = {};
                // eachTemplate.isSelected = false;

                if (drive && drive.folder) {
                    eachDriveItem.DriveName = drive.name;
                    eachDriveItem.DriveId = drive.id;
                    eachDriveItem.DriveUrl = drive.webUrl;
                    eachDriveItem.DriveFolderPath = drive.name;
                    eachDriveItem.DriveLevel = 0;
                    if (drive.parentReference) {
                        eachDriveItem.ParentDriveId = drive.parentReference.driveId;
                        eachDriveItem.ParentDrivePath = drive.parentReference.path;
                    }
                    if (drive.folder) {
                        eachDriveItem.ChildernCount = drive.folder.childCount;
                    }
                    eachDriveItem.ListItemNavigationLink = drive["listItem@odata.navigationLink"];
                    if (drive["#microsoft.graph.createUploadSession"]) {
                        eachDriveItem.UploadSessionLink = drive["#microsoft.graph.createUploadSession"].target;
                    }
                    allOneDrive.push(eachDriveItem);
                }


            });

            return allOneDrive;
        }).catch(error => {
            console.log("Error while fetching Drive Items - ", error);

        });
    }
    private async _getOneDriveChilds(payload: any): Promise<any> {
        return await fetch(ApiURL.GET_USER_ONEDRIVE_HIERARCHY, {
            method: "post",
            body: JSON.stringify(payload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            const allOneDrive: any[] = [];

            data.value.map(drive => {
                const eachDriveItem: IOneDrive = {};
                // eachTemplate.isSelected = false;

                if (drive && drive.folder) {
                    eachDriveItem.DriveName = drive.name;
                    eachDriveItem.DriveId = drive.id;
                    eachDriveItem.DriveUrl = drive.webUrl;
                    if (drive.parentReference) {
                        eachDriveItem.ParentDriveId = drive.parentReference.driveId;
                        eachDriveItem.ParentDrivePath = drive.parentReference.path;
                    }
                    if (drive.folder) {
                        eachDriveItem.ChildernCount = drive.folder.childCount;
                    }
                    eachDriveItem.ListItemNavigationLink = drive["listItem@odata.navigationLink"];
                    if (drive["#microsoft.graph.createUploadSession"]) {
                        eachDriveItem.UploadSessionLink = drive["#microsoft.graph.createUploadSession"].target;
                    }
                    allOneDrive.push(eachDriveItem);
                }


            });

            return allOneDrive;
        }).catch(error => {
            console.log("Error while fetching templates - ", error);

        });
    }
}
