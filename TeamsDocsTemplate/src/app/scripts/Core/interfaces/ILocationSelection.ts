import { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { IPinnedLocations, ITeams, ITeamChannel, IChannelFolder } from "./../interfaces";
import { IOneDrive } from "./IAllTeamsChannels";

export interface ILocationSelectionState extends ITeamsBaseComponentState {
    loading: boolean;
    entityId?: string;
    LocationUrl?: string;
    LocationName?: string;
    selectedTemplate?: any[];
    TeamContext?: microsoftTeams.Context | any;
    PinnedLocations?: IPinnedLocations[];
    Teams?: ITeams[];
    IsLocationPinned?: boolean;
    LoggedInUserName?: string;
    OneDrive?: IOneDrive[];
    HideUnpinDialogBox?: boolean;
    ToBeUnpinnedItemId: number;
    IsSelectionPinnable?: boolean;
    loadingDriveLocations?: boolean;
    loadingTeamLocations?: boolean;

}


export interface ILocationSelectionProps extends ITeamsBaseComponentProps {
    selectedTemplate?: any[];
    getSelectSectionDetails: (obj: any) => any;
    setSelectionLocationState: (pinnedLocations: any, teamLocations: any, driveLocations: any) => any;
    PinnedLocations?: any;
    Teams?: any;
    OneDrive?: any;
}
