export interface ITeamChannel {

    ChannelUrl?: string;
    ChannelDisplayName?: string;
    ChannelId?: string;
    FolderList?: IChannelFolder[];
    ISSelectable?: boolean;
    selected?: boolean;
    FolderCount?: number;
    FolderType?: string;
}

export interface IChannelFolder {
    FolderName?: string;
    FolderUrl?: string;
    FolderRelativeUrl?: string;
    SubFolders?: IChannelFolder[];
    FolderLevel?: number;
    SubFolderLoaded?: boolean;
    ISSelectable?: boolean;
    selected?: boolean;
    ChildCount?: number;
    FolderType?: string;

}
export interface ITeams {
    TeamUrl?: string;
    TeamId?: string;
    TeamName?: string;
    TeamSPOUrl?: string;
    TeamChannels?: ITeamChannel[];
    TeamChannelLoaded?: boolean;
    TeamShareFolderName?: string;
    TeamNameLoaded?: boolean;
    TeamSPOUrlLoaded?: boolean;
    TeamSharedFolderLoaded?: boolean;

}
export interface IOneDrive {
    DriveUrl?: string;
    DriveId?: string;
    DriveFolderPath?: string;
    DriveChannels?: IOneDrive[];
    DriveChannelLoaded?: boolean;
    DriveShareFolderName?: string;
    DriveName?: string;
    ParentDriveId?: string;
    ParentDrivePath?: string;
    ChildernCount?: number;
    ListItemNavigationLink?: string;
    UploadSessionLink?: string;
    ISSelectable?: boolean;
    selected?: boolean;
    DriveLevel?: number;
    ChildCount?: number;
}
