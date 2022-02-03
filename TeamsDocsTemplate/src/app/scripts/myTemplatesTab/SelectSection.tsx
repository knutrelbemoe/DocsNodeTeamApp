import * as React from "react";
import {
    TeamsThemeContext,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    getContext,
    TextArea
} from "msteams-ui-components-react";
import { initializeIcons } from "@uifabric/icons";
import TeamsBaseComponent from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import * as util from "./../utility/formatter";
import { SelectMyLibraryApi } from "./../services";
import { ILocationSelectionState, ILocationSelectionProps, ITeams, ITeamChannel, IChannelFolder, IOneDrive } from "./../Core/interfaces";
import { KeyCodes, createArray, getRTLSafeKeyCode, ISelection, getId } from 'office-ui-fabric-react/lib/Utilities';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { DetailsRow, IColumn, Selection, SelectionMode, DetailsList, CheckboxVisibility, DetailsHeader, IDetailsHeaderProps, DetailsListLayoutMode } from 'office-ui-fabric-react/lib/DetailsList';
import { IPinnedLocations } from "./../Core/interfaces";
import { IconButton, VerticalDivider, Dialog, DialogFooter, DialogType, DefaultButton, PrimaryButton } from "office-ui-fabric-react";
import { Accordion } from "../Core/controls/accordion";
import { PivotItem, IPivotItemProps, Pivot } from 'office-ui-fabric-react/lib/Pivot';
import { IStyleSet } from 'office-ui-fabric-react/lib/Styling';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { Checkbox, ICheckboxProps } from 'office-ui-fabric-react/lib/Checkbox';
import LoadingOverlay from 'react-loading-overlay';
import FadeLoader from 'react-spinners/FadeLoader';
import { IPayload } from "../interfaces/IPayload";
import { getItemClassNames } from "office-ui-fabric-react/lib/components/ContextualMenu/ContextualMenu.classNames";
import { DeleteMyLibraryApi } from "../services/DeleteMyLibraryApi";
import { Spinner, SpinnerSize, SpinnerType } from 'office-ui-fabric-react/lib/Spinner';
import { IStackProps, Stack } from 'office-ui-fabric-react/lib/Stack';
import { String } from "core-js";

const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
    root: { marginTop: 10 }
};
const inputProps: ICheckboxProps['inputProps'] = {
    onFocus: () => console.log('Checkbox is focused'),
    onBlur: () => console.log('Checkbox is blurred'),

};

export class SelectSection extends TeamsBaseComponent<ILocationSelectionProps, ILocationSelectionState> {

    constructor(props: any) {
        super(props, {} as ILocationSelectionState);
        initializeIcons();
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            loading: false,
            fontSize: this.pageFontSize(),
            selectedTemplate: props.selectedTemplate,
            PinnedLocations: props.PinnedLocations,
            Teams: props.Teams,
            OneDrive: props.OneDrive,
            IsLocationPinned: true,
            LoggedInUserName: "",
            HideUnpinDialogBox: true,
            ToBeUnpinnedItemId: -1,
            loadingTeamLocations: true,
            loadingDriveLocations: true
        });

    }

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
    }

    public componentDidMount() {
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);

            const propVals = { ...this.props };


            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId,
                    TeamContext: context,
                    Teams: propVals.Teams,
                    PinnedLocations: propVals.PinnedLocations,
                    OneDrive: propVals.OneDrive
                }, () => {
                    this.loadComponent();
                });
            });

        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }

    }

    public loadTeamUrls(teams: ITeams[], isForTeamShareFolder?: boolean) {
        if (!isForTeamShareFolder) {

            teams.map(async (item, index) => {
                const eachpayload = {
                    tenant: this.state.TeamContext.tid,
                    GroupId: item.TeamId
                };
                try {
                    let teamSiteCollectionUrl: string = await new SelectMyLibraryApi()
                        .getTeamsWebUrl(eachpayload);
                    item.TeamSPOUrl = teamSiteCollectionUrl;

                }
                catch (error) {
                    this.setState({
                        loading: false
                    });
                    console.log("Error while getTeamsWebUrl ", error);
                }
            });

            this.setState({
                Teams: teams
            }, () => {

                this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
            });

        }
        else {

            teams.map(async (item, index) => {
                const eachTeamLibPayload = {
                    tenant: this.state.TeamContext.tid,
                    SPOUrl: `https://${this.state.TeamContext.teamSiteDomain}`,
                    Team: item.TeamSPOUrl
                };
                try {
                    let libArrResp: any[] = await new SelectMyLibraryApi()
                        .getTeamLibraryName(eachTeamLibPayload);



                    if (libArrResp && libArrResp.length > 0) {
                        item.TeamShareFolderName = libArrResp[0].EntityTypeName;
                    }
                }
                catch (error) {
                    this.setState({
                        loading: false
                    });
                    console.log("Error while loading shared document Path ", error);
                }

            });
            this.setState({
                Teams: teams,
                loading: false
            }, () => {

                this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
            });
        }


    }

    public loadComponent() {


        this.setState({
            loading: true,

        });
        const selectMyLibraryApi = new SelectMyLibraryApi();
        const usernamepayload = {
            tenant: this.state.TeamContext.tid,
            SPOUrl: "https://graph.microsoft.com",
            upn: this.state.TeamContext.userPrincipalName
        };


        if (this.state.PinnedLocations && this.state.PinnedLocations.length > 0) {
            this.setState({
                loading: false
            });
        }
        else {
            selectMyLibraryApi
                .getCurrentUserDisplayName(usernamepayload)
                .then((resp) => {

                    this.setState({
                        LoggedInUserName: resp
                    });

                    const payload = {
                        tenant: this.state.TeamContext.tid,
                        SPOUrl: `https://${this.state.TeamContext.teamSiteDomain}`,
                        UserName: resp
                    };
                    selectMyLibraryApi
                        .getMyPinnedLocation(payload)
                        .then((response) => {

                            this.setState({
                                PinnedLocations: response as IPinnedLocations[],
                                loading: false
                            }, () => {

                                this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                            });

                        }).catch(error => {
                            this.setState({
                                loading: false
                            });
                            console.log("Error while loading pinned locations ", error);
                        });
                }).catch(error => {
                    this.setState({
                        loading: false
                    });
                    console.log("Error while loading pinned locations ", error);
                });
        }




        // TODO: Can getting the tenant and SPOURL automated

        const teampayload = {
            tenant: this.state.TeamContext.tid,
            SPOUrl: "https://graph.microsoft.com",
            UsrGUID: this.state.TeamContext && this.state.TeamContext.userObjectId
            // UsrGUID: "d0fd035f-12f9-48dc-84c1-8ac4a43dcc3d"
        };

        const drivepayload = {
            tenant: this.state.TeamContext.tid,
            UsrGUID: this.state.TeamContext && this.state.TeamContext.userObjectId
            // UsrGUID: "d0fd035f-12f9-48dc-84c1-8ac4a43dcc3d"
        };

        if (this.state.Teams && this.state.Teams.length > 0) {
            this.setState({
                loadingTeamLocations: false
            });
        }
        else {
            selectMyLibraryApi
                .getMyMsTeams(teampayload)
                .then((resp) => {
                    const myTeams = resp as ITeams[];
                    this.setState({
                        Teams: myTeams,
                        loadingTeamLocations: false
                    }, () => {

                        this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                    });
                    // Promise.all(
                    //     myTeams.map(async (item, index) => {
                    //         const eachpayload = {
                    //             tenant: this.state.TeamContext.tid,
                    //             GroupId: item.TeamId
                    //         };
                    //         await new SelectMyLibraryApi()
                    //             .getTeamsWebUrl(eachpayload)
                    //             .then((teamurl) => {
                    //                 item.TeamSPOUrl = teamurl;
                    //             }).catch(error => {
                    //                 this.setState({
                    //                     loadingTeamLocations: false
                    //                 });
                    //                 console.log("Error while getTeamsWebUrl ", error);
                    //             });


                    //     })).then(() => {
                    //         this.setState({
                    //             Teams: myTeams

                    //         }, () => {
                    //             console.log("Team Set Complete");
                    //             const { Teams } = this.state;
                    //             if (Teams) {
                    //                 let allTeams = Teams;
                    //                 Promise.all(
                    //                     allTeams.map(async (item, index) => {

                    //                         const eachTeamLibPayload = {
                    //                             tenant: this.state.TeamContext.tid,
                    //                             SPOUrl: `https://${this.state.TeamContext.teamSiteDomain}`,
                    //                             Team: item.TeamSPOUrl
                    //                         };
                    //                         await new SelectMyLibraryApi()
                    //                             .getTeamLibraryName(eachTeamLibPayload)
                    //                             .then((libArrResp: any[]) => {

                    //                                 if (libArrResp && libArrResp.length > 0) {
                    //                                     item.TeamShareFolderName = libArrResp[0].EntityTypeName;
                    //                                 }

                    //                             }).catch(error => {
                    //                                 this.setState({
                    //                                     loadingTeamLocations: false
                    //                                 });
                    //                                 console.log("Error while getting Team Share Path ", error);
                    //                             });


                    //                     })).then(() => {
                    //                         this.setState({
                    //                             Teams: allTeams,
                    //                             loadingTeamLocations: false
                    //                         }, () => {
                    //                             console.log("Team Location Set Complete");
                    //                         });

                    //                     });
                    //             }
                    //         });

                    //     });


                }).catch(error => {
                    this.setState({
                        loadingTeamLocations: false
                    });
                    console.log("Error while loading locations ", error);
                });
        }

        console.log("Start Loading OneDrive");
        if (this.state.OneDrive && this.state.OneDrive.length > 0) {
            this.setState({
                loadingDriveLocations: false
            });
        }
        else {
            selectMyLibraryApi
                .getMyOneDrive(drivepayload)
                .then((resp) => {

                    const myDrive = resp as IOneDrive[];
                    this.setState({
                        OneDrive: myDrive,
                        loadingDriveLocations: false
                    }, () => {

                        this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                    });

                }).catch(error => {
                    this.setState({
                        loadingDriveLocations: false
                    });
                    console.log("Error while loading Drive Location ", error);
                });
        }

    }

    public _onUnPinnedLocation(itemId: number) {


        this.setState({
            loading: true,
            HideUnpinDialogBox: true
        });
        let deleteMyLibraryApi: DeleteMyLibraryApi = new DeleteMyLibraryApi();
        let payload: IPayload = {
            tenant: this.state.TeamContext.tid,
            SPOUrl: `https://${this.state.TeamContext.teamSiteDomain}`,
            ItemID: itemId
        };
        deleteMyLibraryApi.deletePinnedLocation(payload).then((response) => {
            let PinnedLocations: IPinnedLocations[] | undefined = this.state.PinnedLocations ? [...this.state.PinnedLocations] : undefined;
            if (PinnedLocations) {
                let indexOfUnpinnedLocation = PinnedLocations.map(location => location.Id).indexOf(itemId);
                PinnedLocations.splice(indexOfUnpinnedLocation, 1);
            }
            this.setState({
                loading: false,
                PinnedLocations
            }, () => {

                this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
            });
        }).catch(error => {
            this.setState({
                loading: false
            });
            console.error("Error while unpin location", error);
        });

    }
    public _isInnerZoneKeystroke(ev: React.KeyboardEvent<HTMLElement>): boolean {
        return ev.which === getRTLSafeKeyCode(KeyCodes.right);
    }

    public selection = (item: any) => {

        let siteUrl = "";

        if (item.PinnedType.toLowerCase() === "onedrive") {
            siteUrl = "Onedrive";
        }
        else {
            siteUrl = item.LocationUrl;
        }

        this.getSelectSectionDetails({
            DestFolderRelUrl: item.LocationName,
            DestSite: item.LocationUrl,
            LocationAlreadyPinned: true,
            SiteUrl: siteUrl
        });
        // this.getSelectSectionDetails("LocationAlreadyPinned", true);
        // this.getSelectSectionDetails("SiteUrl", item.LocationUrl);

    }

    public onChange = (ev: React.FormEvent<HTMLElement>, isChecked: boolean) => {

        const { IsLocationPinned } = this.state;
        this.getSelectSectionDetails({ IsPinnedLocation: isChecked });
        this.setState({
            IsLocationPinned: isChecked
        });


    }

    public _onRenderHeader(detailsHeaderProps: IDetailsHeaderProps) {
        return (
            null
        );
    }

    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });

        const COLUMNS: IColumn[] = [

            {
                key: "link",
                name: 'Link',
                fieldName: '',
                minWidth: 250,
                maxWidth: 500,
                className: "pinnedFolder",
                onRender: item => <div key={item.id} onClick={ev => (this.selection(item))} ><Label className="folderLabel">{item.LocationName}</Label><div><Link href={item.LocationUrl} target="_blank">{item.LocationUrl}</Link></div></div>
            },
            {
                key: 'IconButton',
                name: 'Link',
                fieldName: '',
                minWidth: 130,
                className: "pinnedIcon",
                onRender: item =>
                    <IconButton
                        iconProps={{ iconName: 'Pinned' }}
                        onClick={(e) =>
                            this.setState({
                                ToBeUnpinnedItemId: item.Id,
                                HideUnpinDialogBox: false
                            })
                        }
                        className="locationpin"
                    />
            }

        ];
        const rowProps: IStackProps = { horizontal: true, verticalAlign: 'center' };

        const tokens = {
            sectionStack: {
                childrenGap: 10
            },
            spinnerStack: {
                childrenGap: 20
            }
        };

        let TeamPinnedLocation: any[] = [];
        let SPPinnedLocation: any[] = [];
        let OneDrivePinnedLocation: any[] = [];
        if (this.state.PinnedLocations) {
            this.state.PinnedLocations.map((item, index) => {
                if (item.PinnedType && item.PinnedType.toLowerCase() === "teams") {
                    // tslint:disable-next-line: no-unused-expression
                    TeamPinnedLocation.push(item);

                }
                if (item.PinnedType && item.PinnedType.toLowerCase() === "documentlibrary") {
                    // tslint:disable-next-line: no-unused-expression
                    SPPinnedLocation.push(item);
                }
                if (item.PinnedType && item.PinnedType.toLowerCase() === "onedrive") {
                    // tslint:disable-next-line: no-unused-expression
                    OneDrivePinnedLocation.push(item);

                }
            });
        }

        const { HideUnpinDialogBox } = this.state;
        // tslint:disable-next-line: variable-name
        const _labelId = getId('dialogLabel');
        // tslint:disable-next-line: variable-name
        const _subTextId = getId('subTextLabel');

        return (
            <TeamsThemeContext.Provider value={context} >
                <Dialog
                    hidden={HideUnpinDialogBox}
                    onDismiss={() => this.setState({ HideUnpinDialogBox: true })}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Are you Sure',
                        closeButtonAriaLabel: 'Close',
                        subText: 'Do you want to unpin the location ?'
                    }}
                    modalProps={{
                        titleAriaId: _labelId,
                        subtitleAriaId: _subTextId,
                        isBlocking: false,
                        styles: { main: { maxWidth: 450 } },
                    }}
                >
                    <DialogFooter>
                        <PrimaryButton
                            onClick={() => this._onUnPinnedLocation(this.state.ToBeUnpinnedItemId)}
                            text="OK" menuIconProps={{ iconName: 'Accept' }}
                        />
                        <DefaultButton
                            onClick={() => this.setState({ HideUnpinDialogBox: true })}
                            text="Cancel" menuIconProps={{ iconName: 'Clear' }}
                        />
                    </DialogFooter>
                </Dialog>
                <Surface className="tabContainer">
                    <Panel>
                        <PanelHeader>
                        </PanelHeader>
                        <LoadingOverlay
                            active={this.state.loading}
                            spinner={<FadeLoader />}
                            text="Loading Locations...."
                        >
                            <PanelBody className="container scrollable">
                                <Label className="sectionHeader">Pinned Locations</Label>
                                <div className="card">

                                    <FocusZone direction={FocusZoneDirection.vertical} isCircularNavigation={true} isInnerZoneKeystroke={this._isInnerZoneKeystroke} role="grid">
                                        <Pivot>
                                            <PivotItem headerText="Microsoft Teams" itemIcon="TeamsLogoInverse">
                                                <DetailsList
                                                    items={TeamPinnedLocation}
                                                    setKey="Team"
                                                    columns={COLUMNS}
                                                    onRenderDetailsHeader={this._onRenderHeader}
                                                    ariaLabelForSelectionColumn="Toggle selection"
                                                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                                    checkboxVisibility={CheckboxVisibility.hidden}
                                                    selectionMode={SelectionMode.single}
                                                />
                                            </PivotItem>

                                            <PivotItem headerText="OneDrive" itemIcon="OneDriveFolder16">
                                                <DetailsList
                                                    items={OneDrivePinnedLocation}
                                                    setKey="Onedrive"
                                                    columns={COLUMNS}
                                                    onRenderDetailsHeader={this._onRenderHeader}
                                                    ariaLabelForSelectionColumn="Toggle selection"
                                                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                                    checkboxVisibility={CheckboxVisibility.hidden}
                                                    selectionMode={SelectionMode.single}
                                                />
                                            </PivotItem>
                                            <PivotItem headerText="SharePoint" itemIcon="SharepointAppIcon16">
                                                <DetailsList
                                                    items={SPPinnedLocation}
                                                    setKey="SP"
                                                    columns={COLUMNS}
                                                    onRenderDetailsHeader={this._onRenderHeader}
                                                    ariaLabelForSelectionColumn="Toggle selection"
                                                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                                                    checkboxVisibility={CheckboxVisibility.hidden}
                                                    selectionMode={SelectionMode.single}
                                                />
                                            </PivotItem>
                                        </Pivot>
                                    </FocusZone>
                                </div>
                                <Label className="sectionHeader">All Locations</Label>
                                <div className="card">
                                    <Pivot onLinkClick={(item?: PivotItem) => this.setOnedriveRootasDestination(item)}>
                                        <PivotItem headerText="Microsoft Teams" itemIcon="TeamsLogoInverse">
                                            <FocusZone direction={FocusZoneDirection.vertical}>

                                                {!this.state.loadingTeamLocations &&
                                                    <div className="channelArea">
                                                        {this.state.Teams && this.state.Teams.map((item: ITeams, index: number) => {
                                                            return this._onRenderTeams(item, index);
                                                        })}
                                                    </div>
                                                }
                                                {this.state.loadingTeamLocations &&
                                                    <div className="channelArea">
                                                        <Stack tokens={tokens.sectionStack}>
                                                            <Stack {...rowProps} tokens={tokens.spinnerStack}></Stack>
                                                            <Spinner label="loading teams Channels..." size={SpinnerSize.large} type={SpinnerType.normal} />
                                                        </Stack>
                                                    </div>}
                                            </FocusZone>
                                        </PivotItem>
                                        <PivotItem headerText="OneDrive" itemIcon="OneDriveFolder16" >
                                            <FocusZone direction={FocusZoneDirection.vertical}>
                                                {!this.state.loadingDriveLocations && <div className="channelArea">
                                                    {this.state.OneDrive && this.state.OneDrive.map((item: IOneDrive, index: number) => {
                                                        return this._onRenderDrive(item, index);
                                                    })}
                                                </div>
                                                }
                                                {this.state.loadingDriveLocations &&
                                                    <div className="channelArea">
                                                        <Stack tokens={tokens.sectionStack}>
                                                            <Stack {...rowProps} tokens={tokens.spinnerStack}></Stack>
                                                            <Spinner label="loading Onedrive Folders..." size={SpinnerSize.large} type={SpinnerType.normal} />
                                                        </Stack>
                                                    </div>}
                                            </FocusZone>
                                        </PivotItem>
                                        <PivotItem headerText="SharePoint" itemIcon="SharepointAppIcon16">
                                            <FocusZone direction={FocusZoneDirection.vertical}>
                                                <Label styles={labelStyles}>Coming Soon....</Label>
                                            </FocusZone>
                                        </PivotItem>
                                    </Pivot>

                                    <Checkbox className="checkBoxArea" label="Make this location as Pinned location" onChange={this.onChange} checked={this.state.IsLocationPinned} />


                                </div>


                            </PanelBody>
                        </LoadingOverlay>
                        <PanelFooter>

                        </PanelFooter>
                    </Panel>
                </Surface>
            </TeamsThemeContext.Provider >
        );
    }



    public _onExpandTeams(team: ITeams) {
        const { Teams } = this.state;
        const eachpayload = {
            tenant: this.state.TeamContext.tid,
            GroupId: team.TeamId
        };
        if (!team.TeamSPOUrlLoaded) {
            new SelectMyLibraryApi()
                .getTeamsWebUrl(eachpayload)
                .then((teamurl) => {
                    team.TeamSPOUrl = teamurl;

                    if (Teams) {
                        Teams.map((item: ITeams, idx: number) => {
                            if (item.TeamId === team.TeamId) {
                                item.TeamSPOUrl = teamurl;
                                item.TeamSPOUrlLoaded = true;
                            }
                        });

                        this.setState({
                            Teams: Teams
                        }, () => {

                            this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                        });
                    }
                    if (!team.TeamSharedFolderLoaded) {

                        const eachTeamLibPayload = {
                            tenant: this.state.TeamContext.tid,
                            SPOUrl: `https://${this.state.TeamContext.teamSiteDomain}`,
                            Team: team.TeamSPOUrl
                        };
                        new SelectMyLibraryApi()
                            .getTeamLibraryName(eachTeamLibPayload)
                            .then((libArrResp: any[]) => {

                                if (libArrResp && libArrResp.length > 0) {
                                    team.TeamShareFolderName = libArrResp[0].EntityTypeName;

                                    if (Teams) {
                                        Teams.map((item: ITeams, idx: number) => {
                                            if (item.TeamId === team.TeamId) {
                                                item.TeamShareFolderName = team.TeamShareFolderName;
                                                item.TeamSharedFolderLoaded = true;
                                            }
                                        });


                                        this.setState({
                                            Teams: Teams
                                        }, () => {
                                            const payload = {
                                                tenant: this.state.TeamContext.tid,
                                                SPOUrl: `https://${this.state.TeamContext.teamSiteDomain}`,
                                                TeamID: team.TeamId,
                                                TeamURL: team.TeamSPOUrl,
                                                LibInternalName: team.TeamShareFolderName

                                            };

                                            this.getSelectSectionDetails({ InvalidLocation: false });

                                            if (!team.TeamChannelLoaded) {

                                                new SelectMyLibraryApi()
                                                    .getMyMsTeamChannels(payload)
                                                    .then((resp) => {
                                                        if (resp && Teams) {
                                                            Teams.map((item: ITeams, idx: number) => {
                                                                if (item.TeamId === team.TeamId) {
                                                                    item.TeamChannels = resp;
                                                                    item.TeamChannelLoaded = true;

                                                                }
                                                            });
                                                        }

                                                    }).finally(() => {
                                                        this.setState({
                                                            Teams: Teams
                                                        }, () => {

                                                            this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                                                        });

                                                    });
                                            }
                                            // this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                                        });
                                    }
                                }

                            }).catch(error => {
                                console.log("Error while getting Team Share Path ", error);
                            });
                    }

                });
        }


    }
    public _onExpandDrive(drive: IOneDrive) {

        const payload = {
            tenant: this.state.TeamContext.tid,
            UsrGUID: this.state.TeamContext && this.state.TeamContext.userObjectId,
            ItemID: drive.DriveId

        };
        this.getSelectSectionDetails({
            DestFolderRelUrl: drive.DriveFolderPath,
            DestSite: drive.DriveUrl,
            SiteUrl: "Onedrive"
        });
        const { OneDrive } = this.state;

        if (!drive.DriveChannelLoaded) {

            new SelectMyLibraryApi()
                .getMyOneDriveChilds(payload)
                .then((resp: IOneDrive[]) => {
                    if (OneDrive) {

                        OneDrive.map((item: IOneDrive, idx: number) => {
                            if (item.DriveId === drive.DriveId) {
                                resp.map((driveItem: IOneDrive) => {
                                    driveItem.DriveFolderPath = drive.DriveFolderPath + "/" + driveItem.DriveName;
                                });

                                item.DriveChannels = resp;
                                item.DriveChannelLoaded = true;
                                item.ISSelectable = true;
                                item.DriveLevel = (drive.DriveLevel ? drive.DriveLevel : 0) + 1;
                                item.ChildCount = 0;
                                if (resp.length > 0) {
                                    item.ChildCount = resp.length;
                                }
                            }
                        });
                    }

                }).finally(() => {
                    this.setState({
                        OneDrive: OneDrive
                    }, () => {

                        this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                    });

                });
        }

    }

    public _onExpandChannel(channel: ITeamChannel, team: ITeams, level: number) {

        let channelInnerName = channel.ChannelUrl && channel.ChannelUrl.substring(channel.ChannelUrl.lastIndexOf("/") + 1, channel.ChannelUrl.lastIndexOf("?"));
        channelInnerName = channelInnerName && channelInnerName.replace(/\+/gi, "%20");
        const payload = {
            tenant: this.state.TeamContext.tid,
            SPOUrl: `https://${this.state.TeamContext.teamSiteDomain}`,
            FolderRelPath: channel.FolderType === "sharepoint.folder" ? channel.ChannelUrl : team.TeamShareFolderName + "/" + channelInnerName,
            SiteUrl: team.TeamSPOUrl,
            ChannelID: channel.ChannelId,
            TeamID: team.TeamId,
            FolderType: channel.FolderType === "sharepoint.folder" ? channel.FolderType : "#microsoft.graph.channel"

        };

        if (channel.FolderType && channel.FolderType.includes("sharepoint.folder")) {
            if (channel.ChannelUrl) {

                let relPath = channel.ChannelUrl && channel.ChannelUrl;

                let siteUrl = team.TeamSPOUrl;

                this.getSelectSectionDetails({
                    DestFolderRelUrl: relPath,
                    DestSite: siteUrl,
                    SiteUrl: siteUrl,
                });
            }
        }
        else {
            this.getSelectSectionDetails({
                DestFolderRelUrl: team.TeamShareFolderName + "/" + channelInnerName,
                DestSite: team.TeamSPOUrl,
                SiteUrl: team.TeamSPOUrl
            });

        }
        // this.getSelectSectionDetails("SiteUrl", team.TeamSPOUrl);

        const { Teams } = this.state;

        if (!channel.FolderList || (channel.FolderList && channel.FolderList.length === 0)) {

            new SelectMyLibraryApi()
                .getTeamsChannelSubFolders(payload, level)
                .then((resp) => {
                    if (Teams) {
                        Teams.map((item: ITeams, idx: number) => {
                            if (item.TeamId === team.TeamId && item.TeamChannels) {
                                item.TeamChannels.map((channelItem: ITeamChannel, idxChannel: number) => {
                                    if (channelItem.ChannelId === channel.ChannelId) {
                                        channelItem.FolderList = resp;
                                        channelItem.ISSelectable = true;
                                        channelItem.FolderCount = channelItem.FolderList ? channelItem.FolderList.length : 0;
                                    }
                                });

                            }
                        });
                    }

                }).finally(() => {

                    const tabPayload = {
                        tenant: this.state.TeamContext.tid,
                        ChannelID: channel.ChannelId,
                        TeamID: team.TeamId
                    };

                    new SelectMyLibraryApi()
                        .getTeamsChannelTabs(tabPayload, level)
                        .then((resp: any[]) => {
                            if (Teams) {
                                Teams.map((item: ITeams, idx: number) => {
                                    if (item.TeamId === team.TeamId && item.TeamChannels) {
                                        item.TeamChannels.map((channelItem: ITeamChannel, idxChannel: number) => {
                                            if (channelItem.ChannelId && channel.ChannelId && channelItem.ChannelId.toString() === channel.ChannelId.toString()) {
                                                if (channelItem.FolderList) {
                                                    if (resp && resp.length > 0) {
                                                        resp.map(foldItem => {
                                                            if (channelItem.FolderList) {
                                                                channelItem.FolderList.push(foldItem);
                                                            }

                                                        });
                                                    }

                                                }
                                                else {
                                                    channelItem.FolderList = resp;
                                                }
                                                channelItem.ISSelectable = true;
                                                channelItem.FolderCount = channelItem.FolderList ? channelItem.FolderList.length : 0;
                                            }
                                        });

                                    }
                                });
                            }
                            else {
                                this.setState({
                                    Teams: Teams
                                }, () => {
                                    this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);

                                });
                            }

                        }).finally(() => {
                            this.setState({
                                Teams: Teams
                            }, () => {

                                this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                            });

                        });
                });

        }

    }
    public _onExpandDriveChild(channel: IOneDrive, drive: IOneDrive) {

        // const channelInnerName = channel.ChannelUrl && channel.ChannelUrl.substring(channel.ChannelUrl.lastIndexOf("/") + 1, channel.ChannelUrl.lastIndexOf("?"));

        const payload = {
            tenant: this.state.TeamContext.tid,
            UsrGUID: this.state.TeamContext && this.state.TeamContext.userObjectId,
            ItemID: channel.DriveId

        };

        this.getSelectSectionDetails({
            DestFolderRelUrl: channel.DriveFolderPath,
            DestSite: channel.DriveUrl,
            SiteUrl: "Onedrive"
        });
        // this.getSelectSectionDetails("SiteUrl", team.TeamSPOUrl);

        const { OneDrive } = this.state;

        if (!channel.DriveChannels || (channel.DriveChannels && channel.DriveChannels.length === 0)) {

            new SelectMyLibraryApi()
                .getMyOneDriveChilds(payload)
                .then((resp: IOneDrive[]) => {
                    if (OneDrive) {
                        OneDrive.map((item: IOneDrive, idx: number) => {
                            if (item.DriveId === drive.DriveId && item.DriveChannels) {

                                item.DriveChannels.map((channelItem: IOneDrive, idxChannel: number) => {

                                    if (channelItem.DriveId === channel.DriveId) {

                                        resp.map((driveItem: IOneDrive) => {
                                            driveItem.DriveFolderPath = channel.DriveFolderPath + "/" + driveItem.DriveName;
                                        });
                                        channelItem.DriveChannels = resp;
                                        channelItem.DriveChannelLoaded = true;
                                        channelItem.ISSelectable = true;
                                        channelItem.DriveLevel = (drive.DriveLevel ? drive.DriveLevel : 0) + 1;
                                        channelItem.ChildCount = 0;
                                        if (resp.length > 0) {
                                            channelItem.ChildCount = resp.length;
                                        }
                                    }
                                    else {
                                        let rootlevel: number = channelItem.DriveLevel ? channelItem.DriveLevel : 0;
                                        this._setSubDriveFolder(channelItem, channel, resp, rootlevel, 1);
                                    }
                                });

                            }
                        });
                    }

                }).finally(() => {
                    this.setState({
                        OneDrive: OneDrive
                    }, () => {

                        this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                    });

                });
        }

    }
    public _onExpandFolder(folder: IChannelFolder, channel: ITeamChannel, team: ITeams, folderlevel: number) {

        let channelInnerName = channel.ChannelUrl && channel.ChannelUrl.substring(channel.ChannelUrl.lastIndexOf("/") + 1, channel.ChannelUrl.lastIndexOf("?")) || '';
        channelInnerName = channelInnerName && channelInnerName.replace(/\+/gi, "%20");

        let folderRelPath = '';

        folderRelPath = folder.FolderRelativeUrl && folder.FolderRelativeUrl.substr(folder.FolderRelativeUrl.indexOf(channelInnerName)) || '';

        let payload = {
            tenant: this.state.TeamContext.tid,
            SPOUrl: `https://${this.state.TeamContext.teamSiteDomain}`,
            FolderRelPath: ((folder.FolderType && folder.FolderType.includes("sharepoint.folder")) || channel.FolderType === "sharepoint.folder") ? folder.FolderUrl : team.TeamShareFolderName + "/" + folderRelPath,
            SiteUrl: team.TeamSPOUrl,
            ChannelID: channel.ChannelId,
            TeamID: team.TeamId,
            FolderType: (folder.FolderType && folder.FolderType.includes("sharepoint.folder")) ? folder.FolderType : channel.FolderType

        };

        if (folder.FolderType && (folder.FolderType.includes("sharepoint.folder") || folder.FolderType.includes("#microsoft.graph.channel.private"))) {
            if (folder.FolderUrl) {
                let splitString = folder.FolderUrl.split("/");
                let teamName = splitString[4] && splitString[4];

                let relPath = folder.FolderRelativeUrl && folder.FolderRelativeUrl;

                let siteUrl = `https://${this.state.TeamContext.teamSiteDomain}` + "/sites/" + teamName;

                this.getSelectSectionDetails({
                    DestFolderRelUrl: relPath,
                    DestSite: siteUrl,
                    SiteUrl: siteUrl,
                });

                payload = {
                    tenant: this.state.TeamContext.tid,
                    SPOUrl: `https://${this.state.TeamContext.teamSiteDomain}`,
                    FolderRelPath: folder.FolderRelativeUrl,
                    SiteUrl: siteUrl,
                    ChannelID: channel.ChannelId,
                    TeamID: team.TeamId,
                    FolderType: folder.FolderType

                };
            }
        }
        else {
            this.getSelectSectionDetails({
                DestFolderRelUrl: team.TeamShareFolderName + "/" + folderRelPath,
                DestSite: team.TeamSPOUrl,
                SiteUrl: team.TeamSPOUrl,
            });
        }
        // this.getSelectSectionDetails("SiteUrl", team.TeamSPOUrl);
        // if (folder.SubFolderLoaded) {
        const { Teams } = this.state;


        new SelectMyLibraryApi()
            .getTeamsChannelSubFolders(payload, folderlevel + 1)
            .then((resp) => {
                if (Teams) {
                    Teams.map((item: ITeams, idx: number) => {
                        if (item.TeamId === team.TeamId && item.TeamChannels) {
                            item.TeamChannels.map((channelItem: ITeamChannel, idxChannel: number) => {
                                if (channelItem.ChannelId === channel.ChannelId && channelItem.FolderList) {
                                    channelItem.FolderList.map((chnfoldr: IChannelFolder) => {
                                        if (chnfoldr.FolderLevel === folderlevel && chnfoldr.FolderRelativeUrl === folder.FolderRelativeUrl) {
                                            chnfoldr.SubFolders = resp;
                                            chnfoldr.ISSelectable = true;
                                            chnfoldr.SubFolderLoaded = true;
                                        } else {
                                            this._setSubFolder(channel, folder, chnfoldr, resp, folderlevel, 2);
                                        }

                                    });

                                }
                            });

                        }
                    });
                }

            }).finally(() => {
                this.setState({
                    Teams: Teams
                }, () => {

                    this.setSelectionLocationState(this.state.PinnedLocations, this.state.Teams, this.state.OneDrive);
                });

            });
        // }


    }
    public getSelectSectionDetails = (obj: any) => {
        this.props.getSelectSectionDetails(obj);
        console.log("-------Save Location-------");
        console.log(obj);
    }

    public setSelectionLocationState = (pinnedLocations: any, teamLocations: any, driveLocations: any) => {
        this.props.setSelectionLocationState(pinnedLocations, teamLocations, driveLocations);
    }

    public setOnedriveRootasDestination = (item?: PivotItem) => {

        if (item && item.props.headerText === "OneDrive") {

            const { OneDrive } = this.state;
            if (OneDrive && OneDrive.length > 0) {
                this.getSelectSectionDetails({
                    DestFolderRelUrl: "/",
                    DestSite: OneDrive[0].DriveUrl,
                    SiteUrl: "Onedrive"
                });
            } else {
                this.getSelectSectionDetails({ InvalidLocation: false });
            }
        }
        else {
            this.getSelectSectionDetails({ InvalidLocation: false });
        }


    }

    private _getSubFolder(folder: IChannelFolder, folderLevel: number, initialLevel: number): IChannelFolder[] | undefined {

        for (let i = initialLevel; i <= folderLevel; i++) {
            if (i === folderLevel) {
                return folder.SubFolders || [];
            }
            else {

                if (folder.SubFolders) {
                    folder.SubFolders.map((subfoldr, inx) => {
                        return this._getSubFolder(subfoldr, folderLevel, i);
                    });
                }


            }



        }
    }

    private _getOneDrives(parentFolder: IOneDrive, drive: IOneDrive, apiResp: IOneDrive[]): IOneDrive | undefined {
        let driveFound: number = 0;

        if (parentFolder && parentFolder.DriveChannels) {
            parentFolder.DriveChannels.map((item: IOneDrive, idx: number) => {
                if (item.DriveId === drive.DriveId && item.DriveChannels) {

                    driveFound += 1;
                    apiResp.map((driveItem: IOneDrive) => {
                        driveItem.DriveFolderPath = parentFolder.DriveFolderPath + "/" + driveItem.DriveName;
                    });
                    parentFolder.DriveChannels = apiResp;
                    parentFolder.DriveChannelLoaded = true;
                    parentFolder.ISSelectable = true;
                    return parentFolder;
                }

                if (driveFound === 0 && item.DriveChannels) {
                    item.DriveChannels.map((channelItem: IOneDrive, idxChannel: number) => {
                        return this._getOneDrives(channelItem, drive, apiResp);
                    });
                }
            });
        }
        else {
            return {};
        }

    }
    private _setSubFolder(rootchannel: ITeamChannel, rootfolderList: IChannelFolder, initialFolderlist: IChannelFolder, respFolderList: IChannelFolder[], folderLevel: number, initialLevel: number): IChannelFolder | undefined {


        for (let i = initialLevel; i <= folderLevel; i++) {
            if (initialLevel === folderLevel && initialFolderlist.SubFolders) {

                initialFolderlist.SubFolders.map((subfoldr, inx) => {
                    if (subfoldr.FolderRelativeUrl === rootfolderList.FolderRelativeUrl) {
                        subfoldr.SubFolders = respFolderList;
                    }
                });

                return initialFolderlist;
            }
            else {

                if (initialFolderlist.SubFolders) {
                    initialFolderlist.SubFolders.map((subfoldr, inx) => {
                        return this._setSubFolder(rootchannel, rootfolderList, subfoldr, respFolderList, folderLevel, i);
                    });
                }


            }



        }
    }
    private _setSubDriveFolder(rootchannel: IOneDrive, initialFolderlist: IOneDrive, respFolderList: IOneDrive[], folderLevel: number, initialLevel: number): IOneDrive | undefined {


        for (let i = initialLevel; i <= folderLevel; i++) {
            if (i === folderLevel && rootchannel.DriveChannels) {

                rootchannel.DriveChannels.map((subfoldr, inx) => {
                    if (subfoldr.DriveId === initialFolderlist.DriveId) {
                        respFolderList.map((driveItem: IOneDrive) => {
                            driveItem.DriveFolderPath = initialFolderlist.DriveFolderPath + "/" + driveItem.DriveName;
                        });

                        subfoldr.DriveChannels = respFolderList;
                        subfoldr.DriveChannelLoaded = true;
                        subfoldr.ISSelectable = true;
                        subfoldr.DriveLevel = (rootchannel.DriveLevel ? rootchannel.DriveLevel : 0) + 1;
                        subfoldr.ChildCount = 0;
                        if (respFolderList.length > 0) {
                            subfoldr.ChildCount = respFolderList.length;
                        }
                    }
                });

                return initialFolderlist;
            }
            else {

                if (rootchannel.DriveChannels) {
                    rootchannel.DriveChannels.map((subfoldr, inx) => {
                        let rootlevel: number = subfoldr.DriveLevel ? subfoldr.DriveLevel : 0;
                        return this._setSubDriveFolder(subfoldr, initialFolderlist, respFolderList, rootlevel, i);
                    });
                }


            }



        }
    }
    private _onRenderFolder(item: IChannelFolder, index: number | undefined, team: ITeams, channel: ITeamChannel): JSX.Element {

        return (
            <Accordion title={item.FolderName || ''} id={item.FolderName || ''}
                defaultCollapsed={true} key={index} onExpand={() => this._onExpandFolder(item, channel, team, item.FolderLevel || 0)}
                childCount={item.ChildCount} foldertype={item.FolderType}
            >
                {item.SubFolders && item.SubFolders.map((folder: IChannelFolder, folderindex: number) => {
                    return this._onRenderFolder(folder, folderindex, team, channel);
                })

                }
            </Accordion>

        );
    }
    private _onRenderOnedriveFolder(item: IOneDrive, index: number | undefined, drive: IOneDrive, channel: IOneDrive): JSX.Element {

        return (
            <Accordion title={item.DriveName || ''} id={item.DriveId || ''}
                defaultCollapsed={true} key={index} onExpand={() => this._onExpandDriveChild(item, drive)}
                childCount={item.ChildCount}
            >
                {item.DriveChannels && item.DriveChannels.map((folder: IOneDrive, folderindex: number) => {
                    return this._onRenderOnedriveFolder(folder, folderindex, drive, channel);
                })

                }
            </Accordion>

        );
    }
    private _onRenderChannels(item: ITeamChannel, index: number | undefined, team: ITeams): JSX.Element {

        return (

            <Accordion title={item.ChannelDisplayName || ''} id={item.ChannelId || ''}
                defaultCollapsed={true} key={index} onExpand={() => this._onExpandChannel(item, team, 1)}
                childCount={item.FolderCount} foldertype={item.FolderType}
            >
                {item.FolderList && item.FolderList.map((folder: IChannelFolder, folderindex: number) => {
                    return this._onRenderFolder(folder, folderindex, team, item);
                })

                }
            </Accordion>

        );
    }
    private _onRenderDriveChilds(item: IOneDrive, index: number | undefined, drive: IOneDrive): JSX.Element {

        return (

            <Accordion title={item.DriveName || ''} id={item.DriveId || ''}
                defaultCollapsed={true} key={index} onExpand={() => this._onExpandDriveChild(item, drive)}
                childCount={item.ChildCount}
            >
                {item.DriveChannels && item.DriveChannels.map((folder: IOneDrive, folderindex: number) => {
                    return this._onRenderOnedriveFolder(folder, folderindex, drive, item);
                })

                }
            </Accordion>

        );
    }

    private _onRenderTeams(item: ITeams, index: number | undefined): JSX.Element {
        const { Teams } = this.state;
        return (
            <Accordion title={item.TeamName || ''} id={item.TeamId || ''}
                defaultCollapsed={true} key={index} onExpand={() => this._onExpandTeams(item)}>
                {item.TeamChannels && item.TeamChannels.map((channel: ITeamChannel, channelindex: number) => {
                    return this._onRenderChannels(channel, channelindex, item);
                })

                }
            </Accordion>
        );
    }

    private _onRenderDrive(item: IOneDrive, index: number | undefined): JSX.Element {
        const { OneDrive } = this.state;
        return (
            <Accordion title={item.DriveName || ''} id={item.DriveId || ''}
                defaultCollapsed={true} key={index} onExpand={() => this._onExpandDrive(item)}
                childCount={item.ChildCount}
            >
                {item.DriveChannels && item.DriveChannels.map((channel: IOneDrive, channelindex: number) => {
                    return this._onRenderDriveChilds(channel, channelindex, item);
                })

                }
            </Accordion>
        );
    }
}
