import * as React from "react";
// import styles from "../../css/DocTemplate.scss";
import {
    PrimaryButton,
    TeamsThemeContext,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface,
    getContext,
    TextArea,
    Checkbox
} from "msteams-ui-components-react";
import { Modal, mergeStyleSets, IconButton, getTheme, DefaultButton, DialogType, Dialog, getId } from 'office-ui-fabric-react';
import { SearchBox } from "office-ui-fabric-react/lib/SearchBox";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import LoadingOverlay from 'react-loading-overlay';
import FadeLoader from 'react-spinners/FadeLoader';
import { Icon, IIconProps } from "office-ui-fabric-react/lib/Icon";
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption, IDropdownProps } from "office-ui-fabric-react/lib/Dropdown";
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import { setup as pnpSetup } from "@pnp/common";
import { graph, Section } from "@pnp/graph";
import { CSSTransition } from 'react-transition-group';
import { ContextualMenuItemType, IContextualMenuItemProps } from 'office-ui-fabric-react/lib/ContextualMenu';
// import { PreviewSection } from "./PreviewSection";
import { PreviewSectionPopup } from "./PreviewSectionPopup";
import { NextSection } from "./NextSection";
import { SelectSection } from "./SelectSection";
import { CreateSection } from "./CreateSection";
import "./../../scss/styles.scss";
import { initializeIcons } from '@uifabric/icons';
import { TEMPLATE_IMAGE_PATH, GET_TEMPLATE_URL, LOAD_IMAGE_INIT_URL, DEFAULT_DOCX_URL, DEFAULT_PPTX_URL, DEFAULT_XLSX_URL, DEFAULT_SITE_URL, GET_TENANT_FULLNAME } from './../utility/Constants';
import * as util from "./../utility/formatter";
import { IBreadCrumb } from "../interfaces/IBreadCrumb";
import { IFolderBreadCrumb } from "../interfaces/IFolderBreadCrumb";
import { IFileFolder } from "../interfaces/IFileFolder";
import { IMyTemplateTabAPI } from "../services/MyTemplateTabAPI";
import { MyTemplateTabAPIImpl } from "../services/MyTemplateTabAPIImpl";
import { ILocationPayload } from "../interfaces/ILocationPayload";

const theme = getTheme();
const filterIcon: IIconProps = { iconName: 'Filter' };

const contentStyles = mergeStyleSets({
    container: {
        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',
        width: "60em"
    },
    header: [
        theme.fonts.xLargePlus,
        {
            flex: '1 1 auto',
            borderTop: `4px solid ${theme.palette.themePrimary}`,
            color: theme.palette.neutralPrimary,
            display: 'flex',
            fontSize: "1.5em",
            alignItems: 'center',
            fontWeight: 700,
            padding: '12px 12px 14px 24px'
        }
    ],
    body: {
        flex: '4 4 auto',
        padding: '0 24px 24px 24px',
        overflowY: 'hidden',
        selectors: {
        }
    }
});

const iconButtonStyles = mergeStyleSets({
    root: {
        color: theme.palette.neutralPrimary,
        marginLeft: 'auto',
        marginTop: '4px',
        marginRight: '2px'
    },
    rootHovered: {
        color: theme.palette.neutralDark
    }
});

const options: IDropdownOption[] = [
    { key: "List", text: "All Documents", data: { icon: "List" } },
    { key: "Tiles", text: "All Documents", data: { icon: "GridViewMedium" } },
    { key: "divider_1", text: "-", itemType: DropdownMenuItemType.Divider },
    { key: "allDocuments", text: "All Documents" },
];

let allTemplates: any[] = [];

/**
 * State for the myTemplatesTabTab React component
 */
export interface IMyTemplatesTabState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the myTemplatesTabTab React component
 */
export interface IMyTemplatesTabProps extends ITeamsBaseComponentProps {

}

/**
 * Implementation of the My Templates content page
 */
export class MyTemplatesTab extends TeamsBaseComponent<any, any> {

    constructor(props: any) {
        super(props, {});
        initializeIcons();
        this.state = {
            templates: [],
            context: null,
            loading: false,
            selectedTemplates: [],
            HideAccessDeniedDialogBox: true,
            view: "Tiles",
            filterMenu: "*",
            searchString: "",

            templateSelectorSection: true,      // search and select template
            // nextSection: false,
            previewSection: false,
            createSection: false,
            selectSection: false,

            showPreviousButton: false,
            showPreviewButton: true,
            showNextButton: true,
            showCreateButton: false,
            // showSelectButton: false,
            isPinnedLocation: false,
            enableCreateButton: false,
            CurrentPosition: 0,
            saveTemplateDetails: {
                DestFolderRelUrl: "",
                DestSite: "",
                IsPinnedLocation: false,
                LocationAlreadyPinned: false,
                SiteUrl: "",
                InvalidLocation: ""
            },
            loggedInUserEmail: "",
            tenantFullName: "",
            BreadCrumb: [{
                Location: "",
                LocationDisplay: "Home",
                Order: 0,
                Contents: [
                    {

                    } as IFileFolder
                ]
            }] as IBreadCrumb[],
            // Selection Location State
            PinnedLocations: [],
            Teams: [],
            OneDrive: [],
            OverlayLoaderText: "We are setting up your templates, please wait.....",
        };
    }

    public componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));
        this.setState({
            fontSize: this.pageFontSize()
        });

        if (this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);

        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
    }

    public componentDidMount() {
        if (this.inTeams()) {
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId,
                    context,
                    loggedInUserEmail: context.loginHint
                }, () => {
                    this.loadInitData();
                });

                pnpSetup({
                    spfxContext: context
                });
                graph.setup({
                    spfxContext: context
                });

            });
        }
    }

    public fetchLocationDetails = () => {
        let apiService: IMyTemplateTabAPI = new MyTemplateTabAPIImpl();
        let payload: ILocationPayload = {
            SPOUrl: "https://".concat(this.state.context.teamSiteDomain),
            tenant: this.state.context.tid,
            FolderPath: "",
            AccountName: this.state.loggedInUserEmail,
            TenantName: this.state.tenantFullName,
        };

        apiService.GetLocationDetails(payload).then(response => {

            let finalBreadCrumb: IBreadCrumb[] = util.loadBreadCrumb(payload.FolderPath, this.state.BreadCrumb, response.d, 0, "");
            let templates = util.populateTemplates(this.state.context, finalBreadCrumb[0].Contents);
            this.setState({
                templates,
                CurrentPosition: 0,
                loading: false,
                BreadCrumb: finalBreadCrumb
            });

        }).catch((error) => {
            console.error("Error fetchLocationDetails", error);
            this.setState({
                loading: false
            });
        });
    }

    public onFolderClick = (folderBreadCrumb: IFolderBreadCrumb): void => {

        this.setState({
            loading: true,
            OverlayLoaderText: "Loading ...."
        });
        let indexOfSelectedFolderInBreadCrumb: number = this.state.BreadCrumb
            .map(data => data.Location).indexOf(folderBreadCrumb.Location);         // check if already exists

        if (indexOfSelectedFolderInBreadCrumb > -1) {

            let tempBreadCrumb: IBreadCrumb[] = [...this.state.BreadCrumb];

            let selectedBreadCrumb: IBreadCrumb = this.state.BreadCrumb[indexOfSelectedFolderInBreadCrumb];
            tempBreadCrumb = util.showHideBreadCrumb(this.state.BreadCrumb, selectedBreadCrumb.Location, selectedBreadCrumb.Order);
            this.setState({
                templates: util.populateTemplates(this.state.context, selectedBreadCrumb.Contents),
                CurrentPosition: selectedBreadCrumb.Order,
                BreadCrumb: tempBreadCrumb,
                loading: false
            });
        }
        else {

            let apiService: IMyTemplateTabAPI = new MyTemplateTabAPIImpl();
            let payload: ILocationPayload = {
                SPOUrl: "https://".concat(this.state.context.teamSiteDomain),
                tenant: this.state.context.tid,
                FolderPath: folderBreadCrumb.Location,
                AccountName: this.state.loggedInUserEmail,
                TenantName: this.state.tenantFullName,
            };
            apiService.GetLocationDetails(payload).then(response => {

                if (response.folderAccess && response.folderAccess === "Access Denined") {
                    let tempBreadCrumb: IBreadCrumb[] = [...this.state.BreadCrumb];
                    let UniqueId: string = this.state.templates.filter(data => data.Name === folderBreadCrumb.Display)[0].UniqueId;
                    tempBreadCrumb = util.updateBreadCrumbForAccessDenied(tempBreadCrumb, UniqueId);

                    let contents: IFileFolder[] = [];

                    // tslint:disable-next-line: prefer-for-of
                    for (let index = 0; index < tempBreadCrumb.length; index++) {
                        if (tempBreadCrumb[index].Contents.filter((folder: IFileFolder) => folder.UniqueId === UniqueId).length > 0) {
                            contents = tempBreadCrumb[index].Contents;
                            break;
                        }
                    }

                    this.setState({
                        templates: util.populateTemplates(this.state.context, contents),
                        loading: false,
                        BreadCrumb: tempBreadCrumb,
                        HideAccessDeniedDialogBox: false
                    }, () => {
                        setTimeout(() => {
                            this.setState({
                                HideAccessDeniedDialogBox: true
                            });
                        }, 2000);
                    }
                    );
                }
                else {
                    let finalBreadCrumb: IBreadCrumb[] = util.loadBreadCrumb(folderBreadCrumb.Location, this.state.BreadCrumb, response.d, this.state.CurrentPosition + 1, folderBreadCrumb.Display);
                    let Contents: IFileFolder[] = finalBreadCrumb.filter((data: IBreadCrumb) => data.Order === this.state.CurrentPosition + 1 && data.InCurrentView === true)[0].Contents;
                    this.setState({
                        templates: util.populateTemplates(this.state.context, Contents),
                        CurrentPosition: this.state.CurrentPosition + 1,
                        loading: false,
                        BreadCrumb: finalBreadCrumb
                    });

                }
            }).catch((error) => {
                console.error("Error fetchLocationDetails", error);
                this.setState({
                    loading: false
                });
            });
        }

    }

    public loadInitData = () => {

        this.setState({
            loading: true
        });
        // let payload = {
        //     tenant: this.state.context.tid,
        //     SPOUrl: `https://${this.state.context.teamSiteDomain}`
        // };

        let teantNamePayload = {
            tenant: this.state.context.tid,
            SPOUrl: "https://graph.microsoft.com"
        };

        fetch(GET_TENANT_FULLNAME, {
            method: 'post',
            body: JSON.stringify(teantNamePayload)
        }).then((response) => {
            return response.json();
        }).then((data) => {
            let tenantFullQuaifiedName = "";
            data.value[0].verifiedDomains.map(domain => {
                if (domain.isInitial) {
                    tenantFullQuaifiedName = domain.name;
                    tenantFullQuaifiedName = tenantFullQuaifiedName.substr(0, tenantFullQuaifiedName.indexOf(".onmicrosoft.com"));
                }
            });
            this.setState({
                tenantFullName: tenantFullQuaifiedName,
                // loading: false
            }, () => this.fetchLocationDetails());
        }).catch(error => {
            console.log("Error while fetching templates - ", error);
            this.setState({
                loading: false
            });
        });
    }

    // CheckBox Operation
    public onCheckUncheckTemplate = (template: any, checked: boolean): void => {

        let currentTemplates = [...this.state.templates];

        const currentState = this.state.active;

        let tempSelectedTemplates: any[] = [...this.state.selectedTemplates];

        currentTemplates.map(eachTemplate => {

            if (eachTemplate.UniqueId === template.UniqueId) {
                let indexOfDeselctedItem = tempSelectedTemplates.map(data => data.UniqueId).indexOf(eachTemplate.UniqueId);
                if (indexOfDeselctedItem === -1) {   // selected and unique id is equal
                    eachTemplate.isSelected = true;
                    tempSelectedTemplates.push(eachTemplate);
                }
                else {
                    eachTemplate.isSelected = false;
                    tempSelectedTemplates.splice(indexOfDeselctedItem, 1);
                }

            }
        });

        this.setState({
            active: !currentState,
            templates: currentTemplates,
            selectedTemplates: tempSelectedTemplates
        });
    }

    public onErrorImageLoad = (template) => {
        let currentTemplates = [...this.state.templates];
        let indexOfSelectedTemplate = currentTemplates.map(data => data.UniqueId).indexOf(template.UniqueId);
        let selectedTemplate = currentTemplates[indexOfSelectedTemplate];
        selectedTemplate.imgUrl = selectedTemplate.altImgUrl;

        currentTemplates[indexOfSelectedTemplate] = selectedTemplate;

        this.setState({
            templates: currentTemplates.sort(util.compare)
        });
    }

    public getSelectSectionDetails = (obj: any) => {

        let saveTemplateDetails = { ...this.state.saveTemplateDetails };

        // tslint:disable-next-line: forin
        for (let key in obj) {
            saveTemplateDetails[key] = obj[key];
        }

        // saveTemplateDetails = obj;

        this.setState({
            saveTemplateDetails: saveTemplateDetails,
            enableCreateButton:
                (obj.DestFolderRelUrl && obj.SiteUrl && obj.DestSite)
                || (obj.IsPinnedLocation && this.state.enableCreateButton)
                || (obj.LocationAlreadyPinned)
                || (obj.InvalidLocation)
        });
    }

    public getFilteredTemplates = (searchString: any) => {
        this.setState({
            searchString
        });
        let currentTemplates: any[] = [];
        allTemplates.map(template => {
            let eachTemplate = template;
            let filterMenu = this.state.filterMenu;
            if (searchString && template.Name.toUpperCase().indexOf(searchString.toUpperCase()) > -1) {
                if (filterMenu !== "*" && template.Name.indexOf(filterMenu) === -1) {
                    eachTemplate.toShow = false;
                }
                else {
                    eachTemplate.toShow = true;
                }
                eachTemplate.isSelected = false;

                if (template.File.ServerRelativeUrl) {
                    eachTemplate.imgUrl =
                        TEMPLATE_IMAGE_PATH + `&tenant=${this.state.context.tid}` +
                        `&SPOUrl=` + `https://${this.state.context.teamSiteDomain}` +
                        `&ImgPath=https://${this.state.context.teamSiteDomain}${template.File.ServerRelativeUrl}`;

                }
                if (template.Name.indexOf(".xls") > -1) {
                    eachTemplate.altImgUrl = DEFAULT_XLSX_URL;
                }
                else if (template.Name.indexOf(".ppt") > -1) {
                    eachTemplate.altImgUrl = DEFAULT_PPTX_URL;

                }
                else if (template.Name.indexOf(".doc") > -1) {
                    eachTemplate.altImgUrl = DEFAULT_DOCX_URL;

                }

                currentTemplates.push(eachTemplate);
            }
            else if (!searchString) {
                if (filterMenu !== "*" && template.Name.indexOf(filterMenu) === -1) {
                    eachTemplate.toShow = false;
                }
                else {
                    eachTemplate.toShow = true;
                }
                currentTemplates.push(eachTemplate);
            }
            else {
                eachTemplate.toShow = false;
                currentTemplates.push(eachTemplate);

            }

        });


        this.setState({
            templates: currentTemplates.sort(util.compare),
        });

    }

    public onClickPreviousIcon = (): void => {
        let currentTemplates = [...this.state.templates];
        let saveTemplateDetails = { ...this.state.saveTemplateDetails };
        saveTemplateDetails = {
            DestFolderRelUrl: "",
            DestSite: "",
            IsPinnedLocation: true
        };
        let tempSelectedTemplates: any[] = [];

        currentTemplates.map(eachTemplate => {
            if (eachTemplate.isSelected) {
                eachTemplate.isSelected = false;
            }
        });

        this.setState({
            templates: currentTemplates.sort(util.compare),
            selectedTemplates: tempSelectedTemplates,
            saveTemplateDetails,
            enableCreateButton: false
        });


        this.setState({
            previewSection: false,
            nextSection: false,
            selectSection: false,
            createSection: false,
            templateSelectorSection: true,
            showPreviousButton: false,
            showPreviewButton: true,
            showNextButton: true,
            showSelectButton: false,
            showCreateButton: false
        });
    }

    // Click on Close Pop Up - Amartya
    public onClosePreviewPopup = (isNext: boolean): void => {
        this.setState({
            previewSection: false,
            selectSection: isNext,
            createSection: false,
            templateSelectorSection: !isNext,
            showPreviousButton: isNext,
            showPreviewButton: !isNext,
            showNextButton: !isNext,
            showCreateButton: isNext,
            enableCreateButton: false
        });
    }

    public selectTemplateFromPopup = (template: any): void => {

        let currentTemplates = [...this.state.templates];
        let indexOfSelectedTemplate = currentTemplates.map(data => data.UniqueId).indexOf(template.UniqueId);
        let selectedTemplate = currentTemplates[indexOfSelectedTemplate];
        selectedTemplate.isSelected = true;

        currentTemplates[indexOfSelectedTemplate] = selectedTemplate;
        let tempSelectedTemplates: any[] = [];

        currentTemplates.map(eachTemplate => {
            if (eachTemplate.UniqueId === selectedTemplate.UniqueId) {
                tempSelectedTemplates.push(eachTemplate);
            }
            else {
                eachTemplate.isSelected = false;

            }
        });

        this.setState({
            templates: currentTemplates.sort(util.compare),
            selectedTemplates: tempSelectedTemplates
        });
    }

    public setSelectionLocationState = (pinnedLocations: any, teamLocations: any, driveLocations: any) => {

        this.setState({
            PinnedLocations: pinnedLocations,
            Teams: teamLocations,
            OneDrive: driveLocations,
        });
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };

        let { BreadCrumb, CurrentPosition, templates, searchString, filterMenu, selectedTemplates, HideAccessDeniedDialogBox } = this.state;

        let FolderBreadCrumbs: IFolderBreadCrumb[] = [];
        let FileTemplates: any[] = [];
        let FolderTemplates: any[] = [];

        BreadCrumb.filter((data: IBreadCrumb) => data.Order <= CurrentPosition && data.InCurrentView)
            .map((data: IBreadCrumb) => {
                FolderBreadCrumbs.push({
                    Display: data.LocationDisplay,
                    Location: data.Location
                });
            });

        if (filterMenu !== "folder") {

            templates.filter((data: any) =>
                filterMenu === "*" ?
                    data.type === "SP.File" &&
                    data.Name.toUpperCase().indexOf(searchString.toUpperCase()) > -1
                    :
                    data.type === "SP.File" &&
                    data.Name.toUpperCase().indexOf(searchString.toUpperCase()) > -1 &&
                    data.Name.toUpperCase().indexOf(filterMenu.toUpperCase()) > -1)
                .map((data: any) => {
                    let currentTemplate = { ...data };
                    currentTemplate.type = data.type;
                    currentTemplate.Name = data.Name;
                    currentTemplate.ServerRelativeUrl = data.ServerRelativeUrl;
                    currentTemplate.Title = data.Title;
                    currentTemplate.UniqueId = data.UniqueId;
                    currentTemplate.toShow = true;
                    if (selectedTemplates.filter(d => d.UniqueId === data.UniqueId).length > 0) {
                        currentTemplate.isSelected = true;
                    }
                    else {
                        currentTemplate.isSelected = currentTemplate.isSelected;

                    }
                    currentTemplate.ParentLocation = data.ParentLocation;
                    FileTemplates.push(currentTemplate);
                });
        }
        if (filterMenu === "*" || filterMenu === "folder") {
            templates.filter((data: any) =>
                data.type === "SP.Folder" &&
                data.Name.toUpperCase().indexOf(searchString.toUpperCase()) > -1)
                .map((data: any) => {
                    let currentTemplate = { ...data };
                    currentTemplate.Display = data.Name;
                    currentTemplate.toShow = true;
                    currentTemplate.Location = data.ParentLocation ? data.ParentLocation.concat(`/${data.Name}`) : data.ParentLocation.concat(`${data.Name}`);
                    FolderTemplates.push(currentTemplate);
                });
        }

        const labelId = getId('dialogLabel');
        // tslint:disable-next-line: variable-name
        const subTextId = getId('subTextLabel');

        return (
            <TeamsThemeContext.Provider value={context} >
                <Dialog
                    hidden={HideAccessDeniedDialogBox}
                    onDismiss={() => this.setState({ HideAccessDeniedDialogBox: true })}
                    dialogContentProps={{
                        type: DialogType.normal,
                        title: 'Access Denied!!',
                        closeButtonAriaLabel: 'Close',
                        subText: 'Sorry!! Yot do not have access to view this folder, Please contact System Administrator'
                    }}
                    modalProps={{
                        titleAriaId: labelId,
                        subtitleAriaId: subTextId,
                        isBlocking: false,
                        styles: { main: { maxWidth: 450 } },
                    }}
                />
                <Surface className="tabContainer">
                    <Panel>
                        <PanelHeader>
                        </PanelHeader>
                        <PanelBody className="container">
                            <LoadingOverlay
                                active={this.state.loading}
                                spinner={<FadeLoader />}
                                text={this.state.OverlayLoaderText}
                            >
                                {/* <div className="ms-Grid">
                                    <div className="ms-Grid-row">
                                        <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12 site-image-div">
                                            <img className="site-image"
                                                src={DEFAULT_SITE_URL} alt="" />
                                        </div>
                                    </div>
                                </div> */}
                                {this.state.templateSelectorSection &&
                                    <div>

                                        <div className="ms-Grid">
                                            <div className="ms-Grid-row">

                                                <div className="ms-Grid-col ms-sm12 ms-md5 ms-lg5">
                                                    <SearchBox
                                                        value={searchString}
                                                        placeholder="Search Templates"
                                                        onChange={(_, newValue) =>
                                                            this.setState({
                                                                searchString: newValue
                                                            })
                                                        }
                                                    />
                                                </div>
                                                <div className="ms-Grid-col ms-sm12 ms-md5 ms-lg5">
                                                    <Dropdown
                                                        placeholder="Select an option"
                                                        ariaLabel="Custom dropdown example"
                                                        onRenderPlaceholder={this.onRenderPlaceholder}
                                                        onRenderOption={this.onRenderOption}
                                                        onRenderCaretDown={this.onRenderCaretDown}
                                                        selectedKey={this.state.view}
                                                        onRenderTitle={this.onRenderTitle}
                                                        onChanged={(item) => {
                                                            if (item.key.toString() === "allDocuments") {
                                                                let currentView = this.state.view;
                                                                // allTemplates = [];
                                                                this.setState({ view: currentView });
                                                                // this.loadInitData();
                                                            }
                                                            else {
                                                                this.setState({ view: item.key.toString() });
                                                            }
                                                        }}
                                                        options={options}
                                                    />
                                                </div>
                                                <div className="ms-Grid-col ms-sm12 ms-md2 ms-lg2">
                                                    <DefaultButton
                                                        text={this.state.filterMenu === "*" ? "ALL" : this.state.filterMenu}
                                                        iconProps={filterIcon}
                                                        menuProps={{
                                                            shouldFocusOnMount: true,
                                                            items: [
                                                                {
                                                                    key: 'noFilter',
                                                                    text: 'Clear Filter',
                                                                    onClick: () => {
                                                                        this.setState({
                                                                            filterMenu: "*"
                                                                        });
                                                                    }

                                                                },
                                                                {
                                                                    key: 'folder',
                                                                    text: 'folder',
                                                                    onClick: () => {
                                                                        this.setState({
                                                                            filterMenu: "folder"
                                                                        });
                                                                    }

                                                                },
                                                                {
                                                                    key: 'docx',
                                                                    text: 'docx',
                                                                    onClick: () => {
                                                                        this.setState({
                                                                            filterMenu: "docx"
                                                                        });
                                                                    }

                                                                },
                                                                {
                                                                    key: 'pptx',
                                                                    text: 'pptx',
                                                                    onClick: () => {
                                                                        this.setState({
                                                                            filterMenu: "pptx"
                                                                        });
                                                                    }
                                                                },
                                                                {
                                                                    key: 'xlsx',
                                                                    text: 'xlsx',
                                                                    onClick: () => {
                                                                        this.setState({
                                                                            filterMenu: "xlsx"
                                                                        });
                                                                    }
                                                                },
                                                                {
                                                                    key: 'dotx',
                                                                    text: 'dotx',
                                                                    onClick: () => {
                                                                        this.setState({
                                                                            filterMenu: "dotx"
                                                                        });
                                                                    }
                                                                },
                                                                {
                                                                    key: 'potx',
                                                                    text: 'potx',
                                                                    onClick: () => {
                                                                        this.setState({
                                                                            filterMenu: "potx"
                                                                        });
                                                                    }
                                                                },
                                                                {
                                                                    key: 'xltx',
                                                                    text: 'xltx',
                                                                    onClick: () => {
                                                                        this.setState({
                                                                            filterMenu: "xltx"
                                                                        });
                                                                    }
                                                                }
                                                            ]
                                                        }}
                                                    />
                                                </div>
                                            </div>
                                        </div>
                                        <ul className="breadcrumb">
                                            {FolderBreadCrumbs.reverse().map((data: IFolderBreadCrumb) => {
                                                return (
                                                    <li onClick={() => this.onFolderClick(data)}>
                                                        <span>
                                                            {data.Display} <Icon className="ChevronRightSmall" />
                                                        </span>
                                                    </li>
                                                );
                                            })}
                                        </ul>
                                        <div className="ms-Grid scroller">
                                            {this.state.view === "Tiles" &&
                                                <React.Fragment>
                                                    <div className="ms-flex pdTop5">
                                                        {FolderTemplates.
                                                            map((template) => {
                                                                return (
                                                                    <div className={["ms-flex-item", "tilesHover animated", template.toShow ? "zoomIn" : "zoomOut"].join(" ").trim()}>

                                                                        <div className="imageContainer" onClick={() => {
                                                                            let folder: IFolderBreadCrumb = {
                                                                                Display: template.Display,
                                                                                Location: template.Location
                                                                            };
                                                                            this.onFolderClick(folder);
                                                                        }
                                                                        }>
                                                                            <img src={template.imgUrl} alt="" className="doc-image" />
                                                                        </div>
                                                                        <div title={template.Name} className="image-name-container">
                                                                            {template.Name}
                                                                        </div>

                                                                    </div>
                                                                );
                                                            })
                                                        }
                                                    </div>
                                                    <hr />
                                                    <div className="ms-flex pdTop5">
                                                        {FileTemplates.
                                                            map((template) => {
                                                                return (

                                                                    <div className={["ms-flex-item", "tilesHover animated", template.toShow ? "zoomIn" : "zoomOut"].join(" ").trim()}>

                                                                        <div className={(template.isSelected) ? "imageContainer isSelected" : "imageContainer "} onClick={() => this.onCheckUncheckTemplate(template, !template.isSelected)}>
                                                                            <img src={template.imgUrl} onError={() => this.onErrorImageLoad(template)} alt="" className="doc-image" />
                                                                        </div>
                                                                        <div title={template.Name} className="image-name-container">
                                                                            {template.Name}
                                                                        </div>
                                                                        <div className="template-checkbox">
                                                                            <Checkbox
                                                                                checked={template.isSelected}
                                                                                onChecked={() => this.onCheckUncheckTemplate(template, !template.isSelected)}
                                                                            />
                                                                        </div>
                                                                    </div>

                                                                );
                                                            })
                                                        }
                                                    </div>
                                                </React.Fragment>
                                            }
                                            {this.state.view === "List" &&
                                                <React.Fragment>
                                                    {FolderTemplates.
                                                        map((template) => {
                                                            return (

                                                                <div
                                                                    onClick={() => {
                                                                        let folder: IFolderBreadCrumb = {
                                                                            Display: template.Display,
                                                                            Location: template.Location
                                                                        };
                                                                        this.onFolderClick(folder);
                                                                    }}
                                                                    className={["ms-Grid-col ms-sm-12 ms-md12 ms-lg12 list-view animated folder", template.toShow ? "slideInUp" : ""].join(" ").trim()}>

                                                                    <img src={template.altImgUrl} alt="" className="list-image"


                                                                    />

                                                                    <div title={template.Name} className="list-item-name">
                                                                        {template.Name}
                                                                    </div>
                                                                </div>

                                                            );
                                                        })
                                                    }
                                                    <hr />
                                                    {FileTemplates.
                                                        map((template) => {
                                                            return (

                                                                <div className={["ms-Grid-col ms-sm-12 ms-md12 ms-lg12 list-view animated", template.toShow ? "slideInUp" : ""].join(" ").trim()}>
                                                                    <Checkbox style={{ float: "left" }}
                                                                        checked={template.isSelected}
                                                                        onChecked={() => this.onCheckUncheckTemplate(template, !template.isSelected)}
                                                                    />
                                                                    <img src={template.altImgUrl} alt="" className="list-image" />

                                                                    <div title={template.Name} className="list-item-name">
                                                                        {template.Name}
                                                                    </div>
                                                                </div>

                                                            );
                                                        })
                                                    }
                                                    <hr />
                                                </React.Fragment>
                                            }
                                        </div>

                                    </div>
                                }
                                {/* {this.state.previewSection &&
                                    <PreviewSection
                                        selectedTemplate={this.state.selectedTemplates[0]}
                                    />
                                } */}
                                {this.state.previewSection &&
                                    <Modal
                                        isOpen={this.state.previewSection}
                                        onDismiss={() => this.onClosePreviewPopup(false)}
                                        isBlocking={false}
                                        containerClassName={contentStyles.container}>
                                        <div className={contentStyles.header}>
                                            <IconButton
                                                styles={iconButtonStyles}
                                                iconProps={{ iconName: 'Cancel' }}
                                                ariaLabel="Close popup modal"
                                                onClick={() => this.onClosePreviewPopup(false)}
                                            />
                                        </div>
                                        <div className={contentStyles.body}>
                                            <PreviewSectionPopup
                                                onClosePreviewPopup={this.onClosePreviewPopup}
                                                selectTemplateFromPopup={this.selectTemplateFromPopup}
                                                templates={this.state.templates.filter(data => data.Name.match(/\.[a-zA-Z]{3,}$/))}
                                                selectedTemplate={this.state.selectedTemplates[0]}
                                            />
                                        </div>
                                    </Modal>

                                }
                                {this.state.nextSection &&
                                    <NextSection
                                        defaultSiteCollection="DocsNode Demo"
                                    />
                                }
                                {this.state.selectSection &&
                                    <SelectSection
                                        getSelectSectionDetails={this.getSelectSectionDetails}
                                        setSelectionLocationState={this.setSelectionLocationState}
                                        PinnedLocations={this.state.PinnedLocations}
                                        Teams={this.state.Teams}
                                        OneDrive={this.state.OneDrive}
                                    />
                                }
                                {this.state.createSection &&
                                    <CreateSection
                                        context={this.state.context}
                                        selectedTemplates={this.state.selectedTemplates}
                                        saveTemplateDetails={this.state.saveTemplateDetails}
                                        isPinnedLocation={this.state.isPinnedLocation}
                                        loggedInUserEmail={this.state.loggedInUserEmail}
                                        onClickPreviousIcon={this.onClickPreviousIcon}
                                        tenantFullName={this.state.tenantFullName}
                                    />
                                }
                                {/* Button Section */}
                                <div className="ms-Grid btn-padding">
                                    <div className="ms-Grid-row">
                                        {this.state.showPreviewButton &&

                                            <div className="ms-Grid-col ms-sm-12 ms-md6 ms-lg2">
                                                <button
                                                    className="normal-button"
                                                    disabled={this.state.selectedTemplates.length !== 1}
                                                    onClick={() => this.setState({
                                                        previewSection: true,
                                                        templateSelectorSection: true,
                                                        // nextSection: false,
                                                        selectSection: false,
                                                        createSection: false,
                                                        showPreviewButton: false,
                                                        showPreviousButton: false,
                                                        showNextButton: true,
                                                        showSelectButton: false,
                                                        showCreateButton: false
                                                    })}>
                                                    <i className="fa fa-eye" aria-hidden="true"></i> Preview
                                                </button>
                                            </div>
                                        }
                                        {this.state.showPreviousButton &&
                                            <div className="ms-Grid-col ms-sm-12 ms-md6 ms-lg2">
                                                <button
                                                    className="normal-button"
                                                    onClick={() => this.setState({
                                                        previewSection: false,
                                                        selectSection: this.state.createSection,
                                                        createSection: false,
                                                        templateSelectorSection: !this.state.createSection,
                                                        showPreviousButton: this.state.createSection,
                                                        showPreviewButton: !this.state.createSection,
                                                        showNextButton: !this.state.createSection,
                                                        showCreateButton: this.state.createSection,
                                                        enableCreateButton: false
                                                    })
                                                    }>
                                                    <i className="fa fa-chevron-left" aria-hidden="true"></i> Previous
                                                </button>
                                            </div>
                                        }
                                        {this.state.showNextButton &&
                                            <div className="ms-Grid-col ms-sm-12 ms-md6 ms-lg2">
                                                <button
                                                    className="normal-button"
                                                    disabled={this.state.selectedTemplates.length === 0}
                                                    onClick={() => this.setState({
                                                        templateSelectorSection: false,
                                                        selectSection: true,
                                                        createSection: false,
                                                        previewSection: false,
                                                        showPreviousButton: true,
                                                        showPreviewButton: false,
                                                        showNextButton: false,
                                                        showSelectButton: false,
                                                        showCreateButton: true,
                                                    })}>
                                                    Next <i className="fa fa-chevron-right" aria-hidden="true"></i>
                                                </button>
                                            </div>
                                        }
                                        {this.state.showCreateButton &&
                                            <div className="ms-Grid-col ms-sm-12 ms-md6 ms-lg2">
                                                <button
                                                    className="normal-button"
                                                    disabled={this.state.selectedTemplates.length === 0 || !this.state.enableCreateButton}
                                                    onClick={() => this.setState({
                                                        templateSelectorSection: false,
                                                        selectSection: false,
                                                        previewSection: false,
                                                        createSection: true,
                                                        showPreviousButton: true,
                                                        showPreviewButton: false,
                                                        showNextButton: false,
                                                        showCreateButton: false
                                                    })}>
                                                    Create <i className="fa fa-plus-circle" aria-hidden="true"></i>
                                                </button>
                                            </div>
                                        }
                                    </div>
                                </div>

                            </LoadingOverlay>
                        </PanelBody>
                        <PanelFooter>

                        </PanelFooter>
                    </Panel>
                </Surface>
            </TeamsThemeContext.Provider >
        );
    }

    private onRenderTitle = (allOptions: IDropdownOption[]): JSX.Element => {
        const option = allOptions[0];

        return (
            <div>
                {option.data && option.data.icon && (
                    <Icon style={{ marginRight: '8px' }} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
                )}
                <span>{option.text}</span>
            </div>
        );
    }

    private onRenderOption = (option: IDropdownOption): JSX.Element => {

        let optionView: string = "";
        switch (option.key) {
            case "List":
            case "Tiles":
                optionView = option.key;
                break;
            default:
                optionView = option.text;
                break;
        }

        return (
            <div>
                {option.data && option.data.icon && (
                    <Icon style={{ marginRight: "8px" }} iconName={option.data.icon} aria-hidden="true" title={option.data.icon} />
                )}
                <span>{optionView}</span>
            </div>
        );
    }



    private onRenderPlaceholder = (props: IDropdownProps): JSX.Element => {
        return (
            <div className="dropdownExample-placeholder">
                <Icon style={{ marginRight: "8px" }} iconName={"MessageFill"} aria-hidden="true" />
                <span>{props.placeholder}</span>
            </div>
        );
    }

    private onRenderCaretDown = (props: IDropdownProps): JSX.Element => {
        return <Icon iconName="ChevronDown" />;
    }

    private changeView = (value: any): void => {
        // console.log(value);
    }


}
