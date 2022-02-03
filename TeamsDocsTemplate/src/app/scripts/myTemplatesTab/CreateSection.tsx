import * as React from "react";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import {
    DefaultButton, IContextualMenuProps, getTheme,
    mergeStyleSets,
    FontWeights,
    ContextualMenu,
    Toggle,
    Modal,
    IDragOptions,
    IconButton,
    IIconProps,
} from 'office-ui-fabric-react';
import { initializeIcons } from '@uifabric/icons';
import LoadingOverlay from 'react-loading-overlay';
import FadeLoader from 'react-spinners/FadeLoader';
import { CREATE_TEMPLATE_URL, CHANGE_CREATED_BY_URL, PIN_LOCATION_URL, SAVE_TEMPLATE_IN_ONE_DRIVE, GET_TEMPLATE_FROM_ONE_DRIVE } from './../utility/Constants';
import { SelectMyLibraryApi } from "../services";
import { IPayloadSaveOneDrive } from "../interfaces/IPayload";



export class CreateSection extends React.Component<any, any> {

    constructor(props: any) {
        super(props, {});
        initializeIcons();
        this.state = {
            loading: false,
            selectedTemplates: [],
            saveTemplateDetails: {
                DestFolderRelUrl: "",
                DestSite: "",
                IsPinnedLocation: false,
                LocationAlreadyPinned: false,
                SiteUrl: "https://docsnode.sharepoint.com/",
                InvalidLocation: ""
            },
            toBeSavedlocation: "documentLib",
            counter: 0,
            showPreviousIcon: false,
            inputDocument: { name: "", isvalid: false, errorMessage: "" },
            savedDocuments: [],
            documentLocation: "#",
            tenantFullName: "",
            teamListName: "",
            documentModalOpen: false,
            modalDocumentLink: "",
        };
    }

    public componentWillMount() {
        this.setState({
            selectedTemplates: this.props.selectedTemplates,
            saveTemplateDetails: this.props.saveTemplateDetails,
            loggedInUserEmail: this.props.loggedInUserEmail,
            tenantFullName: this.props.tenantFullName,
            savedDocuments: []
        });
    }

    public componentDidMount() {

        const eachTeamLibPayload = {
            tenant: this.props.context.tid,
            SPOUrl: `https://${this.props.context.teamSiteDomain}`,
            Team: this.state.saveTemplateDetails.DestSite
        };
        new SelectMyLibraryApi()
            .getTeamLibraryName(eachTeamLibPayload)
            .then((libArrResp: any[]) => {

                let teamListname = "";

                if (libArrResp && libArrResp.length > 0) {
                    teamListname = libArrResp[0].Title;
                }

                this.setState({
                    teamListName: teamListname
                });

            }).catch(error => {
                console.error("Error while getTeamsWebUrl ", error);
            });

    }

    public componentWillReceiveProps(nextProps: any) {
        this.setState({
            selectedTemplates: nextProps.selectedTemplates,
            saveTemplateDetails: nextProps.saveTemplateDetails,
            loggedInUserEmail: nextProps.loggedInUserEmail,
            tenantFullName: nextProps.tenantFullName,
        });
    }

    public saveTemplateInTeams = (currentTemplate: any): void => {

        this.setState({
            loading: true,
        });
        let documentExtention = currentTemplate.Name.substr(currentTemplate.Name.indexOf("."));
        let inputDocument = { ...this.state.inputDocument };
        let savedDocumentName = inputDocument.name + documentExtention;

        let payloadCreateFile = {
            // tenant: "docsnode.com",
            tenant: this.props.context.tid,
            SPOUrl: `https://${this.props.context.teamSiteDomain}`,
            sourceFileName: currentTemplate.ParentLocation + "/" + currentTemplate.Name,
            DestFolderRelUrl: this.state.saveTemplateDetails.DestFolderRelUrl,
            DestSite: this.state.saveTemplateDetails.DestSite,
            FileName: inputDocument.name
        };

        // console.log(payloadCreateFile);

        fetch(CREATE_TEMPLATE_URL, {
            method: 'post',
            body: JSON.stringify(payloadCreateFile)
        }).then((response: any) => {
            return response.json();
        }).then((data) => {
            // console.log(data.d);
            if (data.error && data.error.code === "-2130575257, Microsoft.SharePoint.SPException") {
                inputDocument.isValid = false;
                inputDocument.errorMessage = 'Duplicate file name, file already exists';
                this.setState({
                    inputDocument,
                    loading: false
                });
            }
            else {
                let tempCounter = this.state.counter + 1;

                inputDocument.isValid = false;
                inputDocument.name = "";
                inputDocument.errorMessage = "";

                let savedDocuments = [...this.state.savedDocuments];

                savedDocuments.push({
                    name: savedDocumentName,
                    url: data.d.ServerRedirectedEmbedUri.replace("&action=interactivepreview", ""),
                    clientAppLink: this.GetClientAppLinkforTeams(this.state.saveTemplateDetails.DestSite, this.state.saveTemplateDetails.DestFolderRelUrl, savedDocumentName)
                });
                let payloadChangeCreatedBy = {
                    tenantName: this.state.tenantFullName,
                    siteUrl: this.state.saveTemplateDetails.DestSite.toString(),
                    listName: this.state.teamListName,
                    // listName: "Documents",
                    itemID: data.d.Id,
                    emailid: this.state.loggedInUserEmail.toString()
                };

                // console.log(payloadChangeCreatedBy);
                fetch(CHANGE_CREATED_BY_URL, {
                    headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'
                    },
                    method: 'post',
                    body: JSON.stringify(payloadChangeCreatedBy)

                }).then((responseChange: any) => {
                    // console.log(responseChange);
                    return responseChange.text();
                }).then((dataChangeCreatedBy) => {
                    // console.log(dataChangeCreatedBy);
                    if (this.state.saveTemplateDetails.IsPinnedLocation && !this.state.saveTemplateDetails.LocationAlreadyPinned) {
                        let payloadPinLocation = {
                            tenant: this.props.context.tid,
                            SPOUrl: `https://${this.props.context.teamSiteDomain}`,
                            PinnedType: "Teams",
                            DocumentLibrary: this.state.saveTemplateDetails.DestFolderRelUrl,
                            DocumentLibraryUrl: this.state.saveTemplateDetails.DestSite,
                            SiteUrl: this.state.saveTemplateDetails.SiteUrl

                        };
                        fetch(PIN_LOCATION_URL, {
                            headers: {
                                'Accept': 'application/json',
                                'Content-Type': 'application/json'
                            },
                            method: 'post',
                            body: JSON.stringify(payloadPinLocation)

                        }).then((responsePinLocation: any) => {
                            // console.log(responsePinLocation);
                            return responsePinLocation.json();
                        }).then((dataPinLocation) => {
                            let payloadChangeCreatedByPinLoc = {
                                tenantName: this.state.tenantFullName,
                                siteUrl: `https://${this.props.context.teamSiteDomain}`,
                                listName: "DocsNodePinnedLocations",
                                itemID: dataPinLocation.d.Id,
                                emailid: this.state.loggedInUserEmail.toString()
                            };
                            fetch(CHANGE_CREATED_BY_URL, {
                                headers: {
                                    'Accept': 'application/json',
                                    'Content-Type': 'application/json'
                                },
                                method: 'post',
                                body: JSON.stringify(payloadChangeCreatedByPinLoc)

                            }).then((responseChangePinLoc: any) => {
                                // console.log(responseChangePinLoc);
                                return responseChangePinLoc.text();
                            }).then((dataChangeCreatedByPinLoc) => {

                                let saveTemplateDetails = { ...this.state.saveTemplateDetails };
                                saveTemplateDetails.LocationAlreadyPinned = true;
                                this.setState({
                                    loading: false,
                                    counter: tempCounter,
                                    savedDocuments,
                                    inputDocument,
                                    saveTemplateDetails,
                                    showPreviousIcon: tempCounter === (this.state.selectedTemplates.length),
                                });

                            }).catch(error => {
                                console.error("Error while getting response json in pin location crated by change", error);
                                this.setState({
                                    loading: false
                                });
                            }).catch(error => {
                                console.error("Error while changing created by in pin location", error);
                                this.setState({
                                    loading: false
                                });
                            });

                        }).catch(error => {
                            console.error("Error while getting response json in pin location", error);
                            this.setState({
                                loading: false
                            });
                        }).catch(error => {
                            console.error("Error while pinning location", error);
                            this.setState({
                                loading: false
                            });
                        });

                    }
                    else {
                        this.setState({
                            loading: false,
                            counter: tempCounter,
                            savedDocuments,
                            inputDocument,
                            showPreviousIcon: tempCounter === (this.state.selectedTemplates.length),
                        });
                    }
                }).catch(error => {
                    console.error("Error while getting response json in change created by", error);
                    this.setState({
                        loading: false
                    });
                }).catch(error => {
                    console.error("Error while changing created by", error);
                    this.setState({
                        loading: false
                    });
                });
            }


        }).catch(error => {
            inputDocument.isValid = false;
            inputDocument.errorMessage = 'Save Failed, Server error occurred';
            console.error("Error while getting response json after save", error);
            this.setState({
                inputDocument,
                loading: false
            });
        }).catch(error => {
            inputDocument.isValid = false;
            inputDocument.errorMessage = 'Save Failed, Server error occurred';
            console.error("Error while saving file", error);
            this.setState({
                inputDocument,
                loading: false
            });
        });

    }

    public saveTemplateInOneDrive = (currentTemplate: any): void => {
        this.setState({
            loading: true,
        });
        let documentExtention = currentTemplate.Name.substr(currentTemplate.Name.indexOf("."));
        let inputDocument = { ...this.state.inputDocument };
        let savedDocumentName = inputDocument.name + documentExtention;

        let payloadCreateFile: IPayloadSaveOneDrive = {
            tenant: this.props.context.tid,
            SPOUrl: `https://${this.props.context.teamSiteDomain}`,
            FileName: inputDocument.name,
            sourceFileName: currentTemplate.ParentLocation + "/" + currentTemplate.Name,
            FolderName: this.state.saveTemplateDetails.DestFolderRelUrl,
            userGuidId: this.props.context.userObjectId,
        };

        // console.log(payloadCreateFile);


        fetch(SAVE_TEMPLATE_IN_ONE_DRIVE, {
            method: 'post',
            body: JSON.stringify(payloadCreateFile)
        }).then((response: any) => {
            return response.json();
        }).then((data) => {
            // console.log(data.d);
            if (data.error && data.error.code === "-2130575257, Microsoft.SharePoint.SPException") {
                inputDocument.isValid = false;
                inputDocument.errorMessage = 'Duplicate file name, file already exists';
                this.setState({
                    inputDocument,
                    loading: false
                });
            }
            else {

                let tempCounter = this.state.counter + 1;

                inputDocument.isValid = false;
                inputDocument.name = "";
                inputDocument.errorMessage = "";

                setTimeout(() => {
                    fetch(GET_TEMPLATE_FROM_ONE_DRIVE, {
                        method: 'post',
                        body: JSON.stringify(payloadCreateFile)
                    }).then((responseFileLocation: any) => {
                        return responseFileLocation.json();
                    }).then((templateData) => {

                        let savedDocuments = [...this.state.savedDocuments];
                        savedDocuments.push({
                            name: savedDocumentName,
                            url: templateData.webUrl,
                            clientAppLink: this.GetClientAppLinkforOneDrive(templateData.webUrl, this.state.saveTemplateDetails.DestFolderRelUrl, savedDocumentName)
                        });

                        if (this.state.saveTemplateDetails.IsPinnedLocation && !this.state.saveTemplateDetails.LocationAlreadyPinned) {
                            let payloadPinLocation = {
                                tenant: this.props.context.tid,
                                SPOUrl: `https://${this.props.context.teamSiteDomain}`,
                                PinnedType: "Onedrive",
                                DocumentLibrary: this.state.saveTemplateDetails.DestFolderRelUrl,
                                DocumentLibraryUrl: this.state.saveTemplateDetails.DestSite,
                                SiteUrl: this.state.saveTemplateDetails.DestSite

                            };
                            fetch(PIN_LOCATION_URL, {
                                headers: {
                                    'Accept': 'application/json',
                                    'Content-Type': 'application/json'
                                },
                                method: 'post',
                                body: JSON.stringify(payloadPinLocation)

                            }).then((responsePinLocation: any) => {
                                // console.log(responsePinLocation);
                                return responsePinLocation.json();
                            }).then((dataPinLocation) => {

                                let payloadChangeCreatedByPinLoc = {
                                    tenantName: this.state.tenantFullName,
                                    siteUrl: `https://${this.props.context.teamSiteDomain}`,
                                    listName: "DocsNodePinnedLocations",
                                    itemID: dataPinLocation.d.Id,
                                    emailid: this.state.loggedInUserEmail.toString()
                                };
                                fetch(CHANGE_CREATED_BY_URL, {
                                    headers: {
                                        'Accept': 'application/json',
                                        'Content-Type': 'application/json'
                                    },
                                    method: 'post',
                                    body: JSON.stringify(payloadChangeCreatedByPinLoc)

                                }).then((responseChangePinLoc: any) => {
                                    // console.log(responseChangePinLoc);
                                    return responseChangePinLoc.text();
                                }).then((dataChangeCreatedByPinLoc) => {

                                    let saveTemplateDetails = { ...this.state.saveTemplateDetails };
                                    saveTemplateDetails.LocationAlreadyPinned = true;


                                    this.setState({
                                        loading: false,
                                        counter: tempCounter,
                                        savedDocuments,
                                        inputDocument,
                                        saveTemplateDetails,
                                        showPreviousIcon: tempCounter === (this.state.selectedTemplates.length),
                                    });

                                }).catch(error => {
                                    console.error("Error while getting response json in pin location crated by change", error);
                                    this.setState({
                                        loading: false
                                    });
                                }).catch(error => {
                                    console.error("Error while changing created by in pin location", error);
                                    this.setState({
                                        loading: false
                                    });
                                });

                            }).catch(error => {
                                console.error("Error while getting response json in pin location", error);
                                this.setState({
                                    loading: false
                                });
                            }).catch(error => {
                                console.error("Error while pinning location", error);
                                this.setState({
                                    loading: false
                                });
                            });
                        }
                        else {
                            this.setState({
                                loading: false,
                                counter: tempCounter,
                                savedDocuments,
                                inputDocument,
                                showPreviousIcon: tempCounter === (this.state.selectedTemplates.length),
                            });
                        }
                    }).catch(error => {
                        console.error("Error while getting response json for File Locaion", error);
                        this.setState({
                            loading: false
                        });
                    }).catch(error => {
                        console.error("Error while parsing response json for File Locaion", error);
                        this.setState({
                            loading: false
                        });
                    });

                }, 10000);
            }

        }).catch(error => {
            inputDocument.isValid = false;
            inputDocument.errorMessage = 'Save Failed, Server error occurred';
            console.error("Error while getting response json after save", error);
            this.setState({
                inputDocument,
                loading: false
            });
        }).catch(error => {
            inputDocument.isValid = false;
            inputDocument.errorMessage = 'Save Failed, Server error occurred';
            console.error("Error while saving file", error);
            this.setState({
                inputDocument,
                loading: false
            });
        });
    }

    public onChangeInputDocument = (value: any): void => {
        let inputDocument = { ...this.state.inputDocument };

        inputDocument.name = value;
        if (value.match(/[?*:<>./|"]/)) {
            inputDocument.isValid = false;
            inputDocument.errorMessage = 'Invalid File Name. ? * :  < > . /  |" are not allowed';
        }
        else if (!value.trim()) {
            inputDocument.isValid = false;
            inputDocument.errorMessage = 'Only blank spaces are not allowed';
        }
        else {
            inputDocument.isValid = true;
            inputDocument.errorMessage = "";
        }
        this.setState({
            inputDocument
        });
    }

    public onClickPreviousIcon = (): void => {
        this.props.onClickPreviousIcon();
    }


    public GetClientAppLinkforTeams = (destUrl: string, destFolder: string, filename: string): string => {

        let fileUrl = "";
        if (destFolder.toLowerCase().includes("sites")) {
            fileUrl = `https://${this.props.context.teamSiteDomain}` + "/" + destFolder + "/" + filename;
        }
        else {
            fileUrl = destUrl + "/" + destFolder + "/" + filename;
        }

        fileUrl = fileUrl.replace("//", "/").replace("https:/", "https://");


        let clientAppLink: string = "";

        if (filename.toLowerCase().includes("ppt") || filename.toLowerCase().includes("potx")) {
            clientAppLink = "ms-powerpoint:ofe|u|" + fileUrl;
        }
        if (filename.toLowerCase().includes("xls") || filename.toLowerCase().includes("xltx")) {
            clientAppLink = "ms-excel:ofe|u|" + fileUrl;
        }
        if (filename.toLowerCase().includes("doc") || filename.toLowerCase().includes("dotx")) {
            clientAppLink = "ms-word:ofe|u|" + fileUrl;
        }
        return clientAppLink;
    }

    public GetClientAppLinkforOneDrive = (destUrl: string, destFolder: string, filename: string): string => {

        let fileUrl = "";
        if (destFolder !== "") {
            fileUrl = destUrl.substr(0, destUrl.toLowerCase().indexOf("/_layouts") + 1) + "/documents/" + destFolder + "/" + filename;
        }
        else {
            fileUrl = destUrl.substr(0, destUrl.toLowerCase().indexOf("/_layouts") + 1) + "/documents/" + filename;
        }

        fileUrl = fileUrl.replace("//", "/").replace("https:/", "https://");


        let clientAppLink: string = "";

        if (filename.toLowerCase().includes("ppt") || filename.toLowerCase().includes("potx")) {
            clientAppLink = "ms-powerpoint:ofe|u|" + fileUrl;
        }
        if (filename.toLowerCase().includes("xls") || filename.toLowerCase().includes("xltx")) {
            clientAppLink = "ms-excel:ofe|u|" + fileUrl;
        }
        if (filename.toLowerCase().includes("doc") || filename.toLowerCase().includes("dotx")) {
            clientAppLink = "ms-word:ofe|u|" + fileUrl;
        }
        return clientAppLink;
    }

    public documentModalHide = () => {

        this.setState({
            documentModalOpen: false
        });
    }

    public documentModalOpen = (link: string) => {
        this.setState({
            documentModalOpen: true,
            modalDocumentLink: link
        });
    }

    public contextMenuButton = (clientAppLink: string, webLink: string): IContextualMenuProps => {
        const menuProps: IContextualMenuProps = {
            items: [
                {
                    key: 'openDialog',
                    text: 'Open in Browser',
                    iconProps: { iconName: 'Website' },
                    onClick: (() => this.documentModalOpen(webLink))

                },
                {
                    key: 'openClientApp',
                    text: 'Open in Desktop App',
                    iconProps: { iconName: 'Document' },
                    onClick: (() => this.documentModalOpen(clientAppLink))
                },
            ],
        };

        return menuProps;
    }

    public render() {
        let { selectedTemplates, counter, documentModalOpen, modalDocumentLink } = this.state;
        let currentTemplate: any = {};
        if (selectedTemplates.length === counter) {
            currentTemplate = selectedTemplates[counter - 1];
        }
        else {
            currentTemplate = selectedTemplates[counter];
        }

        const theme = getTheme();
        const contentStyles = mergeStyleSets({
            container: {
                display: 'flex',
                flexFlow: 'column nowrap',
                alignItems: 'stretch',
            },
            header: [
                // tslint:disable-next-line:deprecation
                theme.fonts.xLargePlus,
                {
                    flex: '1 1 auto',
                    border: `0px`,
                    color: theme.palette.neutralPrimary,
                    display: 'flex',
                    alignItems: 'center',
                    fontWeight: FontWeights.semibold,
                    padding: '12px 12px 14px 24px',
                },
            ],
            body: {
                flex: '4 4 auto',
                padding: '0 24px 24px 24px',
                overflowY: 'hidden',
            },
        });
        const toggleStyles = { root: { marginBottom: '20px' } };
        const iconButtonStyles = {
            root: {
                color: theme.palette.neutralPrimary,
                marginLeft: 'auto',
                marginTop: '4px',
                marginRight: '2px',
            },
            rootHovered: {
                color: theme.palette.neutralDark,
            },
        };

        const cancelIcon: IIconProps = { iconName: 'Cancel' };

        return (
            <LoadingOverlay
                active={this.state.loading}
                spinner={<FadeLoader />}
                text="Creating Document...."
            >

                <div className="container">
                    {/* {this.state.showPreviousIcon &&
                        <div className="ms-Grid">
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg12">
                                    <h3>
                                        <i className="ms-Icon ms-Icon--NavigateBack"
                                            aria-hidden="true"
                                            onClick={() => this.onClickPreviousIcon()}
                                            style={{ cursor: "pointer" }}
                                        />
                                    </h3>

                                </div>
                            </div>
                        </div>
                    } */}
                    <div className="ms-Grid">
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg12">
                                <h3>Create The Document</h3>
                            </div>
                        </div>
                    </div>
                    <div className="ms-Grid">
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg12">
                                <h3>Template Name: {currentTemplate.Name}</h3>
                            </div>
                        </div>
                    </div>
                    <div className="ms-Grid">
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg12">
                                <h3>Enter the new file name without extension</h3>
                            </div>
                        </div>
                    </div>
                    <div className="ms-Grid">
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg12">
                                <TextField
                                    value={this.state.inputDocument.name}
                                    onChange={(_, newValue) => this.onChangeInputDocument(newValue)}
                                />
                            </div>
                        </div>
                    </div>
                    {!this.state.inputDocument.isValid && this.state.inputDocument.errorMessage &&
                        <div className="ms-Grid" style={{ color: "#f00" }}>
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg12">
                                    {this.state.inputDocument.errorMessage}
                                </div>
                            </div>
                        </div>
                    }
                    {this.state.counter < selectedTemplates.length &&
                        <div className="ms-Grid">
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg12">
                                    {this.state.counter + 1} of {selectedTemplates.length}
                                </div>
                            </div>
                        </div>
                    }
                    {this.state.counter === selectedTemplates.length &&
                        <div className="ms-Grid">
                            <div className="ms-Grid-row">
                                <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg12">
                                    {selectedTemplates.length} of {selectedTemplates.length}
                                </div>
                            </div>
                        </div>
                    }
                    <div className="ms-Grid">
                        <div className="ms-Grid-row">
                            <div className="ms-Grid-col ms-sm-12 ms-md6 ms-lg2">
                                <button className="normal-button"
                                    disabled={
                                        this.state.counter === selectedTemplates.length ||
                                        !this.state.inputDocument.isValid
                                    }

                                    onClick={
                                        () => this.state.saveTemplateDetails.SiteUrl === "Onedrive" ?
                                            this.saveTemplateInOneDrive(currentTemplate)
                                            :
                                            this.saveTemplateInTeams(currentTemplate)
                                    }
                                >
                                    {counter < (selectedTemplates.length - 1) ? `Next ` : `Save `}
                                    <i className="ms-Icon ms-Icon--Save" />
                                </button>
                            </div>
                        </div>
                    </div>
                    <div className="ms-Grid">

                        {this.state.savedDocuments.map((document) => {
                            return (
                                <div className="ms-Grid-row padRowBottom">
                                    <div className="ms-Grid-col ms-md5 ms-sm6 ms-lg5">
                                        Your document is saved <label style={{ fontWeight: 600 }} > {document.name}</label>
                                    </div>
                                    <div className="ms-Grid-col ms-md7 ms-sm6 ms-lg7">
                                        <a className="anchorbutton" href={document.url} target="_blank">Open in Browser</a>

                                        {/* Your document is saved <a style={{ color: "#ca5010" }} href="https://teams.microsoft.com/l/file/744B7DD5-CB4B-47C4-B588-457328D4D4A0?tenantId=641333af-c280-4e39-8f0c-1a52f0be8dc7&fileType=docx&objectUrl=https%3A%2F%2Fdocsnode.sharepoint.com%2Fsites%2FDocTemplateDEVTeam%2FShared%20Documents%2FGeneral%2Fabcd1.docx&baseUrl=https%3A%2F%2Fdocsnode.sharepoint.com%2Fsites%2FDocTemplateDEVTeam&serviceName=teams&threadId=19:af3e9385527f411fbb7ae0465d89068e@thread.skype&groupId=c0d0ce63-57d7-4321-be52-32e81980357c">{document.name}</a> */}

                                        <DefaultButton
                                            className="clientAppButton"
                                            text="Open in Desktop App"
                                            onClick={() => this.documentModalOpen(document.clientAppLink)}

                                        />
                                    </div>
                                </div>
                            );
                        })}
                    </div>

                </div>
                <div className="modalContainer">
                    <Modal
                        titleAriaId="modal"
                        isOpen={documentModalOpen}
                        onDismiss={() => this.documentModalHide()}
                        isBlocking={false}
                        containerClassName={contentStyles.container}
                    >
                        <div className={contentStyles.header}>
                            <IconButton
                                styles={iconButtonStyles}
                                iconProps={cancelIcon}
                                ariaLabel="Close popup modal"
                                onClick={() => this.documentModalHide()}
                            />
                        </div>
                        <div className={contentStyles.body}>
                            <iframe src={modalDocumentLink} style={{ visibility: "hidden", height: "0px", width: "0px!important" }}>

                            </iframe>
                            <div className="card">

                                <div className="container">
                                    <DefaultButton
                                        className="dialogClose"
                                        text="Close"
                                        onClick={() => this.documentModalHide()}

                                    />
                                </div>
                            </div>
                        </div>
                    </Modal>
                </div>
            </LoadingOverlay >


        );
    }

}
