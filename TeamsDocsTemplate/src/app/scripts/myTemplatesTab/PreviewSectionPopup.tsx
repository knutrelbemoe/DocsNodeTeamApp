import * as React from "react";
import * as util from "./../utility/formatter";
import { IconButton } from "office-ui-fabric-react";

export class PreviewSectionPopup extends React.Component<any, any> {

    constructor(props: any) {
        super(props, {});
        this.state = {
            templates: [],
            selectedTemplate: {
                Id: "",
                ModifiedOn: new Date().toDateString(),
                ModifiedBy: "",
                ContentType: "",
                imgURL: "",
                altImgUrl: ""
            },
            currentTemplateIndex: 0
        };
    }

    public componentWillMount() {
        let template = this.props.selectedTemplate;
        let indexOfSelectedTemplate = this.props.templates.map(data => data.UniqueId).indexOf(template.UniqueId);
        this.setState({
            selectedTemplate: template,
            templates: this.props.templates,
            currentTemplateIndex: indexOfSelectedTemplate
        });
    }

    public componentDidMount() {
    }

    public componentWillReceiveProps(nextProps: any) {
        if (nextProps.selectedTemplate !== this.state.selectedTemplate) {
            let template = nextProps.selectedTemplate;
            let indexOfSelectedTemplate = nextProps.templates.map(data => data.UniqueId).indexOf(template.UniqueId);
            this.setState({
                selectedTemplate: template,
                templates: nextProps.templates,
                currentTemplateIndex: indexOfSelectedTemplate
            });
        }
    }

    public onErrorImageLoad = () => {
        let currentTemplate = { ...this.state.selectedTemplate };

        currentTemplate.imgUrl = currentTemplate.altImgUrl;

        this.setState({
            selectedTemplate: currentTemplate
        });
    }

    public onNavigateTemplate = (flag: number): void => {

        let currentTemplateIndex = this.state.currentTemplateIndex;
        switch (flag) {
            case -1:
                currentTemplateIndex -= 1;
                if (currentTemplateIndex === -1) {
                    currentTemplateIndex = this.state.templates.length - 1;
                }
                break;

            default:
                currentTemplateIndex += 1;
                if (currentTemplateIndex === (this.state.templates.length)) {
                    currentTemplateIndex = 0;
                }
        }

        let currentTemplates = [...this.state.templates];
        let indexOfSelectedTemplate = currentTemplateIndex;
        let selectedTemplate = currentTemplates[indexOfSelectedTemplate];

        this.setState({
            currentTemplateIndex,
            selectedTemplate,
        }, () => this.props.selectTemplateFromPopup(selectedTemplate));
    }

    public render() {

        let { selectedTemplate } = this.state;

        return (
            <div className="container">
                <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-md2 ms-sm12 ms-lg1">
                            <i className="ms-Icon ms-Icon--Back"
                                aria-hidden="true"
                                onClick={() => this.onNavigateTemplate(-1)}
                                style={{ cursor: "pointer", position: "relative", top: "5em", fontSize: "2em" }}
                            />
                        </div>
                        <div className="ms-Grid-col ms-md4 ms-sm12 ms-lg5">
                            <img src={selectedTemplate.imgUrl}
                                onError={() => this.onErrorImageLoad()}
                                alt="" className="preview-image" />
                        </div>
                        <div className="ms-Grid-col ms-md4 ms-sm12 ms-lg5">
                            <table className="table-preview-details">
                                <tbody>
                                    <tr>
                                        <td>
                                            <strong>Template:</strong>
                                        </td>
                                        <td>
                                            {selectedTemplate.Name}
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <strong>Modified:</strong>
                                        </td>
                                        <td>
                                            {util.USformatDate(selectedTemplate.TimeLastModified)}
                                        </td>
                                    </tr>
                                    {/* <tr>
                                        <td>
                                            <strong>Modified By:</strong>
                                        </td>
                                        <td>
                                            {selectedTemplate.Editor.Title}
                                        </td>
                                    </tr> */}
                                    {/* <tr>
                                        <td>
                                            <strong>Content Type:</strong>
                                        </td>
                                        <td>
                                            {selectedTemplate.ContentType.Name}
                                        </td>
                                    </tr> */}
                                </tbody>
                            </table>
                            <button
                                className="normal-button"
                                onClick={() => this.props.onClosePreviewPopup(true)}>
                                Create <i className="fa fa-plus-circle" aria-hidden="true"></i>
                            </button>
                        </div>
                        <div className="ms-Grid-col ms-md2 ms-sm12 ms-lg1">
                            <i className="ms-Icon ms-Icon--Forward"
                                aria-hidden="true"
                                onClick={() => this.onNavigateTemplate(1)}
                                style={{ cursor: "pointer", position: "relative", top: "5em", fontSize: "2em" }}
                            />
                        </div>
                    </div>
                </div>
            </div>
        );
    }

}
