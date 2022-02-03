import * as React from "react";
import * as util from "./../utility/formatter";

export class PreviewSection extends React.Component<any, any> {

    constructor(props: any) {
        super(props, {});
        this.state = {
            selectedTemplate: {
                ModifiedOn: new Date().toDateString(),
                ModifiedBy: "",
                ContentType: "",
                imgURL: "",
                altImgUrl: ""
            }
        };
    }

    public componentWillMount() {
        let template = this.props.selectedTemplate;
        let templateWriterDetails = {
            ModifiedOn: util.USformatDate(template.Modified),
            ModifiedBy: template.Editor.Title,
            ContentType: template.ContentType.Name,
            imgUrl: template.imgUrl,
            altImgUrl: template.altImgUrl
        };
        this.setState({
            selectedTemplate: templateWriterDetails
        });
    }

    public componentDidMount() {
    }

    public componentWillReceiveProps(nextProps: any) {
        if (nextProps.selectedTemplate !== this.state.selectedTemplate) {
            let template = nextProps.selectedTemplate;
            let templateWriterDetails = {
                ModifiedOn: util.USformatDate(template.Modified),
                ModifiedBy: template.Editor.Title,
                ContentType: template.ContentType.Name,
                imgUrl: template.imgUrl,
                altImgUrl: template.altImgUrl
            };
            this.setState({
                selectedTemplate: templateWriterDetails
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

    public render() {

        let { selectedTemplate } = this.state;

        return (
            <div className="container">
                <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-md6 ms-sm8 ms-lg4">
                            <div className="image-container-preview">
                                <img src={selectedTemplate.imgUrl}
                                    onError={() => this.onErrorImageLoad()}
                                    alt="" className="preview-image" />
                            </div>
                        </div>
                    </div>
                </div>
                <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg6">
                            <table className="table-preview-details">
                                <tbody>
                                    <tr>
                                        <td>
                                            <strong>Modified:</strong>
                                        </td>
                                        <td>
                                            {selectedTemplate.ModifiedOn}
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <strong>Modified By:</strong>
                                        </td>
                                        <td>
                                            {selectedTemplate.ModifiedBy}
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <strong>Content Type:</strong>
                                        </td>
                                        <td>
                                            {selectedTemplate.ContentType}
                                        </td>
                                    </tr>
                                </tbody>
                            </table>

                        </div>
                    </div>

                </div>
            </div>
        );
    }

}
