import * as React from "react";
import * as util from "./../utility/formatter";

export class NextSection extends React.Component<any, any> {

    constructor(props: any) {
        super(props, {});
        this.state = {
            defaultSiteCollection: ""
        };
    }

    public componentWillMount() {

        this.setState({
            defaultSiteCollection: this.props.defaultSiteCollection
        });
    }

    public componentDidMount() {
    }

    public componentWillReceiveProps(nextProps: any) {
        if (nextProps.defaultSiteCollection !== this.state.defaultSiteCollection) {
            this.setState({
                defaultSiteCollection: nextProps.defaultSiteCollection
            });
        }
    }

    public render() {

        let { defaultSiteCollection } = this.state;

        return (
            <div className="container">
                <div className="ms-Grid">
                    <div className="ms-Grid-row">
                        <div className="ms-Grid-col ms-md12 ms-sm12 ms-lg12">
                            <div className="default-site-collection-section">
                                <h2>Default Site Collection:</h2>
                                <h2>{defaultSiteCollection}</h2>
                            </div>
                        </div>
                    </div>
                </div>

            </div>
        );
    }

}
