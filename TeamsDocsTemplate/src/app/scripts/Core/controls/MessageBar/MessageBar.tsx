import * as React from 'react';
import { MessageBarButton, Link, Stack, StackItem, MessageBar, MessageBarType, ChoiceGroup, IStackProps } from 'office-ui-fabric-react';
import { IAppMessageBarProps, IAppMessageBarState } from "./index";
interface IExampleProps {
    resetChoice?: () => void;
}

const verticalStackProps: IStackProps = {
    styles: { root: { overflow: 'hidden', width: '100%' } },
    tokens: { childrenGap: 20 }
};

const choiceGroupStyles = {
    label: {
        maxWidth: 250
    }
};




export class AppMessageBar extends React.Component<IAppMessageBarProps, IAppMessageBarState> {

    constructor(props: IAppMessageBarProps) {
        super(props);

        this.state = {
            message: props.message ? props.message : '',
            moreDetailMessage: props.moreDetailMessage ? props.moreDetailMessage : '',
            defaultShowTime: props.defaultShowTime ? props.defaultShowTime : 900,
            isMultiline: props.isMultiline ? props.isMultiline : false,
            linkUrl: props.linkUrl ? props.linkUrl : '',
            linkDisplayText: props.linkDisplayText ? props.linkDisplayText : '',
            messageType: props.messageType ? props.messageType : MessageBarType.info,
        };
    }


    public resetChoice = () => this.props.onDismissOverride();

    public render(): React.ReactElement<IAppMessageBarProps> {

        const MessageBarLoader = () => (
            <MessageBar
                isMultiline={this.state.isMultiline}
                onDismiss={this.resetChoice}
                dismissButtonAriaLabel="Close"
                truncated={this.state.isMultiline}
                overflowButtonAriaLabel="See more"
                messageBarType={this.state.messageType}

            >
                {this.state.moreDetailMessage ? "<b>" + this.state.message + "</b>. " + this.state.moreDetailMessage : this.state.message}
                {this.state.linkUrl && <Link href={this.state.linkUrl} target="_blank">
                    {this.state.linkDisplayText}
                </Link>}

            </MessageBar>
        );

        return (
            <Stack {...verticalStackProps}>
                {<MessageBarLoader />}
            </Stack>

        );
    }
}
