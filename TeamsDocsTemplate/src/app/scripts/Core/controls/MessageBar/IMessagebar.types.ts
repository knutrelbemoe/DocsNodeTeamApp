import { MessageBarType } from "office-ui-fabric-react";

export interface IAppMessageBarProps {
    defaultShowTime?: number;
    message: string;
    moreDetailMessage?: string;
    messageType?: MessageBarType;
    linkUrl?: string;
    linkDisplayText?: string;
    isMultiline?: boolean;
    onDismissOverride: () => void;
}

export interface IAppMessageBarState {
    defaultShowTime?: number;
    message: string;
    moreDetailMessage?: string;
    messageType?: MessageBarType;
    linkUrl?: string;
    linkDisplayText?: string;
    isMultiline?: boolean;
}
