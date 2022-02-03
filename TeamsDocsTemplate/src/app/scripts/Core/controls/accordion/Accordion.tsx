import { css } from "@uifabric/utilities/lib/css";
import { DefaultButton, IIconProps } from "office-ui-fabric-react";
import * as React from "react";
import { IAccordionProps, IAccordionState, Accordstyles } from "./index";
import { FontIcon } from 'office-ui-fabric-react/lib/Icon';

/**
 * Icon styles. Feel free to change them
 */
const collapsedIcon: IIconProps = { iconName: "ChevronRight", className: "accordionChevron" };
const expandedIcon: IIconProps = { iconName: "ChevronDown", className: "accordionChevron" };
const blockIcon: IIconProps = { iconName: "Blocked2", className: "accordionChevron" };

const privateIcon: IIconProps = { iconName: "Lock", className: "accordionLock" };
const TabIcon: IIconProps = { iconName: "BrowserTab", className: "accordionBrowserTab" };
const folderIcon: IIconProps = { iconName: "FabricFolder", className: "accordionFolder" };

export class Accordion extends React.Component<IAccordionProps, IAccordionState> {
  // tslint:disable-next-line: variable-name
  private _drawerDiv: HTMLDivElement | undefined = undefined;
  constructor(props: IAccordionProps) {
    super(props);

    this.state = {
      expanded: props.defaultCollapsed ? !props.defaultCollapsed : false,
      // ChildCount: props.childCount ? props.childCount : (props.childCount === 0 ? props.childCount : -1)
    };

  }

  public componentWillMount() {

    this.getFolderType();
  }

  public getFolderType = () => {

    switch (this.props.foldertype) {
      case "#microsoft.graph.channel.private":
        this.setState({
          foldertype: "Private Channel",
          iconName: "Lock"
        });
        break;
      case "tab.sharepoint.folder":
        this.setState({
          foldertype: "Channel Tab",
          iconName: "BrowserTab"
        });
        break;
      case "sharepoint.folder":
        this.setState({
          foldertype: "SP Folder",
          iconName: "FabricFolder"
        });
        break;
      default:
        break;

    }

  }

  public render(): React.ReactElement<IAccordionProps> {
    return (
      <div className={css("accordion", this.props.className)}>

        <DefaultButton
          toggle
          checked={this.state.expanded}
          text={this.props.title}
          iconProps={this.props.childCount === 0 ? blockIcon : (this.state.expanded ? expandedIcon : collapsedIcon)}
          onClick={() => {
            if (this.props.childCount !== 0) {
              this.setState({
                expanded: !this.state.expanded
              });
            }
            this.props.onExpand();
          }}

          aria-expanded={this.state.expanded}
          aria-controls={this._drawerDiv && this._drawerDiv.id}

        >

          {this.state.foldertype &&
            <span className="badge badge-secondary">
              <FontIcon iconName={this.state.iconName} className="badge-icon" />
              {this.state.foldertype}</span>
          }
        </DefaultButton>
        {this.state.expanded &&
          <div className={"drawer"} ref={(el) => { this._drawerDiv = el || undefined; }}>
            {this.props.children}
          </div>
        }
      </div>
    );
  }
}

