import * as React from "react";
import {
  Stack,
  IStackTokens,
  ActionButton,
  IButtonStyles
} from "office-ui-fabric-react";

const buttonStyle = () => {
  const customStyle: Partial<IButtonStyles> = {};

  customStyle.root = {
    margin: "auto",
    marginBottom: 22,
    color: "#3b3b3b",
    boxShadow: "0px 1px 3px rgba(0, 0, 0, 0.3)",
    border: "1px #f3f3f3 solid",
    width: 220,
    height: 70,
    background: "#ffff"
  };
  customStyle.rootHovered = {
    color: "#3b3b3b",
    boxShadow: "0px 3px 5px rgba(21, 164, 232, 0.3)"
  };

  customStyle.icon = {
    color: "rgb(21, 164, 232)"
  };

  // box-shadow: 0px 1px 3px rgba(0, 0, 0, 0.3);
  // border: 1px #f3f3f3 solid;
  // width: 220px;
  // height: 70px;
  // background: #ffff;
  // &:hover {
  //   box-shadow: 0px 3px 5px rgba(21, 164, 232, 0.3);
  // }

  return customStyle;
};

const iconStyle = {
  fontSize: 40
};

const wrapStackTokens: IStackTokens = { childrenGap: 20 };

export interface IBlockMenuProps {
  activeComponents: { title: string; sortOrderNumber: number };
  onComponentChange(componentName: string, sortOrderNumber: number): void;
}

export default class BlockMenu extends React.Component<IBlockMenuProps, {}> {
  public render(): JSX.Element {
    const { onComponentChange, activeComponents } = this.props;
    // background: "#f3f3f3"
    return (
      <Stack>
        <Stack
          horizontal
          horizontalAlign="space-evenly"
          wrap
          tokens={wrapStackTokens}
          style={{ marginBottom: 30, marginTop: 30 }}
        >
          <ActionButton
            iconProps={{ iconName: "BuildDefinition", style: iconStyle }}
            allowDisabledFocus
            styles={buttonStyle()}
            onClick={() => onComponentChange("pageBuilder", 1)}
            style={
              activeComponents.title === "pageBuilder"
                ? {
                    color: "#3b3b3b",
                    boxShadow: "0px 3px 13px rgba(21, 164, 232, 0.3)"
                  }
                : {}
            }
          >
            Policy Page Builder
          </ActionButton>

          <ActionButton
            iconProps={{ iconName: "ReportLibrary", style: iconStyle }}
            allowDisabledFocus
            styles={buttonStyle()}
            onClick={() => onComponentChange("policies", 2)}
            style={
              activeComponents.title === "policies"
                ? {
                    color: "#3b3b3b",
                    boxShadow: "0px 3px 13px rgba(21, 164, 232, 0.3)"
                  }
                : {}
            }
          >
            Policies
          </ActionButton>

          <ActionButton
            iconProps={{ iconName: "Assign", style: iconStyle }}
            allowDisabledFocus
            styles={buttonStyle()}
            onClick={() => onComponentChange("policyAssignment", 3)}
            style={
              activeComponents.title === "policyAssignment"
                ? {
                    color: "#3b3b3b",
                    boxShadow: "0px 3px 13px rgba(21, 164, 232, 0.3)"
                  }
                : {}
            }
          >
            Policy Assignment
          </ActionButton>

          <ActionButton
            iconProps={{ iconName: "TaskSolid", style: iconStyle }}
            allowDisabledFocus
            styles={buttonStyle()}
            onClick={() => onComponentChange("taskManager", 4)}
            style={
              activeComponents.title === "taskManager"
                ? {
                    color: "#3b3b3b",
                    boxShadow: "0px 3px 13px rgba(21, 164, 232, 0.3)"
                  }
                : {}
            }
          >
            Task manager
          </ActionButton>

          <ActionButton
            iconProps={{ iconName: "HomeGroup", style: iconStyle }}
            allowDisabledFocus
            styles={buttonStyle()}
            onClick={() => onComponentChange("groups", 5)}
            style={
              activeComponents.title === "groups"
                ? {
                    color: "#3b3b3b",
                    boxShadow: "0px 3px 13px rgba(21, 164, 232, 0.3)"
                  }
                : {}
            }
          >
            Group manager
          </ActionButton>

          <ActionButton
            iconProps={{ iconName: "BookmarkReport", style: iconStyle }}
            allowDisabledFocus
            styles={buttonStyle()}
            onClick={() => onComponentChange("reports", 6)}
            style={
              activeComponents.title === "reports"
                ? {
                    color: "#3b3b3b",
                    boxShadow: "0px 3px 13px rgba(21, 164, 232, 0.3)"
                  }
                : {}
            }
          >
            Reports
          </ActionButton>
        </Stack>
      </Stack>
    );
  }
}
