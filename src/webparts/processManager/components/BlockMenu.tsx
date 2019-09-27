import * as React from "react";
import ComposeStyles from "./styles/composeStyles";
import {
  Stack,
  IStackTokens,
  ActionButton,
  IButtonStyles
} from "office-ui-fabric-react";
import cardStyle from "./styles/CardStyle.module.scss";

const buttonStyle = () => {
  const customStyle: Partial<IButtonStyles> = {};

  customStyle.root = {
    margin: "auto",
    color: "#3b3b3b"
  };
  customStyle.rootHovered = {
    color: "#3b3b3b"
  };

  customStyle.icon = {
    color: "rgb(21, 164, 232)"
  };

  return customStyle;
};

const iconStyle = {
  fontSize: 40
};

const sectionStackTokens: IStackTokens = { childrenGap: 5 };
const wrapStackTokens: IStackTokens = { childrenGap: 20 };

export interface IBlockMenuProps {
  onComponentChange(componentName: string): void;
}

export default class BlockMenu extends React.Component<IBlockMenuProps, {}> {
  public render(): JSX.Element {
    const { onComponentChange } = this.props;
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
          <div
            className={cardStyle.card}
            onClick={() => onComponentChange("policies")}
          >
            <ActionButton
              iconProps={{ iconName: "ReportLibrary", style: iconStyle }}
              allowDisabledFocus
              styles={buttonStyle()}
            >
              Policies
            </ActionButton>
          </div>
          <div
            className={cardStyle.card}
            onClick={() => onComponentChange("pageBuilder")}
          >
            <ActionButton
              iconProps={{ iconName: "BuildDefinition", style: iconStyle }}
              allowDisabledFocus
              styles={buttonStyle()}
            >
              Policy Page Builder
            </ActionButton>
          </div>
          <div
            className={cardStyle.card}
            onClick={() => onComponentChange("policyAssignment")}
          >
            <ActionButton
              iconProps={{ iconName: "Assign", style: iconStyle }}
              allowDisabledFocus
              styles={buttonStyle()}
            >
              Policy Assignment
            </ActionButton>
          </div>

          <div
            className={cardStyle.card}
            onClick={() => onComponentChange("taskManager")}
          >
            <ActionButton
              iconProps={{ iconName: "TaskSolid", style: iconStyle }}
              allowDisabledFocus
              styles={buttonStyle()}
            >
              Task manager
            </ActionButton>
          </div>
          <div
            className={cardStyle.card}
            onClick={() => onComponentChange("groups")}
          >
            <ActionButton
              iconProps={{ iconName: "HomeGroup", style: iconStyle }}
              allowDisabledFocus
              styles={buttonStyle()}
            >
              Group manager
            </ActionButton>
          </div>
          {/* <ActionButton
            iconProps={{ iconName: "Add", style: iconStyle }}
            allowDisabledFocus
            styles={{
              root: {
                boxShadow: "0px 1px 3px rgba(0, 0, 0, 0.3)",
                border: "1px #f3f3f3 solid",
                width: "220px",
                height: "70px",
                background: "#ffff"
              }
            }}
          >
            Group manager
          </ActionButton> */}
        </Stack>
      </Stack>
    );
  }
}
