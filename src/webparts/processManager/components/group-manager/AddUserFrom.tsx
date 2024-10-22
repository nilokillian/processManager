import * as React from "react";
import { toast } from "react-toastify";

import SharePointService from "../../../../services/SharePoint/SharePointService";
import {
  DefaultButton,
  Dialog,
  PrimaryButton,
  ContextualMenu,
  DialogType,
  Panel,
  PanelType,
  Stack,
  IStackTokens,
  TextField
} from "office-ui-fabric-react";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 5
};

export interface IUserFormProps {
  groupId: number;
  isOpenForm: boolean;
  onCloseForm(): void;
  updateGroupMembers(): void;
}

export interface IUserFormState {
  users: string[];
  errors: object;
  loading: boolean;
}

export default class UserForm extends React.Component<
  IUserFormProps,
  IUserFormState
> {
  private _dragOptions = {
    moveMenuItemText: "Move",
    closeMenuItemText: "Close",
    menu: ContextualMenu
  };

  constructor(props: IUserFormProps) {
    super(props);

    this.state = {
      users: [],
      errors: {},
      loading: false
    };
  }

  public async componentDidMount() {}

  public render(): JSX.Element {
    const { isOpenForm, onCloseForm } = this.props;
    const { loading, users } = this.state;

    return (
      <div>
        <Panel
          isOpen={isOpenForm}
          type={PanelType.custom}
          customWidth="420px"
          onDismiss={onCloseForm}
          headerText="Add users"
          closeButtonAriaLabel="Close"
          //onRenderHeader={this._onRenderHeaderContent}
          onRenderFooterContent={this._onRenderFooterContent}
          //styles={ComponentStyles.formPanelStyle}
        >
          <Stack
            //  styles={stackContainerStyles}
            tokens={itemAlignmentsStackTokens}
          >
            <PeoplePicker
              context={SharePointService.context}
              titleText="Users"
              personSelectionLimit={10}
              groupName={""} // Leave this blank in case you want to filter from all users
              isRequired={true}
              selectedItems={this._getPeoplePickerItems}
              //defaultSelectedUsers={[Email]}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}

              // peoplePickerCntrlclassName={
              //   styles[ComponentStyles.peoplePickerStyle()]
              // }
              //styles={{backgroundColor: "red"}}
            />
          </Stack>
        </Panel>
      </div>
    );
  }

  private _onRenderFooterContent = () => {
    const { loading, errors } = this.state;
    return (
      <div>
        <PrimaryButton
          onClick={this.submitForm}
          text="Save"
          disabled={loading}
        />
      </div>
    );
  };

  private _getPeoplePickerItems = (items: any[]) => {
    //const { users } = this.state;

    const users = items.map(item => item.id);

    this.setState({ users });
  };

  // private _onChangeTextInput = (
  //   e: React.FormEvent<HTMLInputElement>,
  //   newValue?: string
  // ) => {
  //   // const currentFieldName = e.target["id"];
  //   const { data } = this.state;

  //   data.groupName = newValue;

  //   this.setState({ data });
  // };

  //   private _getPeoplePickerItems = (items: any[]) => {
  //     const { data } = this.state;

  //     if (items.length === 1) {
  //       currentItem.Display_Name = items[0].text;
  //       currentItem.Email = items[0].secondaryText.toLowerCase();
  //       currentItem.First_Name = items[0].text.split(" ")[0];
  //       currentItem.Last_Name = items[0].text.split(" ")[1];
  //     } else if (items.length === 0) {
  //       currentItem.Display_Name = "";
  //       currentItem.Email = "";
  //       currentItem.First_Name = "";
  //       currentItem.Last_Name = "";
  //     }

  //     this.setState({ currentItem });
  //   };

  //   private _showDeleteBtnDialog = () => {
  //     this.setState({ isDeleteCalloutVisible: true });
  //   };

  //   private _closeDeleteBtnDialog = () => {
  //     this.setState({ isDeleteCalloutVisible: false });
  //   };

  public submitForm = async () => {
    const { onCloseForm, groupId, updateGroupMembers } = this.props;
    const { users } = this.state;
    this.setState({ loading: true });

    try {
      await SharePointService.pnp_addGroupMember(groupId, users);
      onCloseForm();
      updateGroupMembers();
      toast.success("created");
    } catch (error) {
      toast.error("error");
      onCloseForm();
      throw error;
    }
    this.setState({ loading: false });
  };
}
