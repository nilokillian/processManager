import * as React from "react";
//import { toast } from "react-toastify";

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

export interface IGroupFormProps {
  isOpenForm: boolean;
  onCloseForm(): void;
}

export interface IGroupOwner {
  id: string;
  name: string;
}

export interface IGroupData {
  title: string;
  owner: IGroupOwner;
  groupName: string;
}

export interface IGroupFormState {
  data: IGroupData;
  errors: object;
  loading: boolean;
}

export default class GroupForm extends React.Component<
  IGroupFormProps,
  IGroupFormState
> {
  private _dragOptions = {
    moveMenuItemText: "Move",
    closeMenuItemText: "Close",
    menu: ContextualMenu
  };

  constructor(props: IGroupFormProps) {
    super(props);

    this.state = {
      data: {} as IGroupData,
      errors: {},
      loading: false
    };
  }

  public async componentDidMount() {}

  public render(): JSX.Element {
    const { isOpenForm, onCloseForm } = this.props;
    const { loading, data } = this.state;

    return (
      <div>
        <Panel
          isOpen={isOpenForm}
          type={PanelType.custom}
          customWidth="420px"
          onDismiss={onCloseForm}
          headerText="New group"
          closeButtonAriaLabel="Close"
          //onRenderHeader={this._onRenderHeaderContent}
          onRenderFooterContent={this._onRenderFooterContent}
          //styles={ComponentStyles.formPanelStyle}
        >
          <Stack
            //  styles={stackContainerStyles}
            tokens={itemAlignmentsStackTokens}
          >
            <TextField
              id="Title"
              label="Title"
              value={data.groupName}
              onChange={this._onChangeTextInput}
              //styles={ComponentStyles.textInputStyle()}
              disabled={loading}
              required={true}
            />

            <PeoplePicker
              context={SharePointService.context}
              titleText="Group owner"
              personSelectionLimit={1}
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

        {/* <Dialog
          hidden={!isDeleteCalloutVisible}
          onDismiss={this._closeDeleteBtnDialog}
          maxWidth={670}
          dialogContentProps={{
            type: DialogType.close,
            title: "Are you sure ?",
            subText:
              "This contact might be connected to existing tracking records. Breaking connections could cause unexpected issues"
          }}
          modalProps={{
            titleAriaId: this._labelId,
            dragOptions: this._dragOptions,
            isBlocking: false
            // styles: { main: { maxWidth: 750 } }
          }}
        >
          <div style={{ display: "flex", justifyContent: "center" }}>
            <DefaultButton
              style={{ backgroundColor: "#dc224d", color: "white" }}
              disabled={loading}
              onClick={this.delete}
              text="Delete"
            />
          </div>
        </Dialog> */}
      </div>
    );
  }

  private _onRenderFooterContent = () => {
    const { loading, errors } = this.state;
    return (
      <div>
        <DefaultButton
          disabled={loading}
          // onClick={this._showDeleteBtnDialog}
          text="Delete"
        />

        <PrimaryButton
          onClick={this.submitForm}
          text="Save"
          disabled={loading}
        />
      </div>
    );
  };

  private _getPeoplePickerItems = (items: any[]) => {
    const { data } = this.state;
    data.owner = { name: items[0].text, id: items[0].id };

    this.setState({ data });
  };

  private _onChangeTextInput = (
    e: React.FormEvent<HTMLInputElement>,
    newValue?: string
  ) => {
    // const currentFieldName = e.target["id"];
    const { data } = this.state;

    data.groupName = newValue;

    this.setState({ data });
  };

  private _createNewGroup = async () => {
    const { groupName, owner } = this.state.data;

    // await SharePointService.pnp_createGroup(
    //   groupName,
    //   SharePointService.context.pageContext.user
    // );

    await SharePointService.pnp_createGroupV2(groupName);
  };

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
    const { onCloseForm } = this.props;
    this.setState({ loading: true });

    try {
      this._createNewGroup();
      this.setState({ loading: false });
    } catch (error) {
      console.log(error);
      //toast.error("error");
      this.setState({ loading: false });
      onCloseForm();
      return;
    }
  };

  public delete = async () => {
    this.setState({ loading: true });
    const { onCloseForm } = this.props;

    const { data } = this.state;
    try {
      onCloseForm();
      //toast.success(`${currentItem.Display_Name} has been deleted`);
      this.setState({ loading: false });

      return;
    } catch (error) {
      console.log(error);
      //toast.error("error");
      onCloseForm();
      return;
    }
  };

  public _getCurrentItem = async () => {};
}
