import * as React from "react";
//import { toast } from "react-toastify";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import SharePointService from "../../../../services/SharePoint/SharePointService";
import {
  DefaultButton,
  Dialog,
  PrimaryButton,
  ContextualMenu,
  DialogType,
  Text,
  Panel,
  PanelType,
  Stack,
  IStackTokens,
  TextField,
  Dropdown,
  IDropdownOption
} from "office-ui-fabric-react";
import { IPolicy } from "./PageBuilder";

const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 5
};

export interface IPolicyFormProps {
  templates: any[];
  isOpenForm: boolean;
  onCloseForm(): void;
}

export interface IPolicyFormState {
  policy: IPolicy;
  //selectedTemplate: { key: string | number | undefined };
  errors: object;
  loading: boolean;
}

export default class PolicyForm extends React.Component<
  IPolicyFormProps,
  IPolicyFormState
> {
  private _dragOptions = {
    moveMenuItemText: "Move",
    closeMenuItemText: "Close",
    menu: ContextualMenu
  };

  constructor(props: IPolicyFormProps) {
    super(props);

    this.state = {
      policy: {} as IPolicy,
      //selectedTemplate: { key: "" },
      errors: {},
      loading: false
    };
  }

  public async componentDidMount() {}

  public render(): JSX.Element {
    const { isOpenForm, onCloseForm } = this.props;
    const { loading, policy } = this.state;

    return (
      <div>
        <Panel
          isOpen={isOpenForm}
          type={PanelType.custom}
          customWidth="420px"
          onDismiss={onCloseForm}
          headerText={"Create policy"}
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
              value={policy.Title}
              onChange={this._onChangeTextInput}
              //styles={ComponentStyles.textInputStyle()}
              disabled={loading}
              required={true}
            />

            <Dropdown
              id="templateId"
              label="Template"
              placeholder="Select template"
              options={this._getTemplateOptions()}
              selectedKey={
                policy.PolicyPagesTitle ? policy.PolicyPagesTitle : undefined
              }
              disabled={loading}
              required={true}
              onChange={this._onTemplateChange}
            />

            <PeoplePicker
              context={SharePointService.context}
              titleText="Contact"
              personSelectionLimit={1}
              groupName={""} // Leave this blank in case you want to filter from all users
              isRequired={true}
              selectedItems={this._getPeoplePickerItems}
              // defaultSelectedUsers={[Email]}
              showHiddenInUI={true}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}

              // peoplePickerCntrlclassName={
              //   styles[ComponentStyles.peoplePickerStyle()]
              // }
              //styles={{backgroundColor: "red"}}
            />

            {/* <PrimaryButton
              onClick={async () =>
                await SharePointService.pnp_createField(policy.Title)
              }
              text="field"
              disabled={loading}
            /> */}
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
  private _getPeoplePickerItems = async (items: any[]) => {
    const { policy } = this.state;
    const user = await SharePointService.pnp_getUserId(items[0].secondaryText);
    policy.PolicyOwnerId = user.Id;
    this.setState({ policy });
  };

  private _onTemplateChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    const { policy } = this.state;
    policy.PolicyPagesTitle = item.text;
    this.setState({ policy });
  };

  private _getTemplateOptions = (): IDropdownOption[] => {
    const { templates } = this.props;

    return templates.map(t => {
      return {
        key: t.name.split(".")[0],
        text: t.name.split(".")[0]
      } as IDropdownOption;
    });
  };

  private _onRenderFooterContent = () => {
    const { loading, errors } = this.state;
    const { onCloseForm } = this.props;
    return (
      <div>
        <PrimaryButton
          onClick={this.submitForm}
          text={"Save"}
          disabled={loading}
        />
      </div>
    );
  };

  private _onChangeTextInput = (
    e: React.FormEvent<HTMLInputElement>,
    newValue?: string
  ) => {
    // const currentFieldName = e.target["id"];
    const { policy } = this.state;
    policy.Title = newValue;

    this.setState({ policy });
  };

  public submitForm = async () => {
    // const fields = [
    //   "UserId",
    //   "Email",
    //   "Status",
    //   "AssignmentDate",
    //   "AcknowledgeDate"
    // ];
    const { onCloseForm } = this.props;
    const { policy } = this.state;
    this.setState({ loading: true });

    try {
      const result = await SharePointService.pnp_postByTitle(
        "Policies",
        policy
      );

      // await SharePointService.pnp_createList(
      //   policy.Title,
      //   `policy assigment list`
      // );
      // const createdFields = await SharePointService.pnp_createField(
      //   policy.Title,
      //   fields
      // );

      // console.log("createdFields", createdFields);

      this.setState({ loading: false });
      onCloseForm();
    } catch (error) {
      console.log(error);
      //toast.error("error");
      this.setState({ loading: false });
      onCloseForm();
      return;
    }
  };

  public _getCurrentItem = async () => {};
}
