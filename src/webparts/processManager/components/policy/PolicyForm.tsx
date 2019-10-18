import * as React from "react";
import { toast } from "react-toastify";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import SharePointService from "../../../../services/SharePoint/SharePointService";
import {
  DefaultButton,
  Dialog,
  PrimaryButton,
  DialogType,
  Panel,
  PanelType,
  Stack,
  IStackTokens,
  TextField,
  Dropdown,
  IDropdownOption
} from "office-ui-fabric-react";

const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 5
};

export interface IPolicyFrom {
  Title: string;
  Id: number;
  PolicyOwner: { Id: number; EMail: string; Title: string };
  PolicyOwnerId?: number;
  PolicyPagesTitle: string;
  GroupAssigned: [
    {
      Id: number;
      Title: string;
    }
  ];
}

export interface IPolicyFormProps {
  selectedPolicyId: number;
  policyPages: any[];

  isOpenForm: boolean;
  onCloseForm(): void;
  resetSelectedPolicyId(): void;
  updateComponent(): void;
}

export interface IPolicyFormState {
  policy: IPolicyFrom;
  selectedPolicyId: number;
  errors: object;
  loading: boolean;
  isDeletePolicyFormOpen: boolean;
}

export default class PolicyForm extends React.Component<
  IPolicyFormProps,
  IPolicyFormState
> {
  constructor(props: IPolicyFormProps) {
    super(props);

    this.state = {
      policy: {} as IPolicyFrom,
      selectedPolicyId: null,
      errors: {},
      loading: false,
      isDeletePolicyFormOpen: false
    };
  }

  private _onPolicyPageChange = (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): void => {
    const { policy } = this.state;
    policy.PolicyPagesTitle = item.text;
    this.setState({ policy });
  };

  private _getPolicyPageOptions = (): IDropdownOption[] => {
    const { policyPages } = this.props;

    return policyPages.map(p => {
      return {
        key: p.name.split(".")[0],
        text: p.name.split(".")[0]
      } as IDropdownOption;
    });
  };

  private _onRenderFooterContent = () => {
    const { loading, selectedPolicyId } = this.state;
    return (
      <div>
        <PrimaryButton
          onClick={this._submitForm}
          text={!selectedPolicyId ? "Save" : "Update"}
          disabled={loading || this._validation()}
          style={{ marginRight: "8px" }}
        />
        {!!selectedPolicyId && (
          <PrimaryButton
            onClick={this._openDeletePolicyForm}
            text={"Delete"}
            disabled={loading}
          />
        )}
      </div>
    );
  };

  private _onChangeTextInput = (
    e: React.FormEvent<HTMLInputElement>,
    newValue?: string
  ) => {
    const { policy } = this.state;
    policy.Title = newValue;

    this.setState({ policy });
  };

  private _submitForm = async () => {
    this.setState({ loading: true });
    const { onCloseForm, updateComponent } = this.props;
    const { policy, selectedPolicyId } = this.state;

    delete policy.GroupAssigned;
    delete policy.PolicyOwner;
    delete policy.Id;

    try {
      if (selectedPolicyId) {
        await SharePointService.pnp_updateByTitle(
          "Policies",
          selectedPolicyId,
          policy
        );
        toast.success("updated");
      } else {
        const result = await SharePointService.pnp_postByTitle(
          "Policies",
          policy
        );
        toast.success("created");
      }

      updateComponent();
      this.setState({ loading: false });
      onCloseForm();
    } catch (error) {
      toast.error("error");
      this.setState({ loading: false });
      onCloseForm();
      throw error;
    }
  };

  private _deletePolicy = async () => {
    this.setState({ loading: true });
    const { selectedPolicyId, onCloseForm, updateComponent } = this.props;

    try {
      await SharePointService.pnp_delete("Policies", selectedPolicyId);
      updateComponent();
      onCloseForm();
      toast.success("deleted");
      this.setState({ loading: false });
    } catch (error) {
      onCloseForm();
      toast.error("error");
      this.setState({ loading: false });
      throw error;
    }
  };

  private _validation = () => {
    const { policy, selectedPolicyId } = this.state;

    if (!selectedPolicyId) {
      return policy.PolicyOwnerId && policy.Title && policy.PolicyPagesTitle
        ? false
        : true;
    } else {
      return false;
    }
  };

  private _closeForm = () => {
    const { resetSelectedPolicyId, onCloseForm } = this.props;
    onCloseForm();
    resetSelectedPolicyId();
  };

  private _closeDeletePolicyForm = () => {
    this.setState({ isDeletePolicyFormOpen: false });
  };

  private _openDeletePolicyForm = () => {
    this.setState({ isDeletePolicyFormOpen: true });
  };

  private _getPeoplePickerItems = async (items: any[]) => {
    const { policy } = this.state;

    if (items[0]) {
      const user = await SharePointService.pnp_getUserId(
        items[0].secondaryText
      );
      policy.PolicyOwnerId = user.Id;
      this.setState({ policy });
    } else {
      policy.PolicyOwnerId = null;

      this.setState({ policy });
    }
  };

  private _concatAssignedGroups = () => {
    const { policy } = this.state;
    let groups = "";
    if (policy.GroupAssigned)
      policy.GroupAssigned.map(g => (groups += g.Title + ";"));
    return groups;
  };

  public render(): JSX.Element {
    const { isOpenForm } = this.props;
    const {
      loading,
      policy,
      isDeletePolicyFormOpen,
      selectedPolicyId
    } = this.state;

    return (
      <Panel
        isOpen={isOpenForm}
        type={PanelType.custom}
        customWidth="420px"
        onDismiss={this._closeForm}
        headerText={!selectedPolicyId ? "Create policy" : policy.Title}
        closeButtonAriaLabel="Close"
        //onRenderHeader={this._onRenderHeaderContent}
        onRenderFooterContent={this._onRenderFooterContent}
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
            disabled={loading}
            required={true}
          />

          <Dropdown
            id="policyPage"
            label="Policy page"
            placeholder="Select policy page"
            options={this._getPolicyPageOptions()}
            selectedKey={
              policy.PolicyPagesTitle ? policy.PolicyPagesTitle : undefined
            }
            disabled={loading}
            required={true}
            onChange={this._onPolicyPageChange}
          />

          {!!selectedPolicyId && (
            <TextField
              id="groupAssigned"
              label="Group Assigned"
              value={this._concatAssignedGroups()}
              disabled={true}
            />
          )}
          <PeoplePicker
            context={SharePointService.context}
            titleText="Policy owner"
            personSelectionLimit={1}
            groupName={""} // Leave this blank in case you want to filter from all users
            isRequired={true}
            selectedItems={this._getPeoplePickerItems}
            defaultSelectedUsers={
              policy.PolicyOwner ? [policy.PolicyOwner.EMail] : []
            }
            showHiddenInUI={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            disabled={loading || !!selectedPolicyId}
          />
        </Stack>
        <Dialog
          hidden={!isDeletePolicyFormOpen}
          onDismiss={this._closeDeletePolicyForm}
          maxWidth={670}
          dialogContentProps={{
            type: DialogType.close,
            title: "Are you sure ?"
            //subText: "some text"
          }}
        >
          <div style={{ display: "flex", justifyContent: "center" }}>
            <DefaultButton
              style={{ backgroundColor: "#dc224d", color: "white" }}
              disabled={loading}
              onClick={this._deletePolicy}
              text="Delete"
            />
          </div>
        </Dialog>
      </Panel>
    );
  }

  public async componentDidMount() {
    const { selectedPolicyId } = this.props;
    if (selectedPolicyId) {
      this.setState({ selectedPolicyId });
      await this.getPolicy();
    }
  }

  public getPolicy = async () => {
    const { selectedPolicyId } = this.props;
    const fields = [
      { key: "ID", fieldType: "Counter" },
      { key: "Title", fieldType: "Text" },
      { key: "PolicyPagesTitle", fieldType: "Text" },
      { key: "PolicyOwner", fieldType: "User" },
      { key: "GroupAssigned", fieldType: "LookupMulti", lookupField: "Title" },
      { key: "GroupAssigned", fieldType: "LookupMulti", lookupField: "ID" }
    ];
    const expend = SharePointService.createExpendedFields(fields);
    const query = SharePointService.createQueriedFields(fields);
    const result = await SharePointService.pnp_getItem(
      "Policies",
      selectedPolicyId,
      expend,
      query
    );

    const policy = {
      Title: result.Title,
      Id: result.ID,
      PolicyOwner: {
        Id: result.PolicyOwner.ID,
        EMail: result.PolicyOwner.EMail,
        Title: result.PolicyOwner.Title
      },
      PolicyPagesTitle: result.PolicyPagesTitle,
      GroupAssigned:
        result.GroupAssigned &&
        result.GroupAssigned.map(g => {
          return {
            Id: g.ID,
            Title: g.Title
          };
        })
    };

    this.setState({ policy });
  };
}
