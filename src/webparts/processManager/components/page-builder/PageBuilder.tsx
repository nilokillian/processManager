// import { DefaultButton, Stack, IStackTokens } from "office-ui-fabric-react";
// import { PageTemplate } from "./PageTemplate";

import * as React from "react";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { toast } from "react-toastify";
import {
  Text,
  DefaultButton,
  Separator,
  Stack,
  IStackTokens,
  MarqueeSelection,
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  TextField,
  CommandBar,
  IconButton,
  SelectionMode,
  Dialog,
  DialogType,
  IDetailsRowProps,
  DetailsRow,
  IDetailsRowStyles
} from "office-ui-fabric-react";
import SharePointService from "../../../../services/SharePoint/SharePointService";
import PageForm from "../page-builder/PageFrom";

const stackTokens: IStackTokens = { childrenGap: 12 };
const wrapStackTokens: IStackTokens = { childrenGap: 20 };

const exampleChildClass = mergeStyles({
  display: "block",
  marginBottom: "10px"
});

export interface IPolicy {
  Title: string;
  PolicyOwnerId: number;
  PolicyPagesTitle: string;
}

export interface IPageBuilderState {
  policyPagesColumns: IColumn[];
  policyPages: any[];
  policies: IPolicy[];
  newPage: {};
  selectedPolicyPageName: string;
  isPageFormOpen: boolean;
  isPolicyFormOpen: boolean;
  isDeleteFormOpen: boolean;
  loading: boolean;
}

export default class PageBuilder extends React.Component<
  {},
  IPageBuilderState
> {
  private _selectionForPolicyPages: Selection;
  private _getSelectionForPolicyPages = () => {
    const selectionCount = this._selectionForPolicyPages.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return "";
      case 1:
        return this._selectionForPolicyPages.getSelection()[0].key.toString();
      default:
        return this._selectionForPolicyPages.getSelection()[0].key.toString();
    }
  };

  private _isPolicyPagesActivated = (item): boolean => {
    return this.state.policies.some(
      policy => policy.PolicyPagesTitle === item.name.split(".")[0]
    );
  };

  private _getPolicyPages = async () => {
    const result = await SharePointService.getPolicyPages(
      "SitePages",
      "Templates"
    );
    const policies = await this._getPolicy();

    const getPolicyTitle = (pageTitle: string) => {
      const policy = policies.find(p => p.policyPagesTitle === pageTitle);

      return policy ? policy.title : "not assigned";
    };

    const policyPages = result.map(policyPage => ({
      key: policyPage.Name,
      title: policyPage.Title,
      name: policyPage.Name,
      activated: policies.some(
        policy => policy.templateTitle === policyPage.Name.split(".")[0]
      ),
      policy: getPolicyTitle(policyPage.Name.split(".")[0]),
      edit: (
        <IconButton
          menuProps={{
            shouldFocusOnMount: true,
            items: [
              {
                key: "delete",
                text: "Delete",
                onClick: this._onOpenDeleteForm
              }
            ]
          }}
        />
      )
    }));

    this.setState({ policyPages, policies });
  };

  private _getMenuItems = () => {
    return [
      {
        key: "createPolicyPage",
        name: "Create new policy page",
        cacheKey: "myCacheKey", // changing this key will invalidate this items cache
        iconProps: {
          iconName: "Add"
        },
        ariaLabel: "Create new policy page",
        onClick: this._onOpenPageForm
      }
    ];
  };

  private _onOpenPageForm = () => {
    this.setState({ isPageFormOpen: true });
  };

  private _getPolicy = async () => {
    const result = await SharePointService.pnp_getPolicies("Policies");
    const policies = result.value.map(p => {
      return {
        title: p.Title,
        peopleAssigned: p.PeopleAssigned,
        policyOwner: p.PolicyOwner,
        policyPagesTitle: p.PolicyPagesTitle
      };
    });
    return policies;
  };

  private _onRenderRow = (props: IDetailsRowProps): JSX.Element => {
    const customStyles: Partial<IDetailsRowStyles> = {};
    console.log(props.item.activated);
    if (props.item.activated) {
      customStyles.root = [
        "root",
        {
          backgroundColor: "#b4f1b4"
        }
      ];
    }

    return <DetailsRow {...props} styles={customStyles} />;
  };

  private _onOpenDeleteForm = () => {
    this.setState({ isDeleteFormOpen: true });
  };

  private _onCloseDeleteForm = () => {
    this.setState({ isDeleteFormOpen: false });
  };

  private _deletePolicyPage = async () => {
    const { selectedPolicyPageName } = this.state;

    try {
      await SharePointService.deletePolicyPage(
        "SitePages",
        "Templates",
        selectedPolicyPageName
      );
      this._updatePolicies();
      this._onCloseDeleteForm();
      await this._getPolicyPages();
      toast.success("deleted");
    } catch (error) {
      this._onCloseDeleteForm();
      toast.error("deleted");
      throw error;
    }
  };

  private _updatePolicies = () => {
    const { selectedPolicyPageName } = this.state;
    const policyPageTitle = selectedPolicyPageName.split(".")[0].trim();

    SharePointService.pnp_update_collection_filter(
      "Policies",
      "PolicyPagesTitle",
      policyPageTitle,
      "PolicyPagesTitle",
      ""
    );
  };

  constructor(props: {}) {
    super(props);

    this._selectionForPolicyPages = new Selection({
      onSelectionChanged: async () => {
        this.setState({
          selectedPolicyPageName: this._getSelectionForPolicyPages()
        });
      }
    });

    const _policyPagesColumns: IColumn[] = [
      {
        key: "title",
        name: "Title",
        fieldName: "title",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
        //onColumnClick: this._onColumnClick
      },

      {
        key: "edit",
        name: "",
        fieldName: "edit",
        minWidth: 50,
        maxWidth: 50,
        isResizable: false
      },
      {
        key: "name",
        name: "Name",
        fieldName: "name",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
        //onColumnClick: this._onColumnClick
      },

      {
        key: "activeted",
        name: "Activeted",
        fieldName: "activeted",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true,
        onRender: item => {
          return (
            <span>{this._isPolicyPagesActivated(item) ? "Yes" : "No"}</span>
          );
        }
        //onColumnClick: this._onColumnClick
      },

      {
        key: "policy",
        name: "Policy",
        fieldName: "policy",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
        //onColumnClick: this._onColumnClick
      }
    ];

    this.state = {
      newPage: {},
      policyPages: [],
      policies: [] as IPolicy[],
      selectedPolicyPageName: "",
      policyPagesColumns: _policyPagesColumns,
      isPageFormOpen: false,
      isPolicyFormOpen: false,
      isDeleteFormOpen: false,
      loading: false
    };
  }

  public render(): JSX.Element {
    const {
      policyPages,
      selectedPolicyPageName,
      policyPagesColumns,
      isPageFormOpen,
      isPolicyFormOpen,
      isDeleteFormOpen,
      loading
    } = this.state;
    //selectionId && this._getGroupMembers(Number(selectionId));
    // background: "#f3f3f3"
    return (
      <div>
        <Separator>
          <Text>Policy Page Builder</Text>
        </Separator>

        <CommandBar items={this._getMenuItems()} />
        <Stack
          horizontal
          //horizontalAlign="space-evenly"
          wrap
          tokens={wrapStackTokens}
          style={{ marginBottom: 30, marginTop: 30 }}
        >
          <Stack verticalAlign="center" tokens={stackTokens}>
            <DetailsList
              items={policyPages}
              columns={policyPagesColumns}
              setKey="setOfpolicyPages"
              layoutMode={DetailsListLayoutMode.justified}
              selection={this._selectionForPolicyPages}
              selectionMode={SelectionMode.single}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              checkButtonAriaLabel="Row checkbox"
              onRenderRow={this._onRenderRow}
              //onItemInvoked={this._onItemInvoked}
            />
          </Stack>
        </Stack>

        {isPageFormOpen && (
          <PageForm
            onCloseForm={this.onClosePageForm}
            isOpenForm={isPageFormOpen}
          />
        )}

        {isDeleteFormOpen && (
          <Dialog
            hidden={!isDeleteFormOpen}
            onDismiss={this._onCloseDeleteForm}
            maxWidth={670}
            dialogContentProps={{
              type: DialogType.close,
              title: "Are you sure ?",
              subText:
                "Performing this action you might affect policies that have this policy page assigned"
            }}
          >
            <div style={{ display: "flex", justifyContent: "center" }}>
              <DefaultButton
                style={{ backgroundColor: "#dc224d", color: "white" }}
                disabled={loading}
                onClick={this._deletePolicyPage}
                text="Delete"
              />
            </div>
          </Dialog>
        )}
      </div>
    );
  }

  public async componentDidMount() {
    await this._getPolicyPages();
  }

  public onClosePageForm = async () => {
    await this._getPolicyPages();
    this.setState({ isPageFormOpen: false });
  };
}
