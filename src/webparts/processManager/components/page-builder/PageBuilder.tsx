// import { DefaultButton, Stack, IStackTokens } from "office-ui-fabric-react";
// import { PageTemplate } from "./PageTemplate";

import * as React from "react";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";

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
  selectedPolicyPageId: string;
  isPageFormOpen: boolean;
  isPolicyFormOpen: boolean;
  isDeleteFormOpen: boolean;
}

export default class PageBuilder extends React.Component<
  {},
  IPageBuilderState
> {
  private _allItems: any[];
  private _selectionForPolicyPages: Selection;

  private _getSelectionForPolicyPages = () => {
    const selectionCount = this._selectionForPolicyPages.getSelectedCount();
    // if (selectionCount !== 0)
    //   return (this._selection.getSelection()[0] as ITrackingRecordsItem).id;

    switch (selectionCount) {
      case 0:
        return "";
      case 1:
        return this._selectionForPolicyPages.getSelection()[0].key.toString();
      default:
        return this._selectionForPolicyPages.getSelection()[0].key.toString();
    }
  };

  constructor(props: {}) {
    super(props);

    this._selectionForPolicyPages = new Selection({
      onSelectionChanged: async () => {
        this.setState({
          selectedPolicyPageId: this._getSelectionForPolicyPages()
        });

        //await this._getGroupMembers(Number(this._getSelectionForTemplates()));
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
      selectedPolicyPageId: "",
      policyPagesColumns: _policyPagesColumns,
      isPageFormOpen: false,
      isPolicyFormOpen: false,
      isDeleteFormOpen: false
    };
  }

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
      key: policyPage.Name.split(".")[0],
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
                text: "Delete"
                //onClick: this._onOpenDeleteForm
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
        key: "createTemplate",
        name: "Create Template",
        cacheKey: "myCacheKey", // changing this key will invalidate this items cache
        iconProps: {
          iconName: "PageHeaderEdit"
        },
        ariaLabel: "Create Template",
        onClick: this.onOpenPageForm
      }
      // {
      //   key: "newPolicy",
      //   name: "Create new policy",
      //   iconProps: {
      //     iconName: "EntitlementPolicy"
      //   },

      //   onClick: this.onOpenPolicyForm
      // }
    ];
  };

  private _getPolicy = async () => {
    const result = await SharePointService.pnp_getPolicy("Policies");
    const policies = result.value.map(p => {
      return {
        title: p.Title,
        peopleAssigned: p.PeopleAssigned,
        policyOwner: p.PolicyOwner,
        policyPagesTitle: p.PolicyPagesTitle
      };
    });
    console.log("resultP", result.value);

    // this.setState({ policies });
    return policies;
  };

  public async componentDidMount() {
    await this._getPolicyPages();
  }

  public render(): JSX.Element {
    const {
      policyPages,
      selectedPolicyPageId,
      policyPagesColumns,
      isPageFormOpen,
      isPolicyFormOpen,
      isDeleteFormOpen
    } = this.state;
    //selectionId && this._getGroupMembers(Number(selectionId));
    // background: "#f3f3f3"
    return (
      <div>
        {/* <DefaultButton
          onClick={() => this._getRelevantPolicy()}
          text="getPolicy"
        /> */}

        <Separator>
          <Text>Policy Page Builder</Text>
        </Separator>

        <CommandBar
          items={this._getMenuItems()}
          //overflowItems={this.getOverlflowItems()}
        />
        <Stack
          horizontal
          //horizontalAlign="space-evenly"
          wrap
          tokens={wrapStackTokens}
          style={{ marginBottom: 30, marginTop: 30 }}
        >
          {/* <div className={exampleChildClass}>{selectionDetails}</div> */}
          {/* <TextField
          className={exampleChildClass}
          label="Filter by name:"
          onChange={this._onFilter}
          styles={{ root: { maxWidth: "300px" } }}
        /> */}

          <Stack verticalAlign="center" tokens={stackTokens}>
            <Text>Templates</Text>

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
              onRenderRow={this.onRenderRow}
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
              subText: ""
            }}
            //   modalProps={{
            //     titleAriaId: this._labelId,
            //     dragOptions: this._dragOptions,
            //     isBlocking: false
            //     // styles: { main: { maxWidth: 750 } }
            //   }}
          >
            <div style={{ display: "flex", justifyContent: "center" }}>
              <DefaultButton
                style={{ backgroundColor: "#dc224d", color: "white" }}
                //disabled={loading}
                onClick={null}
                text="Delete"
              />
            </div>
          </Dialog>
        )}
      </div>
    );
  }

  public onRenderRow = (props: IDetailsRowProps): JSX.Element => {
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

  public onOpenPageForm = () => {
    this.setState({ isPageFormOpen: true });
  };

  public onClosePageForm = () => {
    this.setState({ isPageFormOpen: false });
  };
}
