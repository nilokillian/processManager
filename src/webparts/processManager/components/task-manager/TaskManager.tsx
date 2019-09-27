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
  IDetailsRowStyles,
  Toggle
} from "office-ui-fabric-react";
import SharePointService from "../../../../services/SharePoint/SharePointService";

const stackTokens: IStackTokens = { childrenGap: 12 };
const wrapStackTokens: IStackTokens = { childrenGap: 20 };

const exampleChildClass = mergeStyles({
  display: "block",
  marginBottom: "10px"
});

export interface IComletedTask {
  title: string;
  userId: string;
  email: string;
  acknowledgeDate: string;
  policy: string;
  status: string;
  userGroupTitle: string;
}

export interface IActiveTask {
  title: string;
  userId: string;
  email: string;
  assignmentDate: string;
  policy: string;
  userGroupTitle: string;
}

export interface ITaskManagerState {
  activeTaskColumns: IColumn[];
  completedTaskColumns: IColumn[];
  activeTasks: IActiveTask[];
  completedTasks: IComletedTask[];
  selectedActiveTaskId: string;
  selectedCompletedTaskId: string;
  isActiveTaskFormOpen: boolean;
  isCompletedTaskFormOpen: boolean;
  isDeleteFormOpen: boolean;
  isCompetedTasks: boolean;
}

export default class TaskManager extends React.Component<
  {},
  ITaskManagerState
> {
  private _allItems: any[];
  private _selectionForActiveTasks: Selection;
  private _selectionForCompletedTasks: Selection;

  private _getSelectionForActiveTasks = () => {
    const selectionCount = this._selectionForActiveTasks.getSelectedCount();
    // if (selectionCount !== 0)
    //   return (this._selection.getSelection()[0] as ITrackingRecordsItem).id;

    switch (selectionCount) {
      case 0:
        return "";
      case 1:
        return this._selectionForActiveTasks.getSelection()[0].key.toString();
      default:
        return this._selectionForActiveTasks.getSelection()[0].key.toString();
    }
  };

  private _getSelectionForCompletedTasks = () => {
    const selectionCount = this._selectionForCompletedTasks.getSelectedCount();
    // if (selectionCount !== 0)
    //   return (this._selection.getSelection()[0] as ITrackingRecordsItem).id;

    switch (selectionCount) {
      case 0:
        return "";
      case 1:
        return this._selectionForCompletedTasks
          .getSelection()[0]
          .key.toString();
      default:
        return this._selectionForCompletedTasks
          .getSelection()[0]
          .key.toString();
    }
  };

  constructor(props: {}) {
    super(props);

    this._selectionForActiveTasks = new Selection({
      onSelectionChanged: async () => {
        this.setState({
          selectedActiveTaskId: this._getSelectionForActiveTasks()
        });
      }
    });

    this._selectionForCompletedTasks = new Selection({
      onSelectionChanged: async () => {
        this.setState({
          selectedCompletedTaskId: this._getSelectionForCompletedTasks()
        });
      }
    });

    const _activeTaskColumns: IColumn[] = [
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
        key: "userId",
        name: "User ID",
        fieldName: "userId",
        minWidth: 70,
        maxWidth: 100,
        isResizable: true
        //onColumnClick: this._onColumnClick
      },

      {
        key: "email",
        name: "Email",
        fieldName: "email",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
        //onColumnClick: this._onColumnClick
      },

      {
        key: "assignmentDate",
        name: "Assignment Date",
        fieldName: "assignmentDate",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
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
      },
      {
        key: "userGroupTitle",
        name: "User Group",
        fieldName: "userGroupTitle",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
        //onColumnClick: this._onColumnClick
      }
    ];

    const _completedTaskColumns: IColumn[] = [
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
        key: "userId",
        name: "User ID",
        fieldName: "userId",
        minWidth: 70,
        maxWidth: 100,
        isResizable: true
        //onColumnClick: this._onColumnClick
      },

      {
        key: "email",
        name: "Email",
        fieldName: "email",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
        //onColumnClick: this._onColumnClick
      },

      {
        key: "acknowledgeDate",
        name: "Acknowledge Date",
        fieldName: "acknowledgeDate",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
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
      },
      {
        key: "userGroupTitle",
        name: "User Group",
        fieldName: "userGroupTitle",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
        //onColumnClick: this._onColumnClick
      },
      {
        key: "status",
        name: "Status",
        fieldName: "status",
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
        //onColumnClick: this._onColumnClick
      }
    ];

    this.state = {
      activeTaskColumns: _activeTaskColumns,
      completedTaskColumns: _completedTaskColumns,
      activeTasks: [],
      completedTasks: [],
      selectedActiveTaskId: "",
      selectedCompletedTaskId: "",
      isActiveTaskFormOpen: false,
      isCompletedTaskFormOpen: false,
      isDeleteFormOpen: false,
      isCompetedTasks: false
    };
  }

  private _getComletedTasks = async () => {
    const result = await SharePointService.pnp_getListItems("CompletedTasks");

    const completedTasks = result.value.map(
      task =>
        ({
          key: task.ID,
          title: task.Title,
          userId: task.UserId.toString(),
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
          ),
          email: task.Email,
          acknowledgeDate: task.AcknowledgeDate,
          policy: task.Policy,
          userGroupTitle: task.UserGroupTitle,
          status: task.Status
        } as IComletedTask)
    );

    this.setState({ completedTasks });
  };

  private _getActiveTasks = async () => {
    const result = await SharePointService.pnp_getListItems("ActiveTasks");

    const activeTasks = result.value.map(
      task =>
        ({
          key: task.ID,
          title: task.Title,
          userId: task.UserId.toString(),
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
          ),
          email: task.Email,
          assignmentDate: task.AssignmentDate,
          policy: task.Policy,
          userGroupTitle: task.UserGroupTitle
        } as IActiveTask)
    );

    this.setState({ activeTasks });
  };

  public async componentDidMount() {
    await this._getActiveTasks();
    await this._getComletedTasks();
  }

  public render(): JSX.Element {
    const {
      activeTaskColumns,
      completedTaskColumns,
      activeTasks,
      completedTasks,
      selectedActiveTaskId,
      selectedCompletedTaskId,
      isActiveTaskFormOpen,
      isCompletedTaskFormOpen,
      isDeleteFormOpen,
      isCompetedTasks
    } = this.state;

    return (
      <div>
        <Separator>
          <Text>Task manager</Text>
        </Separator>
        <Toggle
          label="Tasks"
          defaultChecked
          onText="on"
          offText="off"
          onChange={this._onCompetedTasks}
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
            <Text>{isCompetedTasks ? "Completed tasks" : "Active Tasks"}</Text>

            <DetailsList
              items={isCompetedTasks ? completedTasks : activeTasks}
              columns={
                isCompetedTasks ? completedTaskColumns : activeTaskColumns
              }
              setKey="setOfpolicyPages"
              layoutMode={DetailsListLayoutMode.justified}
              selection={
                isCompetedTasks
                  ? this._selectionForCompletedTasks
                  : this._selectionForActiveTasks
              }
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

        {/* {isPageFormOpen && (
          <PageForm
            onCloseForm={this.onClosePageForm}
            isOpenForm={isPageFormOpen}
          />
        )} */}

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

  private _onCompetedTasks = (
    ev: React.MouseEvent<HTMLElement>,
    checked: boolean
  ) => {
    this.setState({ isCompetedTasks: checked });
  };

  private _onOpenDeleteForm = () => {
    this.setState({ isDeleteFormOpen: true });
  };

  private _onCloseDeleteForm = () => {
    this.setState({ isDeleteFormOpen: false });
  };

  public onOpenPageForm = () => {
    this.setState({ isActiveTaskFormOpen: true });
  };

  public onClosePageForm = () => {
    this.setState({ isActiveTaskFormOpen: false });
  };
}
