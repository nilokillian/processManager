import * as React from "react";
import { toast } from "react-toastify";
import {
  Text,
  DefaultButton,
  Separator,
  Stack,
  IStackTokens,
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  IconButton,
  SelectionMode,
  Dialog,
  DialogType,
  Toggle
} from "office-ui-fabric-react";
import SharePointService from "../../../../services/SharePoint/SharePointService";

const wrapStackTokens: IStackTokens = { childrenGap: 20 };

export interface IComletedTask {
  title: string;
  userId: string;
  email: string;
  acknowledgeDate: string;
  policy: string;
  status: string;
  userGroupTitle: string;
}

export interface ITask {
  title: string;
  userId: string;
  email: string;
  assignmentDate: string;
  policy: string;
  userGroupTitle: string;
  status: string;
}

export interface ITaskManagerState {
  taskColumns: IColumn[];
  tasks: ITask[];
  selectedTaskId: number;
  isTaskFormOpen: boolean;
  isDeleteFormOpen: boolean;
  isAcknowledgedTasks: boolean;
  isCanceledTasks: boolean;
  loading: boolean;
}

export default class TaskManager extends React.Component<
  {},
  ITaskManagerState
> {
  private _selectionForTasks: Selection;

  private _getSelectionForTasks = () => {
    const selectionCount = this._selectionForTasks.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return null;
      case 1:
        return this._selectionForTasks.getSelection()[0].key;
      default:
        return this._selectionForTasks.getSelection()[0].key;
    }
  };

  constructor(props: {}) {
    super(props);

    this._selectionForTasks = new Selection({
      onSelectionChanged: async () => {
        this.setState({
          selectedTaskId: Number(this._getSelectionForTasks())
        });
      }
    });

    const _taskColumns: IColumn[] = [
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
        key: "acknowledgeDate",
        name: "Acknowledge Date",
        fieldName: "acknowledgeDate",
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

    this.state = {
      taskColumns: _taskColumns,
      tasks: [],
      selectedTaskId: null,
      isTaskFormOpen: false,
      isDeleteFormOpen: false,
      isAcknowledgedTasks: false,
      isCanceledTasks: false,
      loading: false
    };
  }

  private _deleteTask = async () => {
    const { selectedTaskId } = this.state;
    this.setState({ loading: true });
    try {
      await SharePointService.pnp_delete("UserTasks", selectedTaskId);

      toast.success("deleted");
      this._getTasks();

      this._onCloseDeleteForm();
    } catch (error) {
      toast.error("error");
      this._onCloseDeleteForm();
      throw error;
    }
    this.setState({ loading: false });
  };

  private _getTasks = async () => {
    const result = await SharePointService.pnp_getListItems("UserTasks");

    const tasks = result.value.map(task => ({
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
                text: "Delete",
                onClick: this._onOpenDeleteForm
              }
            ]
          }}
        />
      ),
      email: task.Email,
      assignmentDate: task.AssignmentDate,
      acknowledgeDate: task.AcknowledgeDate,
      policy: task.Policy,
      userGroupTitle: task.UserGroupTitle,
      status: task.Status
    }));

    this.setState({ tasks });
  };

  public async componentDidMount() {
    await this._getTasks();
  }

  public render(): JSX.Element {
    const {
      taskColumns,
      loading,
      isDeleteFormOpen,
      isCanceledTasks,
      isAcknowledgedTasks
    } = this.state;
    const filteredTasks = this._filterTasks();
    return (
      <div>
        <Separator>
          <Text>Task manager</Text>
        </Separator>

        <Stack
          horizontal
          horizontalAlign="space-evenly"
          tokens={wrapStackTokens}
          style={{ marginBottom: 30, marginTop: 30 }}
        >
          <Toggle
            label="acknowledged tasks"
            inlineLabel
            checked={isAcknowledgedTasks}
            onText=""
            offText=""
            onChange={this._onAcknowledgeTasks}
          />
          <Toggle
            label="canceled tasks"
            inlineLabel
            onText=""
            offText=""
            checked={isCanceledTasks}
            onChange={this._onCanceledTasks}
          />
        </Stack>

        <Stack
          horizontal
          wrap
          tokens={wrapStackTokens}
          style={{ marginBottom: 30, marginTop: 30 }}
        >
          <DetailsList
            items={filteredTasks}
            columns={taskColumns}
            setKey="setOfpolicyPages"
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selectionForTasks}
            selectionMode={SelectionMode.single}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
          />
        </Stack>
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
          >
            <div style={{ display: "flex", justifyContent: "center" }}>
              <DefaultButton
                style={{ backgroundColor: "#dc224d", color: "white" }}
                disabled={loading}
                onClick={this._deleteTask}
                text="Delete"
              />
            </div>
          </Dialog>
        )}
      </div>
    );
  }
  private _onAcknowledgeTasks = (
    ev: React.MouseEvent<HTMLElement>,
    checked: boolean
  ) => {
    this.setState({ isAcknowledgedTasks: checked, isCanceledTasks: false });
  };

  private _onCanceledTasks = (
    ev: React.MouseEvent<HTMLElement>,
    checked: boolean
  ) => {
    this.setState({ isCanceledTasks: checked, isAcknowledgedTasks: false });
  };

  private _filterTasks = () => {
    const { isCanceledTasks, isAcknowledgedTasks, tasks } = this.state;

    if (isCanceledTasks) {
      return tasks.filter(task => task.status === "Canceled");
    } else if (isAcknowledgedTasks) {
      return tasks.filter(task => task.status === "Acknowledged");
    } else {
      return tasks;
    }
  };

  private _onOpenDeleteForm = () => {
    this.setState({ isDeleteFormOpen: true });
  };

  private _onCloseDeleteForm = () => {
    this.setState({ isDeleteFormOpen: false });
  };

  public onOpenPageForm = () => {
    this.setState({ isTaskFormOpen: true });
  };

  public onClosePageForm = () => {
    this.setState({ isTaskFormOpen: false });
  };
}
