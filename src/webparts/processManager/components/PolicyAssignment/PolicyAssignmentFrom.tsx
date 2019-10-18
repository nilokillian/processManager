import * as React from "react";
import { toast } from "react-toastify";
import * as moment from "moment";
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
  Toggle
} from "office-ui-fabric-react";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IGroupsAssigned, IPeopeleAssigned } from "./PolicyAssignment";

const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 5
};

export interface IPolicyAssignmentFormProps {
  currentPolicy: {
    id: number;
    name: string;
    groupsAssigned: IGroupsAssigned[];
    peopeleAssigned: IPeopeleAssigned[];
  };
  isOpenForm: boolean;
  onCloseForm(): void;
}

export interface IGrpoup {
  id: number;
  name: string;
}

export interface IPolicyAssignmentFormState {
  users: any[];
  groups: IGroupsAssigned[];
  usersInsideGroups: groupUser[];
  activeTasks: ITask[];
  errors: object;
  assignGroup: boolean;
  assignUser: boolean;
  loading: boolean;
}

export interface ITask {
  UserId: string;
  Title: string;
  Email: string;
  AssignmentDate: string;
  Policy: string;
  UserGroupTitle: string;
}

export interface groupUser {
  title: string;
  email: string;
  userId: number;
  groupName: string;
}

export default class PolicyAssignmentForm extends React.Component<
  IPolicyAssignmentFormProps,
  IPolicyAssignmentFormState
> {
  private _getActiveTasks = async () => {
    const { currentPolicy } = this.props;

    //currentActiveTasks

    const result = await SharePointService.pnp_activeTasks(
      "ActiveTasks",
      currentPolicy.name
    );
    const activeTasks = result.value.map(task => ({
      Id: task.ID,
      UserId: task.UserId,
      Title: task.Title,
      Email: task.Email,
      AssignmentDate: task.AssignmentDate,
      Policy: task.Policy,
      UserGroupTitle: task.UserGroupTitle
    }));
    //console.log("currentActiveTasks", result.value);
    this.setState({ activeTasks });
  };

  private _updateTasks = async () => {
    const groups: any[] = this.state.groups;
    const activeTasks: any[] = this.state.activeTasks;
    const { groupsAssigned, peopeleAssigned } = this.props.currentPolicy;
    const newGroups = [];

    //check if there are new groups in peoplePicker
    groups.forEach(group => {
      const isExistingGroup = groupsAssigned.some(
        groupAssigned => groupAssigned.name === group.name
      );
      if (!isExistingGroup) {
        newGroups.push(group);
      }
    });

    //check there are newly added groups, get users from them and create tasks
    if (newGroups.length !== 0) {
      const usersInsideNewGroups = await SharePointService.pnp_getGroupMembers(
        newGroups
      );
      await this._createTask(usersInsideNewGroups);
    } else {
      // no new groups added but the exisiting ones might have more or less users
      const tasksToRemove = [];
      const tasksToAdd = [];
      const usersInsideExistingGroups = await SharePointService.pnp_getGroupMembers(
        groups
      );

      //checking if less users. If so => remove the corresponding tasks
      activeTasks.forEach(task => {
        const usrExists = usersInsideExistingGroups.some(
          usr => usr.userId === Number(task.UserId)
        );
        if (!usrExists) tasksToRemove.push(task);
      });

      //cheking if more users
      console.log("usersInsideExistingGroups", usersInsideExistingGroups);
      usersInsideExistingGroups.forEach(user => {
        user.userId;
        user.groupName;
        const isTask = activeTasks.some(
          task => Number(task.UserId) === user.userId
        );

        if (!isTask) {
          tasksToAdd.push(user);
        }
      });

      await this._createTask(tasksToAdd);

      console.log("tasksToAdd", tasksToAdd);
      console.log("tasksToRemove", tasksToRemove);
    }
  };

  private _submitForm = async () => {
    this.setState({ loading: true });
    const { groups, usersInsideGroups } = this.state;
    const { onCloseForm, currentPolicy } = this.props;
    try {
      if (groups.length !== 0) {
        const updatedPolicy = await SharePointService.pnp_updateByTitle(
          "Policies",
          currentPolicy.id,
          {
            GroupAssignedId: { results: groups.map(g => g.id) }
          }
        );
        await this._createTask(usersInsideGroups);
      } else {
        this._updateTasks();
      }

      this.setState({ loading: false });
      onCloseForm();
    } catch (error) {
      toast.error("error");
      this.setState({ loading: false });
      onCloseForm();
      throw error;
    }
  };

  private _onAssignGroupToggle = (
    ev: React.MouseEvent<HTMLElement>,
    checked: boolean
  ) => {
    const assignGroup = checked;

    this.setState({ assignGroup, assignUser: false });
  };

  private _onAssignUserToggle = (
    ev: React.MouseEvent<HTMLElement>,
    checked: boolean
  ) => {
    const assignUser = checked;

    this.setState({ assignGroup: false, assignUser });
  };

  private _onRenderFooterContent = () => {
    const { loading } = this.state;
    const { groupsAssigned, peopeleAssigned } = this.props.currentPolicy;
    return (
      <div>
        <PrimaryButton
          onClick={this._submitForm}
          text={groupsAssigned || peopeleAssigned ? "Update" : "Save"}
          disabled={loading}
        />
        <PrimaryButton
          onClick={this._updateTasks}
          text="test button"
          disabled={loading}
        />

        <PrimaryButton
          onClick={this._removeTask}
          text="remove tasks button"
          disabled={loading}
        />
      </div>
    );
  };

  private _getPeoplePickerGroup = async (items: any[]) => {
    //console.log("items", items);

    const groups = items.map(
      item => ({ id: item.id, name: item.text } as IGroupsAssigned)
    );

    const usersInsideGroups = await SharePointService.pnp_getGroupMembers(
      groups
    );
    this.setState({ groups, usersInsideGroups });
  };

  private _createTask = async usersInsideGroups => {
    //const { usersInsideGroups } = this.state;
    const { currentPolicy } = this.props;

    const qObjGroupUsers = usersInsideGroups.map(u => {
      return {
        UserId: u.userId.toString(),
        Title: u.title,
        Email: u.email,
        AssignmentDate: moment(new Date())
          .format("DD/MM/YYYY")
          .toString(),
        Policy: currentPolicy.name,
        UserGroupTitle: u.groupName
      };
    });
    //console.log("ObjGroupUsers ", groupUsers);
    const tasks = await SharePointService.pnp_postByTitle_multiple(
      "ActiveTasks",
      qObjGroupUsers
    );
    // console.log("tasks", tasks);
  };

  private _removeTask = async () => {
    await SharePointService.pnp_delete_multiple("ActiveTasks", [8, 9]);

    SP.Client;
  };

  constructor(props: IPolicyAssignmentFormProps) {
    super(props);

    this.state = {
      users: [],
      groups: [{ id: null, name: "" }],
      usersInsideGroups: [],
      activeTasks: [],
      errors: {},
      loading: false,
      assignGroup: true,
      assignUser: false
    };
  }

  public render(): JSX.Element {
    const { isOpenForm, onCloseForm, currentPolicy } = this.props;
    const { loading, groups, assignGroup, assignUser } = this.state;

    return (
      <Panel
        isOpen={isOpenForm}
        type={PanelType.custom}
        customWidth="420px"
        onDismiss={onCloseForm}
        headerText="Assigne people"
        closeButtonAriaLabel="Close"
        //onRenderHeader={this._onRenderHeaderContent}
        onRenderFooterContent={this._onRenderFooterContent}
      >
        <Stack
          //  styles={stackContainerStyles}
          tokens={itemAlignmentsStackTokens}
        >
          <PeoplePicker
            context={SharePointService.context}
            titleText="Assign Groups"
            personSelectionLimit={3}
            groupName={""} // Leave this blank in case you want to filter from all users
            isRequired={true}
            selectedItems={this._getPeoplePickerGroup}
            defaultSelectedUsers={
              currentPolicy.groupsAssigned &&
              currentPolicy.groupsAssigned.map(g => g.name)
            }
            showHiddenInUI={false}
            principalTypes={[PrincipalType.SharePointGroup]}
            resolveDelay={1000}
            disabled={!assignGroup}
          />
          <Toggle
            //label="Enabled and checked"
            checked={assignGroup}
            onText="On"
            offText="Off"
            onChange={this._onAssignGroupToggle}
          />

          <PeoplePicker
            context={SharePointService.context}
            titleText="Assign Individuals"
            personSelectionLimit={1}
            groupName={""} // Leave this blank in case you want to filter from all users
            isRequired={true}
            selectedItems={null}
            // defaultSelectedUsers={[Email]}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            disabled={!assignUser}
          />
          <Toggle
            //label="Enabled and checked"
            checked={assignUser}
            onText="On"
            offText="Off"
            onChange={this._onAssignUserToggle}
          />
        </Stack>
      </Panel>
    );
  }

  public async componentDidMount() {
    const { groupsAssigned, peopeleAssigned } = this.props.currentPolicy;

    if (groupsAssigned || peopeleAssigned) {
      await this._getActiveTasks();
    }

    if (groupsAssigned) {
      const usersInsideGroups = await SharePointService.pnp_getGroupMembers(
        groupsAssigned
      );
      this.setState({ usersInsideGroups, groups: groupsAssigned });
    }
  }
}
