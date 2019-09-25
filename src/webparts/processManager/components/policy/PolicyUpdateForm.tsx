import * as React from "react";
//import { toast } from "react-toastify";
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
  TextField
} from "office-ui-fabric-react";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 5
};

export interface IPolicyUpdateFormProps {
  currentPolicy: { id: number; name: string };
  isOpenForm: boolean;
  onCloseForm(): void;
}

export interface IPolicyUpdateFormState {
  users: any[];
  group: { id: number };
  groupUsers: groupUser[];
  errors: object;
  loading: boolean;
}

export interface groupUser {
  Id: string;
  Title: string;
  Email: string;
  Status: string;
  AssignmentDate: string;
  AcknowledgeDate?: string;
  UserId: string;
}

export default class PolicyUpdateForm extends React.Component<
  IPolicyUpdateFormProps,
  IPolicyUpdateFormState
> {
  private _dragOptions = {
    moveMenuItemText: "Move",
    closeMenuItemText: "Close",
    menu: ContextualMenu
  };

  constructor(props: IPolicyUpdateFormProps) {
    super(props);

    this.state = {
      users: [],
      group: { id: null },
      groupUsers: [],
      errors: {},
      loading: false
    };
  }

  public async componentDidMount() {}

  public render(): JSX.Element {
    const { isOpenForm, onCloseForm, currentPolicy } = this.props;
    const { loading, group } = this.state;

    return (
      <div>
        <Panel
          isOpen={isOpenForm}
          type={PanelType.custom}
          customWidth="420px"
          onDismiss={onCloseForm}
          headerText={currentPolicy.name}
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
              value={title}
              onChange={this._onChangeTextInput}
              //styles={ComponentStyles.textInputStyle()}
              disabled={loading}
              required={true}
            />

            <TextField
              id="Title"
              label="Title"
              value={title}
              onChange={this._onChangeTextInput}
              //styles={ComponentStyles.textInputStyle()}
              disabled={loading}
              required={true}
            />
            <DefaultButton text="tasks" onClick={this._createTask} />
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

  private _getPeoplePickerGroup = async (items: any[]) => {
    const { group } = this.state;
    group.id = items[0].id;
    const groupUsers = await SharePointService.pnp_getGroupMembers(group.id);
    this.setState({ group, groupUsers });
  };

  private _createTask = async () => {
    const { groupUsers } = this.state;
    const { currentPolicy } = this.props;
    const qObjGroupUsers = groupUsers.map(u => {
      return {
        UserId: u.UserId,
        Title: u.Title,
        Email: u.Email,
        Status: "Assigned",
        AssignmentDate: moment(new Date())
          .format("DD/MM/YYYY")
          .toString()
      };
    });
    //console.log("ObjGroupUsers ", groupUsers);
    const tasks = await SharePointService.pnp_postByTitle_multiple(
      currentPolicy.name,
      qObjGroupUsers
    );
    console.log("tasks", tasks);
  };

  public submitForm = async () => {
    this.setState({ loading: true });
    const { group } = this.state;
    const { onCloseForm, currentPolicy } = this.props;
    try {
      const updatedPolicy = await SharePointService.pnp_updateByTitle(
        "Policies",
        currentPolicy.id,
        {
          PeopleAssignedId: { results: [group.id] }
        }
      );

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
