import * as React from "react";
import { toast } from "react-toastify";
import {
  Stack,
  Separator,
  Text,
  DetailsList,
  DefaultButton,
  IColumn,
  Selection,
  SelectionMode,
  DetailsListLayoutMode,
  IconButton,
  IStackTokens,
  CommandBar,
  Dialog,
  DialogType,
  IStackStyles
} from "office-ui-fabric-react";
import GroupForm from "./NewGroupFrom";
import UserForm from "./AddUserFrom";
import SharePointService from "../../../../services/SharePoint/SharePointService";

const stackTokens: IStackTokens = { childrenGap: 12 };
const wrapStackTokens: IStackTokens = { childrenGap: 20 };
const stackStyles: IStackStyles = {
  root: {
    width: 300
  }
};
// export interface IGroupManagerProps {
//   selectedGroups: any[];
// }

export interface IGroupManagerState {
  groupColumns: IColumn[];
  memberColumns: IColumn[];
  groups: any[];
  members: any[];
  selectedGroupId: number;
  selectedUserId: number;
  isFormGroupOpen: boolean;
  isFormUserOpen: boolean;
  isDeleteFormOpen: boolean;
  loading: boolean;
}

export default class GroupManager extends React.Component<
  {},
  IGroupManagerState
> {
  private _allItems: any[];
  private _selectionForGroups: Selection;
  private _selectionForUsers: Selection;

  private _getSelectionForGroups = () => {
    const selectionCount = this._selectionForGroups.getSelectedCount();
    // if (selectionCount !== 0)
    //   return (this._selection.getSelection()[0] as ITrackingRecordsItem).id;

    switch (selectionCount) {
      case 0:
        return "";
      case 1:
        return this._selectionForGroups.getSelection()[0].key;
      default:
        return this._selectionForGroups.getSelection()[0].key;
    }
  };

  private _getSelectionForUsers = () => {
    const selectionCount = this._selectionForUsers.getSelectedCount();
    // if (selectionCount !== 0)
    //   return (this._selection.getSelection()[0] as ITrackingRecordsItem).id;

    switch (selectionCount) {
      case 0:
        return "";
      case 1:
        return this._selectionForUsers.getSelection()[0].key;
      default:
        return this._selectionForUsers.getSelection()[0].key;
    }
  };

  constructor(props: {}) {
    super(props);

    this._selectionForGroups = new Selection({
      onSelectionChanged: async () => {
        this.setState({
          selectedGroupId: Number(this._getSelectionForGroups())
        });

        await this._getGroupMembers(Number(this._getSelectionForGroups()));
      }
    });

    this._selectionForUsers = new Selection({
      onSelectionChanged: () => {
        this.setState({ selectedUserId: Number(this._getSelectionForUsers()) });
      }
    });

    const _groupColumns: IColumn[] = [
      {
        key: "groupName",
        name: "Group name",
        fieldName: "groupName",
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
      }
    ];

    const _memberColumns: IColumn[] = [
      {
        key: "displayName",
        name: "Display Name",
        fieldName: "displayName",
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
      }
    ];

    this.state = {
      groups: [],
      members: [],
      selectedGroupId: null,
      selectedUserId: null,
      groupColumns: _groupColumns,
      memberColumns: _memberColumns,
      isFormGroupOpen: false,
      isFormUserOpen: false,
      isDeleteFormOpen: false,
      loading: false
    };
  }

  public getGroups = async () => {
    const result = await SharePointService.pnp_getGroups();
    const groups = result.map(group => ({
      key: group.Id,
      groupName: group.Title,
      ownerTitle: group.OwnerTitle,
      principalType: group.principalType,
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

    this.setState({ groups });
  };

  private _getGroupMembers = async (groupId: number) => {
    const result = await SharePointService.pnp_getGroupMembers([
      { id: groupId, groupName: "" }
    ]);

    const members = result.map(member => ({
      key: member.userId,
      displayName: member.title,
      email: member.email,
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

    this.setState({ members });
  };

  private _deleteGroupMember = async () => {
    const { selectedGroupId, selectedUserId } = this.state;
    this.setState({ loading: true });
    try {
      await SharePointService.pnp_deleteGroupMember(
        selectedGroupId,
        selectedUserId
      );
      await this._getGroupMembers(Number(this._getSelectionForGroups()));
      toast.success("deleted");
      this._onCloseDeleteForm();
    } catch (error) {
      toast.error("error");
      this._onCloseDeleteForm();
      throw error;
    }
    this.setState({ loading: false });
  };

  private _deleteGroup = async () => {
    const { selectedGroupId } = this.state;
    this.setState({ loading: true });
    try {
      await SharePointService.pnp_deleteGroup(Number(selectedGroupId));
      this.setState({ selectedGroupId: null });
      await this.getGroups();
      toast.success("deleted");
      this._onCloseDeleteForm();
    } catch (error) {
      toast.error("error");
      this._onCloseDeleteForm();
      throw error;
    }
    this.setState({ loading: false });
  };

  private _getMenuItems = () => {
    const { selectedGroupId } = this.state;
    return [
      {
        key: "newItem",
        name: "New Group",
        cacheKey: "myCacheKey", // changing this key will invalidate this items cache
        iconProps: {
          iconName: "Group"
        },
        ariaLabel: "New Group",
        onClick: this.onOpenGroupForm
      },
      {
        key: "addUserToGroup",
        name: "Add user to group",
        iconProps: {
          iconName: "AddGroup"
        },
        disabled: !selectedGroupId,
        onClick: this.onOpenUserForm
      }
    ];
  };

  public render(): JSX.Element {
    const {
      groupColumns,
      memberColumns,
      groups,
      members,
      selectedGroupId,
      selectedUserId,
      isFormGroupOpen,
      isFormUserOpen,
      isDeleteFormOpen,
      loading
    } = this.state;
    //selectionId && this._getGroupMembers(Number(selectionId));
    // background: "#f3f3f3"
    return (
      <div>
        <Separator>
          <Text>Group manager</Text>
        </Separator>

        <CommandBar
          items={this._getMenuItems()}
          //overflowItems={this.getOverlflowItems()}
        />

        <Stack
          horizontal
          horizontalAlign="space-evenly"
          tokens={wrapStackTokens}
          style={{ marginBottom: 30, marginTop: 30 }}
        >
          <Stack.Item styles={stackStyles}>
            <Text>Groups</Text>
          </Stack.Item>
          <Stack.Item styles={stackStyles}>
            <Text>
              {selectedGroupId
                ? `Members of ${
                    groups.find(g => g.key === selectedGroupId).groupName
                  }`
                : "Select a group to display its members"}
            </Text>
          </Stack.Item>
        </Stack>

        <Stack
          horizontal
          horizontalAlign="space-evenly"
          // wrap
          tokens={wrapStackTokens}
          style={{ marginBottom: 30, marginTop: 30 }}
        >
          <DetailsList
            items={groups}
            columns={groupColumns}
            setKey="setOfGroups"
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selectionForGroups}
            selectionMode={SelectionMode.single}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
            //onItemInvoked={this._onItemInvoked}
          />

          <DetailsList
            items={members}
            columns={memberColumns}
            setKey="setOfMembers"
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selectionForUsers}
            selectionMode={SelectionMode.single}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Members"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
            //onItemInvoked={this._onItemInvoked}
          />
        </Stack>

        {isFormGroupOpen && (
          <GroupForm
            onCloseForm={this.onCloseGroupForm}
            updateComponent={this.getGroups}
            isOpenForm={isFormGroupOpen}
          />
        )}

        {isFormUserOpen && (
          <UserForm
            groupId={Number(selectedGroupId)}
            onCloseForm={this.onCloseUserForm}
            updateGroupMembers={this.updateGroupMembers}
            isOpenForm={isFormUserOpen}
          />
        )}

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
              disabled={loading}
              onClick={
                selectedGroupId && selectedUserId
                  ? this._deleteGroupMember
                  : this._deleteGroup
              }
              text="Delete"
            />
          </div>
        </Dialog>
      </div>
    );
  }

  public async componentDidMount() {
    await this.getGroups();
  }
  public clearUsrSelection() {
    this.setState({ selectedUserId: null });
  }

  public updateGroupMembers = () => {
    this._getGroupMembers(Number(this._getSelectionForGroups()));
  };

  private _onOpenDeleteForm = () => {
    this.setState({ isDeleteFormOpen: true });
  };

  private _onCloseDeleteForm = () => {
    this.setState({ isDeleteFormOpen: false });
  };

  public onOpenGroupForm = () => {
    this.setState({ isFormGroupOpen: true });
  };

  public onCloseGroupForm = () => {
    this.setState({ isFormGroupOpen: false });
  };

  public onOpenUserForm = () => {
    this.setState({ isFormUserOpen: true });
  };
  public onCloseUserForm = () => {
    this.setState({ isFormUserOpen: false });
  };
}
