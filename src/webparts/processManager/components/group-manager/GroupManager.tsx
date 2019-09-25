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
  DialogType
} from "office-ui-fabric-react";
import GroupForm from "./NewGroupFrom";
import UserForm from "./AddUserFrom";
import SharePointService from "../../../../services/SharePoint/SharePointService";

const stackTokens: IStackTokens = { childrenGap: 12 };
const wrapStackTokens: IStackTokens = { childrenGap: 20 };

// export interface IGroupManagerProps {
//   selectedGroups: any[];
// }

export interface IGroupManagerState {
  groupColumns: IColumn[];
  memberColumns: IColumn[];
  groups: any[];
  members: any[];
  selectedGroupId: string;
  selectedUserId: string;
  isFormGroupOpen: boolean;
  isFormUserOpen: boolean;
  isDeleteFormOpen: boolean;
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
        return this._selectionForGroups.getSelection()[0].key.toString();
      default:
        return this._selectionForGroups.getSelection()[0].key.toString();
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
        return this._selectionForUsers.getSelection()[0].key.toString();
      default:
        return this._selectionForUsers.getSelection()[0].key.toString();
    }
  };

  constructor(props: {}) {
    super(props);

    this._selectionForGroups = new Selection({
      onSelectionChanged: async () => {
        this.setState({ selectedGroupId: this._getSelectionForGroups() });

        await this._getGroupMembers(Number(this._getSelectionForGroups()));
      }
    });

    this._selectionForUsers = new Selection({
      onSelectionChanged: () => {
        this.setState({ selectedUserId: this._getSelectionForUsers() });
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
      //   {
      //     key: "ownerTitle",
      //     name: "Owner Title",
      //     fieldName: "ownerTitle",
      //     minWidth: 100,
      //     maxWidth: 400,
      //     isResizable: true

      //   }
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
      selectedGroupId: "",
      selectedUserId: "",
      groupColumns: _groupColumns,
      memberColumns: _memberColumns,
      isFormGroupOpen: false,
      isFormUserOpen: false,
      isDeleteFormOpen: false
    };
  }

  private _getGroups = async () => {
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
                text: "Delete"
                //onClick: this._onOpenDeleteForm
              }
            ]
          }}
        />
      )
    }));

    this.setState({ groups });
  };

  private _getGroupMembers = async (groupId: number) => {
    const result = await SharePointService.pnp_getGroupMembers([groupId]);
    console.log("resultM", result);
    const members = result.map(member => ({
      key: member.Id,
      displayName: member.Title,
      email: member.Email,
      edit: (
        <IconButton
          menuProps={{
            shouldFocusOnMount: true,
            items: [
              {
                key: "delete",
                text: "Details",
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

    try {
      const result = await SharePointService.pnp_deleteGroupMember(
        Number(selectedGroupId),
        Number(selectedUserId)
      );
      toast.success("deleted");
      this._onCloseDeleteForm();
    } catch (error) {
      toast.error("error");
      this._onCloseDeleteForm();
      console.log("error");
    }
  };

  private _getMenuItems = () => {
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

        onClick: this.onOpenUserForm
      }
    ];
  };

  public async componentDidMount() {
    await this._getGroups();
  }

  public render(): JSX.Element {
    const {
      groupColumns,
      memberColumns,
      groups,
      members,
      selectedGroupId,
      isFormGroupOpen,
      isFormUserOpen,
      isDeleteFormOpen
    } = this.state;
    //selectionId && this._getGroupMembers(Number(selectionId));
    // background: "#f3f3f3"
    return (
      <div>
        <DefaultButton
          onClick={async () => await SharePointService.pnp_getGroupOwner()}
          text="getOwner"
        />
        <DefaultButton
          onClick={async () => await SharePointService.pnp_createGroup()}
          text="ceate"
        />

        <Separator>
          <Text>Group manager</Text>
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

          <Stack verticalAlign="space-around" tokens={stackTokens}>
            <Text>Groups</Text>

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
          </Stack>

          <Stack tokens={stackTokens}>
            <Text>
              {selectedGroupId
                ? `Members of ${
                    groups.find(g => g.key.toString() === selectedGroupId)
                      .groupName
                  }`
                : "Select a group to display its members"}
            </Text>
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
        </Stack>

        {isFormGroupOpen && (
          <GroupForm
            onCloseForm={this.onCloseGroupForm}
            isOpenForm={isFormGroupOpen}
          />
        )}

        {isFormUserOpen && (
          <UserForm
            groupId={Number(selectedGroupId)}
            onCloseForm={this.onCloseUserForm}
            isOpenForm={isFormUserOpen}
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
                onClick={this._deleteGroupMember}
                text="Delete"
              />
            </div>
          </Dialog>
        )}
      </div>
    );
  }

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
