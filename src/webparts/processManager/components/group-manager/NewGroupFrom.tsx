import * as React from "react";
import { toast } from "react-toastify";
import SharePointService from "../../../../services/SharePoint/SharePointService";
import {
  DefaultButton,
  Dialog,
  PrimaryButton,
  ContextualMenu,
  DialogType,
  Panel,
  PanelType,
  Stack,
  IStackTokens,
  TextField
} from "office-ui-fabric-react";

const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 5
};

export interface IGroupFormProps {
  isOpenForm: boolean;
  onCloseForm(): void;
  updateComponent(): void;
}

export interface IGroupOwner {
  id: string;
  name: string;
}

export interface IGroupData {
  title: string;
  owner: IGroupOwner;
  groupName: string;
}

export interface IGroupFormState {
  data: IGroupData;
  errors: object;
  loading: boolean;
}

export default class GroupForm extends React.Component<
  IGroupFormProps,
  IGroupFormState
> {
  private _dragOptions = {
    moveMenuItemText: "Move",
    closeMenuItemText: "Close",
    menu: ContextualMenu
  };

  constructor(props: IGroupFormProps) {
    super(props);

    this.state = {
      data: {} as IGroupData,
      errors: {},
      loading: false
    };
  }

  public async componentDidMount() {}

  public render(): JSX.Element {
    const { isOpenForm, onCloseForm } = this.props;
    const { loading, data } = this.state;

    return (
      <div>
        <Panel
          isOpen={isOpenForm}
          type={PanelType.custom}
          customWidth="420px"
          onDismiss={onCloseForm}
          headerText="New group"
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
              value={data.groupName}
              onChange={this._onChangeTextInput}
              //styles={ComponentStyles.textInputStyle()}
              disabled={loading}
              required={true}
            />
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

  private _onChangeTextInput = (
    e: React.FormEvent<HTMLInputElement>,
    newValue?: string
  ) => {
    const { data } = this.state;

    data.groupName = newValue;

    this.setState({ data });
  };

  private _createNewGroup = async () => {
    const { groupName } = this.state.data;
    await SharePointService.pnp_createGroupV2(groupName);
  };

  public submitForm = async () => {
    const { onCloseForm, updateComponent } = this.props;
    this.setState({ loading: true });

    try {
      await this._createNewGroup();
      updateComponent();
      toast.success("created");
      onCloseForm();
    } catch (error) {
      toast.error("error");
      onCloseForm();
      throw error;
    }
    this.setState({ loading: false });
  };

  public _getCurrentItem = async () => {};
}
