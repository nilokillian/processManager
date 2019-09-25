import * as React from "react";
//import { toast } from "react-toastify";

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
import PageTemplate from "../PageTemplate";

const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 5
};

export interface IPageFormProps {
  isOpenForm: boolean;
  onCloseForm(): void;
}

export interface IPageFormState {
  title: string;
  createdPageUrl: string;
  errors: object;
  loading: boolean;
}

export default class PageForm extends React.Component<
  IPageFormProps,
  IPageFormState
> {
  private _dragOptions = {
    moveMenuItemText: "Move",
    closeMenuItemText: "Close",
    menu: ContextualMenu
  };

  constructor(props: IPageFormProps) {
    super(props);

    this.state = {
      title: "",
      createdPageUrl: "",
      errors: {},
      loading: false
    };
  }

  public async componentDidMount() {}

  public render(): JSX.Element {
    const { isOpenForm, onCloseForm } = this.props;
    const { loading, title, createdPageUrl } = this.state;

    return (
      <div>
        <Panel
          isOpen={isOpenForm}
          type={PanelType.custom}
          customWidth="420px"
          onDismiss={onCloseForm}
          headerText={createdPageUrl ? "Template builder" : "Page Title"}
          closeButtonAriaLabel="Close"
          //onRenderHeader={this._onRenderHeaderContent}
          onRenderFooterContent={this._onRenderFooterContent}
          //styles={ComponentStyles.formPanelStyle}
        >
          <Stack
            //  styles={stackContainerStyles}
            tokens={itemAlignmentsStackTokens}
          >
            {!createdPageUrl && (
              <TextField
                id="Title"
                label="Title"
                value={title}
                onChange={this._onChangeTextInput}
                //styles={ComponentStyles.textInputStyle()}
                disabled={loading}
                required={true}
              />
            )}

            {createdPageUrl && (
              <div>
                <Text>
                  Click on created page, edit it and make sure you save it as a
                  template{" "}
                </Text>

                <PageTemplate pageTitle={title} url={createdPageUrl} />
              </div>
            )}
          </Stack>
        </Panel>

        {/* <Dialog
          hidden={!isDeleteCalloutVisible}
          onDismiss={this._closeDeleteBtnDialog}
          maxWidth={670}
          dialogContentProps={{
            type: DialogType.close,
            title: "Are you sure ?",
            subText:
              "This contact might be connected to existing tracking records. Breaking connections could cause unexpected issues"
          }}
          modalProps={{
            titleAriaId: this._labelId,
            dragOptions: this._dragOptions,
            isBlocking: false
            // styles: { main: { maxWidth: 750 } }
          }}
        >
          <div style={{ display: "flex", justifyContent: "center" }}>
            <DefaultButton
              style={{ backgroundColor: "#dc224d", color: "white" }}
              disabled={loading}
              onClick={this.delete}
              text="Delete"
            />
          </div>
        </Dialog> */}
      </div>
    );
  }

  private _onRenderFooterContent = () => {
    const { loading, errors, createdPageUrl } = this.state;
    const { onCloseForm } = this.props;
    return (
      <div>
        <PrimaryButton
          onClick={createdPageUrl ? onCloseForm : this.submitForm}
          text={createdPageUrl ? "Close" : "Save"}
          disabled={loading}
        />
      </div>
    );
  };

  private _onChangeTextInput = (
    e: React.FormEvent<HTMLInputElement>,
    newValue?: string
  ) => {
    // const currentFieldName = e.target["id"];

    const title = newValue;

    this.setState({ title });
  };

  //   private _getPeoplePickerItems = (items: any[]) => {
  //     const { data } = this.state;

  //     if (items.length === 1) {
  //       currentItem.Display_Name = items[0].text;
  //       currentItem.Email = items[0].secondaryText.toLowerCase();
  //       currentItem.First_Name = items[0].text.split(" ")[0];
  //       currentItem.Last_Name = items[0].text.split(" ")[1];
  //     } else if (items.length === 0) {
  //       currentItem.Display_Name = "";
  //       currentItem.Email = "";
  //       currentItem.First_Name = "";
  //       currentItem.Last_Name = "";
  //     }

  //     this.setState({ currentItem });
  //   };

  //   private _showDeleteBtnDialog = () => {
  //     this.setState({ isDeleteCalloutVisible: true });
  //   };

  //   private _closeDeleteBtnDialog = () => {
  //     this.setState({ isDeleteCalloutVisible: false });
  //   };

  public submitForm = async () => {
    const { onCloseForm } = this.props;
    const { title } = this.state;
    this.setState({ loading: true });

    try {
      const pageTitle = await SharePointService.pnp_CreatePage(title);
      const createdPageUrl = `${SharePointService.context.pageContext.web.absoluteUrl}/SitePages/${pageTitle.title}.aspx`;
      this.setState({ loading: false, createdPageUrl });
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
