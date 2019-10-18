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
          headerText={createdPageUrl ? "Policy page builder" : "Page Title"}
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
              <p>
                <ul>
                  <li>click on edit</li>
                  <li>build the page a way you want</li>
                  <li>save it as a template</li>
                </ul>
                <PageTemplate pageTitle={title} url={createdPageUrl} />
              </p>
            )}
          </Stack>
        </Panel>
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
      onCloseForm();
      throw error;
    }
    this.setState({ loading: false });
  };
}
