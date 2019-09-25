import * as React from "react";
import SharePointService from "../../../../services/SharePoint/SharePointService";

import {
  DefaultButton,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle,
  IDocumentCardPreviewProps,
  Separator,
  CommandBar,
  Text,
  Stack,
  IStackTokens
} from "office-ui-fabric-react";
import PolicyForm from "../page-builder/PolicyForm";

const wrapStackTokens: IStackTokens = { childrenGap: 20 };
const templateImage: ITemplateImage = {
  image: require("../../../images/templateImg.png"),
  icon: require("../../../images/templateIcon.jpg")
};

export interface ITemplateImage {
  image: string;
  icon: string;
}

export interface IPolicyState {
  polices: any[];
  isPolicyFormOpen: boolean;
  templates: any[];
}

export default class Policies extends React.Component<{}, IPolicyState> {
  constructor(props: {}) {
    super(props);

    this.state = {
      polices: [],
      isPolicyFormOpen: false,
      templates: []
    };
  }

  public async componentDidMount() {
    await this._getPolicies();
    await this._getPolicyPagesOptions();
  }

  public onOpenPolicyForm = () => {
    this.setState({ isPolicyFormOpen: true });
  };
  public onClosePolicyForm = () => {
    this.setState({ isPolicyFormOpen: false });
  };

  public render(): JSX.Element {
    const { polices, isPolicyFormOpen, templates } = this.state;
    return (
      <div>
        <Separator>
          <Text>Polies</Text>
        </Separator>

        <CommandBar
          items={this._getMenuItems()}
          //overflowItems={this.getOverlflowItems()}
        />
        <Stack
          horizontal
          horizontalAlign="space-evenly"
          wrap
          tokens={wrapStackTokens}
          style={{ marginBottom: 30, marginTop: 30 }}
        >
          {polices.map(policy => (
            <DocumentCard>
              {/* <DocumentCardPreview {...previewProps} /> */}
              <DocumentCardTitle title={policy.title} shouldTruncate={true} />
              <DocumentCardActivity
                activity="Created a few minutes ago"
                people={[
                  { name: policy.title, profileImageSrc: templateImage.icon }
                ]}
              />
              <DefaultButton text="Update policy" onClick={null} />
            </DocumentCard>
          ))}
        </Stack>
        {isPolicyFormOpen && (
          <PolicyForm
            templates={templates}
            onCloseForm={this.onClosePolicyForm}
            isOpenForm={isPolicyFormOpen}
          />
        )}
      </div>
    );
  }

  private _getMenuItems = () => {
    return [
      {
        key: "newPolicy",
        name: "Create new policy",
        iconProps: {
          iconName: "EntitlementPolicy"
        },

        onClick: this.onOpenPolicyForm
      }
    ];
  };

  private _getPolicies = async () => {
    const fields = [
      { key: "ID", fieldType: "Counter" },
      { key: "Title", fieldType: "Text" },
      { key: "PeopleAssigned", fieldType: "UserMulti", lookupField: "Title" },
      { key: "GroupAssigned", fieldType: "UserMulti", lookupField: "Title" },
      { key: "PolicyOwner", fieldType: "User", lookupField: "Title" },
      { key: "PolicyPagesTitle", fieldType: "Text" }
    ];

    const expandField = SharePointService.createExpendedFields(fields);
    const queriedField = SharePointService.createQueriedFields(fields);

    const result = await SharePointService.pnp_getItemsByTitle(
      "Policies",
      expandField,
      queriedField
    );

    console.log("result", result);
    const polices = result.map(v => {
      return {
        title: v.Title,
        peopeleAssigned: v.PeopeleAssigned && v.PeopeleAssigned.Title,
        groupAssigned: v.GroupAssigned && v.GroupAssigned.Title,
        policyOwner: v.PolicyOwner && v.PolicyOwner.Title
      };
    });
    this.setState({ polices });
  };

  private _getPolicyPagesOptions = async () => {
    const result = await SharePointService.getPolicyPages(
      "SitePages",
      "Templates"
    );

    const templates = result.map(template => ({
      key: template.Name.split(".")[0],
      title: template.Title,
      name: template.Name
    }));

    this.setState({ templates });
  };

  private _updateTasks = () => {
    //get all items
    //filter based on status
  };
}
