import * as React from "react";
import SharePointService from "../../../../services/SharePoint/SharePointService";
import {
  DefaultButton,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardTitle,
  Separator,
  CommandBar,
  Text,
  Stack,
  IStackTokens,
  DocumentCardDetails
} from "office-ui-fabric-react";
import PolicyForm from "./PolicyForm";

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
  policyPages: any[];
  selectedPolicyId: number;
}

export default class Policies extends React.Component<{}, IPolicyState> {
  private _getMenuItems = () => {
    return [
      {
        key: "newPolicy",
        name: "Create new policy",
        iconProps: {
          iconName: "Add"
        },
        onClick: this._onOpenPolicyForm
      }
    ];
  };

  private _editPolicy = (policyId: number) => {
    this.setState({ selectedPolicyId: policyId, isPolicyFormOpen: true });
  };

  private _onOpenPolicyForm = () => {
    this.setState({ isPolicyFormOpen: true });
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

    const polices = result.map(v => {
      return {
        id: v.ID,
        title: v.Title,
        peopeleAssigned: v.PeopeleAssigned && v.PeopeleAssigned.Title,
        groupAssigned: v.GroupAssigned && v.GroupAssigned.Title,
        policyOwner: { title: v.PolicyOwner.Title, email: v.PolicyOwner.EMail }
      };
    });
    this.setState({ polices });
  };

  private _getPolicyPagesOptions = async () => {
    const result = await SharePointService.getPolicyPages(
      "SitePages",
      "Templates"
    );

    const policyPages = result.map(p => ({
      key: p.Name.split(".")[0],
      title: p.Title,
      name: p.Name
    }));

    this.setState({ policyPages });
  };

  constructor(props: {}) {
    super(props);

    this.state = {
      polices: [],
      isPolicyFormOpen: false,
      policyPages: [],
      selectedPolicyId: null
    };
  }

  public render(): JSX.Element {
    const {
      polices,
      isPolicyFormOpen,
      policyPages,
      selectedPolicyId
    } = this.state;
    return (
      <div>
        <Separator>
          <Text>Polices</Text>
        </Separator>

        <CommandBar items={this._getMenuItems()} />
        <Stack
          horizontal
          horizontalAlign="space-evenly"
          wrap
          tokens={wrapStackTokens}
          style={{ marginBottom: 30, marginTop: 30 }}
        >
          {polices.map(policy => (
            <DocumentCard>
              <DocumentCardDetails>
                <DocumentCardTitle title={policy.title} shouldTruncate={true} />
                <DocumentCardActivity
                  activity={policy.policyOwner.email}
                  people={[
                    {
                      name: policy.policyOwner.title,
                      profileImageSrc: templateImage.icon
                    }
                  ]}
                />
                <DefaultButton
                  text="details"
                  onClick={() => this._editPolicy(policy.id)}
                />
              </DocumentCardDetails>
            </DocumentCard>
          ))}
        </Stack>
        {isPolicyFormOpen && (
          <PolicyForm
            selectedPolicyId={selectedPolicyId}
            policyPages={policyPages}
            onCloseForm={this.onClosePolicyForm}
            resetSelectedPolicyId={this.resetSelectedPolicyId}
            updateComponent={this._getPolicies}
            isOpenForm={isPolicyFormOpen}
          />
        )}
      </div>
    );
  }

  public async componentDidMount() {
    await this._getPolicies();
    await this._getPolicyPagesOptions();
  }

  public resetSelectedPolicyId = () => {
    this.setState({ selectedPolicyId: null });
  };

  public onClosePolicyForm = () => {
    this.setState({ isPolicyFormOpen: false });
  };
}
