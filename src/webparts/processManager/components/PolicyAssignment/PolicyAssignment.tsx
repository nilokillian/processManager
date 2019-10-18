import * as React from "react";
import SharePointService from "../../../../services/SharePoint/SharePointService";
import {
  DefaultButton,
  Separator,
  Text,
  DocumentCard,
  DocumentCardActivity,
  DocumentCardTitle,
  Stack,
  IStackTokens,
  DocumentCardDetails
} from "office-ui-fabric-react";
import PolicyAssignmentForm from "./PolicyAssignmentFrom";

const wrapStackTokens: IStackTokens = { childrenGap: 20 };
const templateImage: ITemplateImage = {
  image: require("../../../images/templateImg.png"),
  icon: require("../../../images/templateIcon.jpg")
};

export interface ITemplateImage {
  image: string;
  icon: string;
}

export interface IGroupsAssigned {
  id: string;
  name: string;
}

export interface IPeopeleAssigned {
  id: string;
  name: string;
}

export interface IPolicyAssignmentState {
  policies: any[];
  isPolicyAssignmentFormOpen: boolean;
  currentPolicy: {
    id: number;
    name: string;
    groupsAssigned: IGroupsAssigned[];
    peopeleAssigned: IPeopeleAssigned[];
  };
}

export default class PolicyAssignment extends React.Component<
  {},
  IPolicyAssignmentState
> {
  constructor(props: {}) {
    super(props);

    this.state = {
      policies: [],
      isPolicyAssignmentFormOpen: false,
      currentPolicy: {
        id: null,
        name: "",
        groupsAssigned: [],
        peopeleAssigned: []
      }
    };
  }

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

    const policies = result.map(v => {
      return {
        id: v.ID,
        title: v.Title,
        peopeleAssigned:
          v.PeopeleAssigned &&
          v.PeopeleAssigned.map(user => ({ id: user.ID, name: user.Title })),
        groupsAssigned:
          v.GroupAssigned &&
          v.GroupAssigned.map(group => ({ id: group.ID, name: group.Title })),
        policyOwner: { title: v.PolicyOwner.Title, email: v.PolicyOwner.EMail }
      };
    });
    // const policies = result.map(v => {
    //   return {
    //     id: v.Id,
    //     title: v.Title,
    //     peopeleAssigned: v.PeopeleAssigned && v.PeopeleAssigned,
    //     groupsAssigned:
    //       v.GroupAssigned &&
    //       v.GroupAssigned.map(group => ({ id: group.ID, name: group.Title })),
    //     policyOwner: v.PolicyOwner && v.PolicyOwner.Title
    //   };
    // });

    this.setState({ policies });
  };

  public render(): JSX.Element {
    const { policies, isPolicyAssignmentFormOpen, currentPolicy } = this.state;
    return (
      <div>
        <Separator>
          <Text>Policy Assignment</Text>
        </Separator>
        <Stack>
          <Stack
            horizontal
            horizontalAlign="space-evenly"
            wrap
            tokens={wrapStackTokens}
            style={{ marginBottom: 30, marginTop: 30 }}
          >
            {policies.map(p => (
              <DocumentCard>
                <DocumentCardDetails>
                  <DocumentCardTitle title={p.title} shouldTruncate={true} />
                  <DocumentCardActivity
                    activity={p.policyOwner.email}
                    people={[
                      {
                        name: p.policyOwner.title,
                        profileImageSrc: templateImage.icon
                      }
                    ]}
                  />
                  <DefaultButton
                    text="Assign people"
                    onClick={() =>
                      this.onOpenPolicyAssignmentForm(
                        p.id,
                        p.title,
                        p.groupsAssigned,
                        p.peopeleAssigned
                      )
                    }
                  />
                </DocumentCardDetails>
              </DocumentCard>
            ))}
          </Stack>
        </Stack>
        {currentPolicy.name && isPolicyAssignmentFormOpen && (
          <PolicyAssignmentForm
            currentPolicy={currentPolicy}
            onCloseForm={this.onClosePolicyAssignmentForm}
            isOpenForm={isPolicyAssignmentFormOpen}
          />
        )}
      </div>
    );
  }

  public async componentDidMount() {
    await this._getPolicies();
  }

  public onOpenPolicyAssignmentForm = (
    currentPolicyId: number,
    currentPolicyTitle: string,
    groupsAssigned: any[],
    peopeleAssigned: any[]
  ) => {
    const { currentPolicy } = this.state;
    currentPolicy.id = currentPolicyId;
    currentPolicy.name = currentPolicyTitle;
    currentPolicy.groupsAssigned = groupsAssigned;
    currentPolicy.peopeleAssigned = peopeleAssigned;

    this.setState({ isPolicyAssignmentFormOpen: true, currentPolicy });
  };

  public onClosePolicyAssignmentForm = () => {
    this.setState({ isPolicyAssignmentFormOpen: false });
  };
}
