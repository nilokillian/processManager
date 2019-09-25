import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption
} from "@microsoft/sp-webpart-base";
import { PropertyFieldMultiSelect } from "@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect";
import * as strings from "ProcessManagerWebPartStrings";
import ProcessManager from "./components/ProcessManager";
import { IProcessManagerProps } from "./components/IProcessManagerProps";
import SharePointService from "../../services/SharePoint/SharePointService";

export interface IProcessManagerWebPartProps {
  description: string;
  securityGroups: any[];
}

export default class ProcessManagerWebPart extends BaseClientSideWebPart<
  IProcessManagerWebPartProps
> {
  private groups: any[];
  private groupOptions: IPropertyPaneDropdownOption[];
  private groupOptionsLoading: boolean = false;

  public render(): void {
    const element: React.ReactElement<
      IProcessManagerProps
    > = React.createElement(ProcessManager, {
      description: this.properties.description
    });

    ReactDom.render(element, this.domElement);
  }

  public async onInit(): Promise<void> {
    await super.onInit();
    SharePointService.setup(this.context);
    SharePointService.pnp_setup(this.context);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  // protected onPropertyPaneConfigurationStart(): void {
  //   this.groupOptionsLoading = true;
  //   this._getGroups()
  //     .then(groupOptions => {
  //       this.groupOptions = groupOptions;
  //       this.context.propertyPane.refresh();
  //       this.groupOptionsLoading = false;
  //       this.render();
  //     })
  //     .catch(error => console.log(error));
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldMultiSelect("securityGroups", {
                  key: "groups",
                  label: "Assignable groups",
                  options: this.groupOptions,
                  selectedKeys: this.properties.securityGroups,
                  disabled: this.groupOptionsLoading
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
