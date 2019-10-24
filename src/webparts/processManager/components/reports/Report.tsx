import * as React from "react";
import SharePointService from "../../../../services/SharePoint/SharePointService";
import {
  ChartControl,
  ChartType
} from "@pnp/spfx-controls-react/lib/ChartControl";
import { Separator, Text, Stack, IStackTokens } from "office-ui-fabric-react";

const wrapStackTokens: IStackTokens = { childrenGap: 20 };

export interface IReportsState {
  polices: any[];
  isPolicyFormOpen: boolean;
  policyPages: any[];
  selectedPolicyId: number;
}

export default class Reports extends React.Component<{}, IReportsState> {
  private _data: Chart.ChartData = {
    labels: ["January", "February", "March", "April", "May", "June", "July"],
    datasets: [
      {
        label: "My First Dataset",
        data: [65, 59, 80, 81, 56, 55, 40],
        backgroundColor: [
          "rgba(255, 99, 132, 0.2)",
          "rgba(255, 159, 64, 0.2)",
          "rgba(255, 205, 86, 0.2)",
          "rgba(75, 192, 192, 0.2)",
          "rgba(54, 162, 235, 0.2)",
          "rgba(153, 102, 255, 0.2)",
          "rgba(201, 203, 207, 0.2)"
        ],
        borderColor: [
          "rgb(255, 99, 132)",
          "rgb(255, 159, 64)",
          "rgb(255, 205, 86)",
          "rgb(75, 192, 192)",
          "rgb(54, 162, 235)",
          "rgb(153, 102, 255)",
          "rgb(201, 203, 207)"
        ],
        borderWidth: 1
      }
    ]
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
    return (
      <div>
        <Separator>
          <Text>Reports</Text>
        </Separator>
        <ChartControl
          type={ChartType.Bar}
          data={this._data}
          //options={this._options}
        />
      </div>
    );
  }

  public async componentDidMount() {}
}
