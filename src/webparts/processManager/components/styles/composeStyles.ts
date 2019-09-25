import { getTheme } from "@uifabric/styling";

export class StylesManger {
  private currentSPTeam = getTheme();
  private semanticColors = this.currentSPTeam.semanticColors;

  public getBlockMenuStyle = () => {
    return {
      background: { background: this.semanticColors.buttonBackground },
      blockSize: { width: "300px", height: "200px" }
    };
  };
}

const ComposeStyles = new StylesManger();
export default ComposeStyles;
