import * as React from "react";
import "react-toastify/dist/ReactToastify.css";
import { ToastContainer } from "react-toastify";
import styles from "./ProcessManager.module.scss";
import { IProcessManagerProps } from "./IProcessManagerProps";
import PageBuilder from "./page-builder/PageBuilder";
import Policies from "./policy/Policies";
import BlockMenu from "./BlockMenu";
import PolicyAssignment from "./PolicyAssignment/PolicyAssignment";
import GroupManager from "./group-manager/GroupManager";
import TaskManager from "./task-manager/TaskManager";
import Reports from "./reports/Report";

export interface IProcessManagerState {
  activeComponents: { title: string; sortOrderNumber: number };
}

export default class ProcessManager extends React.Component<
  IProcessManagerProps,
  IProcessManagerState
> {
  constructor(props: IProcessManagerProps) {
    super(props);

    this.state = {
      activeComponents: { title: "pageBuilder", sortOrderNumber: 1 }
    };
  }
  public render(): React.ReactElement<IProcessManagerProps> {
    const { activeComponents } = this.state;
    return (
      <div className={styles.processManager}>
        <ToastContainer />
        <BlockMenu
          onComponentChange={(newComponentName, sortOrderNumber) =>
            this.onComponentChange(newComponentName, sortOrderNumber)
          }
          activeComponents={activeComponents}
        />
        {activeComponents.title === "pageBuilder" && <PageBuilder />}
        {activeComponents.title === "policies" && <Policies />}
        {activeComponents.title === "policyAssignment" && <PolicyAssignment />}
        {activeComponents.title === "taskManager" && <TaskManager />}
        {activeComponents.title === "groups" && <GroupManager />}
        {activeComponents.title === "reports" && <Reports />}
        {/* <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div> */}
      </div>
    );
  }

  public onComponentChange = (
    componentName: string,
    sortOrderNumber: number
  ) => {
    let { activeComponents } = this.state;

    activeComponents.title = componentName;
    activeComponents.sortOrderNumber = sortOrderNumber;
    this.setState({ activeComponents });
  };
}
