import * as React from 'react';
import styles from './RootCauseAnalysis.module.scss';
import type { IRootCauseAnalysisProps } from './IRootCauseAnalysisProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import RCATable from './RootCauseAnalysisTables/RCATable';
import Header from './Header/Header';
import RaidLogs from './RaidLogs/RaidLogs';


export interface IRootCauseAnalysisState {
  activeTab: string;
}

export default class RootCauseAnalysis extends React.Component<IRootCauseAnalysisProps, IRootCauseAnalysisState> {
 constructor(props: IRootCauseAnalysisProps) {
    super(props);
    this.state = {
      activeTab: 'raidLogs'
    };
  }
  private handleTabChange = (tabKey: string): void => {
    this.setState({ activeTab: tabKey });
  };

  public render(): React.ReactElement<IRootCauseAnalysisProps> {
    const { context, hasTeamsContext } = this.props;
    
    const renderContent = (): React.ReactElement => {
      switch (this.state.activeTab) {
        case 'rootCauseAnalysis':
          return <RCATable context={context} />
        case 'raidLogs':
          return <RaidLogs context={this.props.context} />;
        default:
          return <RCATable context={context} />
      }
    };
   

    return (
      <section className={`${styles.rootCauseAnalysis} ${hasTeamsContext ? styles.teams : ''}`}>
        <Header  activeTab={this.state.activeTab}
          onTabChange={this.handleTabChange}/>
        {renderContent()}
      </section>
    );
  }
}
