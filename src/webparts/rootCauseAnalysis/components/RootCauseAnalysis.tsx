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
    const {
     // description,
     // isDarkTheme,
     // environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    // example RCA items matching the RCACOLUMNS fieldNames
    const items = [
      {
        problemStatement: 'Spike in failed logins (Causal Analysis Trigger)',
        causeCategory: 'Authentication',
        source: 'AuthService logs',
        priority: 'High',
        relatedMetric: 'Login Failure Rate',
        causes: 'Incorrect token handling',
        rootCauses: 'Token validation race condition',
        analysisTechnique: '5 Whys; see ticket RCA-001',
        actionType: 'Corrective',
        actionPlan: 'Patch auth service to serialize token validation',
        responsibility: userDisplayName || 'Team A',
        plannedClosureDate: '2025-11-15',
        actualClosureDate: '',
        performanceBefore: 'Login success 92%',
        performanceAfter: '',
        quantitativeEffectiveness: '',
        remarks: 'Monitoring added'
      },
      {
        problemStatement: 'Delayed order processing',
        causeCategory: 'Process',
        source: 'OrderQueue metrics',
        priority: 'Medium',
        relatedMetric: 'Order Throughput',
        causes: 'Slow DB queries',
        rootCauses: 'Missing index on orders table',
        analysisTechnique: 'Fishbone diagram; DB explain plan',
        actionType: 'Preventive',
        actionPlan: 'Add index and optimize queries',
        responsibility: 'DB Team',
        plannedClosureDate: '2025-10-30',
        actualClosureDate: '2025-10-28',
        performanceBefore: 'Throughput 120 ops/hr',
        performanceAfter: 'Throughput 320 ops/hr',
        quantitativeEffectiveness: '166% increase',
        remarks: 'Validated on staging'
      }
    ];

    const renderContent = (): React.ReactElement => {
      switch (this.state.activeTab) {
        case 'rootCauseAnalysis':
          return <RCATable items={items} />;
        case 'raidLogs':
          return <RaidLogs context={this.props.context} />;
        default:
          return <RCATable items={items} />;
      }
    };

    return (
      <section className={`${styles.rootCauseAnalysis} ${hasTeamsContext ? styles.teams : ''}`}>
        <Header 
          activeTab={this.state.activeTab}
          onTabChange={this.handleTabChange}
        />
        {renderContent()}
      </section>
    );
  }
}
