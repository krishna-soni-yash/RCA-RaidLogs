import * as React from 'react';
import styles from './RootCauseAnalysis.module.scss';
import type { IRootCauseAnalysisProps } from './IRootCauseAnalysisProps';
import RCATable from './RootCauseAnalysisTables/RCATable';
import Header from './Header/Header';
import RaidLogs from './RaidLogs/RaidLogs';
import LlBpRc from './LL, BP & RC/LlBpRc';
import PPOApproversRepository from '../../../repositories/PPOApproversRepository';
import { IPPOApprovers } from '../../../models/PPOApprovers';
import { Current_User_Role } from '../../../common/Constants';
import PpoApproversContext from './PpoApproversContext';


export interface IRootCauseAnalysisState {
  activeTab: string;
  approver: IPPOApprovers | null;
  currentUserRole: Current_User_Role;
  isLoadingApprover: boolean;
}

export default class RootCauseAnalysis extends React.Component<IRootCauseAnalysisProps, IRootCauseAnalysisState> {
  private readonly ppoApproversRepository: PPOApproversRepository;

  constructor(props: IRootCauseAnalysisProps) {
    super(props);
    this.ppoApproversRepository = new PPOApproversRepository();
    this.state = {
      activeTab: 'raidLogs',
      approver: null,
      currentUserRole: Current_User_Role.None,
      isLoadingApprover: false
    };
  }
  public componentDidMount(): void {
    void this.loadPpoApprovers();
  }
  private handleTabChange = (tabKey: string): void => {
    this.setState({ activeTab: tabKey });
  };

  private loadPpoApprovers = async (): Promise<void> => {
    const { context } = this.props;

    if (!context) {
      return;
    }

    this.setState({ isLoadingApprover: true });

    try {
      const result = await this.ppoApproversRepository.getApproversForCurrentSite(context);

      this.setState({
        approver: result.approver,
        currentUserRole: result.currentUserRole,
        isLoadingApprover: false
      });
    } catch (error) {
      console.error('Failed to fetch PPO Approvers', error);
      this.setState({
        approver: null,
        currentUserRole: Current_User_Role.None,
        isLoadingApprover: false
      });
    }
  };

  public render(): React.ReactElement<IRootCauseAnalysisProps> {
    const { context, hasTeamsContext } = this.props;
    const { approver, currentUserRole, isLoadingApprover } = this.state;
    
    const renderContent = (): React.ReactElement => {
      switch (this.state.activeTab) {
        case 'rootCauseAnalysis':
          return <RCATable context={context} />
        case 'raidLogs':
          return <RaidLogs context={this.props.context} />;
        case 'lessonsLearnt':
          return <LlBpRc context={this.props.context} />;
        default:
          return <RCATable context={context} />
      }
    };
   

    return (
      <PpoApproversContext.Provider value={{
        approver,
        currentUserRole,
        isLoading: isLoadingApprover,
        reload: this.loadPpoApprovers
      }}>
        <section className={`${styles.rootCauseAnalysis} ${hasTeamsContext ? styles.teams : ''}`}>
          <Header  activeTab={this.state.activeTab}
            onTabChange={this.handleTabChange}/>
          {renderContent()}
        </section>
      </PpoApproversContext.Provider>
    );
  }
}
