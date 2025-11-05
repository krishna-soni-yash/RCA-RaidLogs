import * as React from 'react';
import styles from './RootCauseAnalysis.module.scss';
import type { IRootCauseAnalysisProps } from './IRootCauseAnalysisProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import RCATable from './RootCauseAnalysisTables/RCATable';
import Header from './Header/Header';


export default class RootCauseAnalysis extends React.Component<IRootCauseAnalysisProps> {
  public render(): React.ReactElement<IRootCauseAnalysisProps> {
    const {
     // description,
     // isDarkTheme,
     // environmentMessage,
      hasTeamsContext,
      //userDisplayName
      context
    } = this.props;

   

    return (
      <section className={`${styles.rootCauseAnalysis} ${hasTeamsContext ? styles.teams : ''}`}>
        <Header />
        <RCATable context={context} />
      </section>
    );
  }
}
