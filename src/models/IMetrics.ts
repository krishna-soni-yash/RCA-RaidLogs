export interface IMetrics {
  ID?: number;
  LinkTitle?: string; // Key
  IsActive?: boolean;
  ProjectType?: string;
  NameOfMetrics?: string;
  MetricFormulae?: string;
  BG?: string;
  PG?: string;
  PerformanceGoals?: string;
  Metrics?: string;
  UnitOfMeasure?: string;
  Goal?: string;
  USL?: string;
  LSL?: string;
  MetricsFormulae?: string;
  Priority?: string;
  AssociatedPPM?: string;
  DataInput?: string;
  DataSource?: string;
  DataCollectionFrequency?: string;
  DataAnalysisFrequency?: string;
  BaselineAndRevisionFrequency?: string;
  Statistical?: string;
  Quantitative?: string;
  InterpretationGuidelines?: string;
  CausalAnalysisTrigger?: string;
  ProbabilityOfSuccessThreshold?: string;
  Process?: string;
  SubMetrics?: string;
  Subprocess?: string;
  SubGoal?: string;
  SubUnitOfMeasure?: string;
  SubMetricsFormulae?: string;
  SubUSL?: string;
  SubDataInput?: string;
  SubLSL?: string;
  SubDataSource?: string;
  SubDataCollectionFrequency?: string;
  SubDataAnalysisFrequency?: string;
  SubBaselineAndRevisionFrequency?: string;
  Applicability?: string;
  HasSubProcess?: boolean;
  OrgStatistical?: string;
  OrgInterpretationGuidelines?: string;
  OrgCausalAnalysisTrigger?: string;
  
}
export interface IObjectivesMasterForMetrics {
  ID?: number;
  LinkTitle?: string;
  
 
}