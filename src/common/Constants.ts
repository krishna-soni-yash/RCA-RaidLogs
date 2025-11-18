import { IColumnConfig } from '../webparts/rootCauseAnalysis/components/RootCauseAnalysisTables/RCATable';

export default class ParentListNames {
  public static AppSettings: string = "AppSettings";
  public static ObjectivesMaster: string = "ObjectivesMaster";
  public static PPOApprovers: string = "PPOApprovers";
  public static ProjectType: string = "ProjectType";
  public static AssociatedPPM: string = "AssociatedPPM"
  public static Metrics: string = "Metrics";
  public static MetricsMailSender: string = "MetricsMailSender";
}

export class SubSiteListNames {
  public static ProjectMetrics: string = "ProjectMetrics";
  public static ProjectMetricLogs: string = "ProjectMetricLogs";
  public static RootCauseAnalysis: string = "Root Cause Analysis";
}



export class SiteConfiguration {
  public static readonly PARENT_LISTS = [
    ParentListNames.AppSettings,
    ParentListNames.ObjectivesMaster,
    ParentListNames.PPOApprovers,
    ParentListNames.ProjectType,
    ParentListNames.AssociatedPPM,
    ParentListNames.Metrics,
    ParentListNames.MetricsMailSender
  ];
}
export const RCACOLUMNS: (IColumnConfig)[] = [
  {
    key: 'problemStatement',
    name: 'Problem statement (Causal Analysis Trigger)',
    fieldName: 'LinkTitle',
    minWidth: 150,
    maxWidth: 350,
    isResizable: true
  },
  { key: 'causeCategory', name: 'Cause Category', fieldName: 'CauseCategory', minWidth: 100, maxWidth: 200, isResizable: true },
  { key: 'source', name: 'Source', fieldName: 'RCASource', minWidth: 100, maxWidth: 200, isResizable: true },
  { key: 'priority', name: 'Priority', fieldName: 'RCAPriority', minWidth: 80, maxWidth: 120, isResizable: true },
  { key: 'relatedMetric', name: 'Related Metric (if any)', fieldName: 'RelatedMetric', minWidth: 140, maxWidth: 250, isResizable: true },
  { key: 'causes', name: 'Cause(s)', fieldName: 'Cause', minWidth: 150, maxWidth: 300, isResizable: true },
  { key: 'rootCauses', name: 'Root Cause(s)', fieldName: 'RootCause', minWidth: 150, maxWidth: 300, isResizable: true },
  { key: 'analysisTechnique', name: 'Root Cause Analysis Technique Used and Reference (if any)', fieldName: 'RCATechniqueUsedAndReference', minWidth: 180, maxWidth: 350, isResizable: true },
  { key: 'actionType', name: 'Type of Action', fieldName: 'RCATypeOfAction', minWidth: 120, maxWidth: 200, isResizable: true },
  { key: 'performanceBefore', name: 'Performance before action plan', fieldName: 'PerformanceBeforeActionPlan', minWidth: 150, maxWidth: 220, isResizable: true },
  { key: 'performanceAfter', name: 'Performance after action plan', fieldName: 'PerformanceAfterActionPlan', minWidth: 150, maxWidth: 220, isResizable: true },
  { key: 'quantitativeEffectiveness', name: 'Quantitative / Statistical effectiveness', fieldName: 'QuantitativeOrStatisticalEffecti', minWidth: 180, maxWidth: 260, isResizable: true },
  { key: 'remarks', name: 'Remarks', fieldName: 'Remarks', minWidth: 120, maxWidth: 300, isResizable: true }
];
export const selectedFields = [
  'Id',
  'LinkTitle',
  'ProblemStatementNumber', 
  'CauseCategory',
  'RCASource',
  'RCAPriority',  
  'RelatedMetric',
  'Cause',
  'RootCause',
  'RCATechniqueUsedAndReference',
  'RCATypeOfAction',  
  'ActionPlanCorrection',
  'ResponsibilityCorrection/Id',
  'ResponsibilityCorrection/Title',
  'ResponsibilityCorrection/EMail',
  
  'PlannedClosureDateCorrection', 
  'ActualClosureDateCorrection',
  'ActionPlanCorrective',

  'ResponsibilityCorrective/Id',
  'ResponsibilityCorrective/Title',
  'ResponsibilityCorrective/EMail',

  'PlannedClosureDateCorrective',
  'ActualClosureDateCorrective',
  'ActionPlanPreventive',

  'ResponsibilityPreventive/Id', 
  'ResponsibilityPreventive/Title',
  'ResponsibilityPreventive/EMail',

  'PlannedClosureDatePreventive', 
  'ActualClosureDatePreventive',
  'PerformanceBeforeActionPlan',    
  'PerformanceAfterActionPlan',  
  'QuantitativeOrStatisticalEffecti',
  'Remarks'
];
export const expandFields = [
  'ResponsibilityCorrection',
  'ResponsibilityCorrective',
  'ResponsibilityPreventive'
];




