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

//RCA Site Configuration

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
    minWidth: 120,
    maxWidth: 220,
    isResizable: true
  },
  { key: 'causeCategory', name: 'Cause Category', fieldName: 'CauseCategory', minWidth: 80, maxWidth: 150, isResizable: true },
  { key: 'source', name: 'Source', fieldName: 'RCASource', minWidth: 80, maxWidth: 150, isResizable: true },
  { key: 'priority', name: 'Priority', fieldName: 'RCAPriority', minWidth: 60, maxWidth: 100, isResizable: true },
  { key: 'relatedMetric', name: 'Related Metric (if any)', fieldName: 'RelatedMetric', minWidth: 110, maxWidth: 180, isResizable: true },
  { key: 'causes', name: 'Cause(s)', fieldName: 'Cause', minWidth: 120, maxWidth: 200, isResizable: true },
  { key: 'rootCauses', name: 'Root Cause(s)', fieldName: 'RootCause', minWidth: 120, maxWidth: 200, isResizable: true },
  { key: 'analysisTechnique', name: 'Root Cause Analysis Technique Used and Reference (if any)', fieldName: 'RCATechniqueUsedAndReference', minWidth: 140, maxWidth: 240, isResizable: true },
  { key: 'actionType', name: 'Type of Action', fieldName: 'RCATypeOfAction', minWidth: 100, maxWidth: 140, isResizable: true },
  { key: 'performanceBefore', name: 'Performance before action plan', fieldName: 'PerformanceBeforeActionPlan', minWidth: 120, maxWidth: 180, isResizable: true },
  { key: 'performanceAfter', name: 'Performance after action plan', fieldName: 'PerformanceAfterActionPlan', minWidth: 120, maxWidth: 180, isResizable: true },
  { key: 'quantitativeEffectiveness', name: 'Quantitative / Statistical effectiveness', fieldName: 'QuantitativeOrStatisticalEffecti', minWidth: 140, maxWidth: 200, isResizable: true },
  { key: 'remarks', name: 'Remarks', fieldName: 'Remarks', minWidth: 100, maxWidth: 180, isResizable: true }
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
  'Remarks',
];
export const expandFields = [
  'ResponsibilityCorrection',
  'ResponsibilityCorrective',
  'ResponsibilityPreventive'
];

//RAID Site Configuration
/**
 * SharePoint List Names
 */
export const LIST_NAMES = {
  RAID_LOGS: 'RAIDLogs',
  // Add other list names here as needed
} as const;

/**
 * RAID Types
 */
export const RAID_TYPES = {
  RISK: 'Risk',
  ASSUMPTION: 'Assumption',
  ISSUE: 'Issue',
  DEPENDENCY: 'Dependency',
  OPPORTUNITY: 'Opportunity',
  CONSTRAINTS: 'Constraints'
} as const;

/**
 * RAID Status Values
 */
export const RAID_STATUS = {
  OPEN: 'Open',
  IN_PROGRESS: 'In Progress',
  CLOSED: 'Closed',
  ON_HOLD: 'On Hold',
  CANCELLED: 'Cancelled'
} as const;

/**
 * RAID Priority Values
 */
export const RAID_PRIORITY = {
  LOW: 'Low',
  MEDIUM: 'Medium',
  HIGH: 'High',
  CRITICAL: 'Critical'
} as const;

/**
 * Action Types for RAID Items
 */
export const ACTION_TYPES = {
  MITIGATION: 'Mitigation',
  CONTINGENCY: 'Contingency',
} as const;

/**
 * Form Types
 */
export const FORM_TYPES = {
  RISK_FORM: 'RiskForm',
  OPPORTUNITY_FORM: 'OpportunityForm',
  ISSUE_FORM: 'IssueForm',
  ASSUMPTION_FORM: 'AssumptionForm',
  DEPENDENCY_FORM: 'DependencyForm',
  CONSTRAINTS_FORM: 'ConstraintsForm'
} as const;

/**
 * Field Names (Internal SharePoint Field Names)
 */
export const FIELD_NAMES = {
  ID: 'Id',
  TITLE: 'Title',
  SELECT_TYPE: 'SelectType',
  IDENTIFICATION_DATE: 'IdentificationDate',
  RISK_DESCRIPTION: 'RiskDescription',
  ASSOCIATED_GOAL: 'AssociatedGoal',
  RISK_SOURCE: 'RiskSource',
  RISK_CATEGORY: 'RiskCategory',
  IMPACT: 'Impact',
  RISK_PRIORITY: 'RiskPriority',
  IMPACT_VALUE: 'ImpactValue',
  PROBABILITY_VALUE: 'ProbabilityValue',
  RISK_EXPOSURE: 'RiskExposure',
  POTENTIAL_COST: 'PotentialCost',
  POTENTIAL_BENEFIT: 'PotentialBenefit',
  OPPORTUNITY_VALUE: 'OpportunityValue',
  TYPE_OF_ACTION: 'TypeOfAction',
  ACTION_PLAN: 'ActionPlan',
  RESPONSIBILITY: 'Responsibility',
  TARGET_DATE: 'TargetDate',
  ACTUAL_DATE: 'ActualDate',
  RISK_STATUS: 'RiskStatus',
  ISSUE_DETAILS: 'IssueDetails',
  IDA_DATE: 'IDADate',
  BY_WHOM: 'ByWhom',
  IMPLEMENTATION_ACTIONS: 'ImplementationActions',
  PLANNED_CLOSURE_DATE: 'PlannedClosureDate',
  ACTUAL_CLOSURE_DATE: 'ActualClosureDate',
  EFFECTIVENESS: 'Effectiveness',
  REMARKS: 'Remarks',
  RAID_ID: 'RAIDId',
  CREATED: 'Created',
  MODIFIED: 'Modified',
  AUTHOR_ID: 'AuthorId',
  EDITOR_ID: 'EditorId'
} as const;

/**
 * Default Values
 */
export const DEFAULT_VALUES = {
  LIST_NAME: LIST_NAMES.RAID_LOGS,
  PAGE_SIZE: 50,
  MAX_ITEMS: 5000
} as const;

/**
 * Error Messages
 */
export const ERROR_MESSAGES = {
  LIST_NOT_FOUND: 'SharePoint list not found',
  ITEM_NOT_FOUND: 'RAID item not found',
  CREATE_FAILED: 'Failed to create RAID item',
  UPDATE_FAILED: 'Failed to update RAID item',
  DELETE_FAILED: 'Failed to delete RAID item',
  INVALID_DATA: 'Invalid data provided',
  NETWORK_ERROR: 'Network error occurred',
  PERMISSION_DENIED: 'Permission denied'
} as const;

/**
 * Success Messages
 */
export const SUCCESS_MESSAGES = {
  ITEM_CREATED: 'RAID item created successfully',
  ITEM_UPDATED: 'RAID item updated successfully',
  ITEM_DELETED: 'RAID item deleted successfully',
  BULK_UPDATE_COMPLETED: 'Bulk update completed successfully'
} as const;

/**
 * Utility Functions
 */

/**
 * Generate a unique RAID ID for Risk items
 * Format: RAID-{timestamp}-{random4digits}
 * Example: RAID-20231111143025-4521
 */
export const generateRaidId = (): string => {
  const timestamp = new Date().toISOString().replace(/[-:TZ.]/g, '').substring(0, 14);
  const random = Math.floor(1000 + Math.random() * 9000); // 4-digit random number
  return `RAID-${timestamp}-${random}`;
};

/**
 * Dropdown Options for RAID Forms
 */
export const DROPDOWN_OPTIONS = {
  ASSOCIATED_GOAL: [
    { key: 'BG01', text: 'BG01' },
    { key: 'BG02', text: 'BG02' }
  ],
  SOURCE: [
    { key: 'Internal', text: 'Internal' },
    { key: 'External', text: 'External' }
  ],
  CATEGORY: [
    { key: 'Resource', text: 'Resource' },
    { key: 'Customer', text: 'Customer' },
    { key: 'Information Security', text: 'Information Security' },
    { key: 'Technology', text: 'Technology' },
    { key: 'Business', text: 'Business' },
    { key: 'Process', text: 'Process' },
    { key: 'Others', text: 'Others' }
  ],
  PRIORITY: [
    { key: 'High', text: 'High' },
    { key: 'Medium', text: 'Medium' },
    { key: 'Low', text: 'Low' }
  ],
  STATUS: [
    { key: 'Open', text: 'Open' },
    { key: 'Closed', text: 'Closed' },
    { key: 'In Progress', text: 'In Progress' },
    { key: 'Monitoring', text: 'Monitoring' }
  ],
  ACTION_TYPE: [
    { key: 'Mitigation', text: 'Mitigation' },
    { key: 'Contingency', text: 'Contingency' }
  ],
  POTENTIAL_COST: [
    { key: '1', text: '1 - No Cost' },
    { key: '2', text: '2 - Very Low Cost' },
    { key: '3', text: '3 - Low Cost' },
    { key: '4', text: '4 - Below Medium Cost' },
    { key: '5', text: '5 - Medium Cost' },
    { key: '6', text: '6 - Above Medium Cost' },
    { key: '7', text: '7 - High Cost' },
    { key: '8', text: '8 - Above High Cost' },
    { key: '9', text: '9 - Very High Cost' },
    { key: '10', text: '10 - Extreme High Cost' }
  ],
  POTENTIAL_BENEFIT: [
    { key: '1', text: '1 - No Benefit' },
    { key: '2', text: '2 - Very Low Benefit' },
    { key: '3', text: '3 - Low Benefit' },
    { key: '4', text: '4 - Below Moderate Benefit' },
    { key: '5', text: '5 - Moderate Benefit' },
    { key: '6', text: '6 - Above Moderate Benefit' },
    { key: '7', text: '7 - High Benefit' },
    { key: '8', text: '8 - Above High Benefit' },
    { key: '9', text: '9 - Very High Benefit' },
    { key: '10', text: '10 - Extreme High Benefit' }
  ],
  PROBABILITY_VALUE: [
    { key: '1', text: '1 - Not likely to occur' },
    { key: '2', text: '2 - Not very likely to occur' },
    { key: '3', text: '3 - Somewhat less than an even chance' },
    { key: '4', text: '4 - An even chance to occur' },
    { key: '5', text: '5 - Somewhat greater than an even chance' },
    { key: '6', text: '6 - Likely to occur' },
    { key: '7', text: '7 - Very likely to occur' },
    { key: '8', text: '8 - Almost sure to occur' },
    { key: '9', text: '9 - Extremely sure to occur' },
    { key: '10', text: '10 - Certain to occur' }
  ],
  IMPACT_VALUE: [
    { key: '1', text: '1 - Very Low Impact' },
    { key: '2', text: '2 - Low Impact - Needs No Attention' },
    { key: '3', text: '3 - Has a minor impact and things are taken into considerations by the respective projects / departments' },
    { key: '4', text: '4 - Medium impact and needs attention for resolving the issue' },
    { key: '5', text: '5 - Has an impact and resolves in co-ordination with BUH / CEO- and DH' },
    { key: '6', text: '6 - Has an high impact and which effects business value' },
    { key: '7', text: '7 - Has an major impact and which has significant effect on business' },
    { key: '8', text: '8 - Has a critical impact and call for a review with senior management' },
    { key: '9', text: '9 - Has a very critical impact and needs senior management attention / CISO' },
    { key: '10', text: '10 - Leads to serious problems and call for emergency mode' }
  ]
};





