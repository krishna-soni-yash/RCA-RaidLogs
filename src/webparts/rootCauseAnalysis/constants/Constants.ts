/**
 * Application Constants
 * Contains all constant values used across the application
 */

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
  OPPORTUNITY: 'Opportunity'
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
  ACCEPTANCE: 'Acceptance',
  TRANSFER: 'Transfer',
  AVOIDANCE: 'Avoidance'
} as const;

/**
 * Form Types
 */
export const FORM_TYPES = {
  RISK_FORM: 'RiskForm',
  OPPORTUNITY_FORM: 'OpportunityForm',
  ISSUE_FORM: 'IssueForm',
  ASSUMPTION_FORM: 'AssumptionForm',
  DEPENDENCY_FORM: 'DependencyForm'
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
  APPLICABILITY: 'Applicability',
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