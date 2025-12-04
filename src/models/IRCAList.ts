export interface IRCAList {
    ID: number;
    Title: string;
    // RCAItemID: string;
    // RCAItemCreatedBy: string;
    // RCAItemCreatedDate: string;
    // RCAItemModifiedBy: string;
    // RCAItemModifiedDate: string;
    LinkTitle: string;
    ProblemStatementNumber: string;
    CauseCategory: string;
    RCASource: string;
    RCAPriority: string;
    RelatedMetric: string;
    Cause: string;
    RootCause: string;
    RCATechniqueUsedAndReference: string;
    RCATypeOfAction: string;
    
    ActionPlanCorrection: string;
    ResponsibilityCorrection: string;
    PlannedClosureDateCorrection: string;
    ActualClosureDateCorrection: string;
    
    ActionPlanCorrective: string;
    ResponsibilityCorrective: string;
    PlannedClosureDateCorrective: string;
    ActualClosureDateCorrective: string;
    
    ActionPlanPreventive: string;
    ResponsibilityPreventive: string;
    PlannedClosureDatePreventive: string;
    ActualClosureDatePreventive: string;

    PerformanceBeforeActionPlan: string;
    PerformanceAfterActionPlan: string;
    QuantitativeOrStatisticalEffecti: string;
    Remarks: string;
    RelatedSubMetric: string;
    attachments: any[];
}
