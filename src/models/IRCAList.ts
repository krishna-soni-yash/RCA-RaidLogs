export interface IRCAList {
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
}
