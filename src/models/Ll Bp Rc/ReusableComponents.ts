export interface IReusableComponents {
    ID?: number;
    RcComponentName?: string;
    RcLocation?: string;
    RcPurposeMainFunctionality?: string;
    RcResponsibility?: string;
    RcResponsibilityId?: number | number[];
    RcResponsibilityEmail?: string | string[];
    RcResponsibilityLoginName?: string | string[];
    RcRemarks?: string;
    DataType?: string;
}

export const ReusableComponentsDataType = 'ReusableComponents';