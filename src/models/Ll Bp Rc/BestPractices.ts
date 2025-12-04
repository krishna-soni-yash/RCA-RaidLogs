export interface IBestPractices {
    ID?: number;
    BpBestPracticesDescription?: string;
    BpReferences?: string;
    BpResponsibility?: string;
    BpResponsibilityId?: number | number[];
    BpResponsibilityEmail?: string | string[];
    BpResponsibilityLoginName?: string | string[];
    BpRemarks?: string;
    DataType?: string;
}

export const BestPracticesDataType = 'BestPractices';