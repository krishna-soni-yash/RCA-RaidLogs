export interface IBestPracticeAttachment {
    FileName: string;
    ServerRelativeUrl: string;
}

export interface IBestPractices {
    ID?: number;
    BpBestPracticesDescription?: string;
    BpCategory?: string;
    BpReferences?: string;
    BpRemarks?: string;
    DataType?: string;
    attachments?: IBestPracticeAttachment[];
    newAttachments?: File[];
}

export const BestPracticesDataType = 'BestPractices';