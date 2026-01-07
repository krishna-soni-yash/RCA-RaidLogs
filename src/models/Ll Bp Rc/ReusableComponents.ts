export interface IReusableComponentAttachment {
    FileName: string;
    ServerRelativeUrl: string;
}

export interface IReusableComponents {
    ID?: number;
    RcComponentName?: string;
    RcLocation?: string;
    RcPurposeMainFunctionality?: string;
    RcRemarks?: string;
    DataType?: string;
    attachments?: IReusableComponentAttachment[];
    newAttachments?: File[];
}

export const ReusableComponentsDataType = 'ReusableComponents';