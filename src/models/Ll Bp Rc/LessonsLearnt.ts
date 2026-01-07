export interface ILessonsLearntAttachment {
    FileName: string;
    ServerRelativeUrl: string;
}

export interface ILessonsLearnt {
    ID?: number;
    LlProblemFacedLearning?: string;
    LlCategory?: string;
    LlSolution?: string;
    LlRemarks?: string;
    DataType?: string;
    attachments?: ILessonsLearntAttachment[];
    newAttachments?: File[];
}

export const LessonsLearntDataType = 'LessonsLearnt';