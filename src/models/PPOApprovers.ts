export interface IPPOApproverUser {
    id: number;
    title: string;
    email: string;
    loginName?: string;
}

export interface IPPOApprovers {
    ID: number;
    Title: string;
    InternalProjectName: string;
    Reviewer: IPPOApproverUser[] | null;
    BUH: IPPOApproverUser[] | null;
    ProjectManager: IPPOApproverUser[] | null;
}