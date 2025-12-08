import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Current_User_Role } from '../../common/Constants';
import { IPPOApprovers } from '../../models/PPOApprovers';

export interface IPPOApproverResult {
	approver: IPPOApprovers | null;
	currentUserRole: Current_User_Role;
	currentUserRoles: Current_User_Role[];
}

export default interface IPPOApproversRepository {
	getApproversForCurrentSite(context: WebPartContext): Promise<IPPOApproverResult>;
}
