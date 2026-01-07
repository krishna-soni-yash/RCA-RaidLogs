import * as React from 'react';
import { IPPOApprovers } from '../../../models/PPOApprovers';
import { Current_User_Role } from '../../../common/Constants';

export interface IPpoApproversContext {
  approver: IPPOApprovers | null;
  currentUserRole: Current_User_Role;
  currentUserRoles: Current_User_Role[];
  isLoading: boolean;
  reload: () => Promise<void> | void;
}

export const defaultPpoApproversContext: IPpoApproversContext = {
  approver: null,
  currentUserRole: Current_User_Role.None,
  currentUserRoles: [],
  isLoading: false,
  reload: () => undefined
};

const PpoApproversContext = React.createContext<IPpoApproversContext>(defaultPpoApproversContext);

export default PpoApproversContext;
