import { WebPartContext } from '@microsoft/sp-webpart-base';
import ParentListNames, { Current_User_Role } from '../common/Constants';
import ErrorMessages from '../common/ErrorMessages';
import genericService from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { IPPOApprovers, IPPOApproverUser } from '../models/PPOApprovers';
import IPPOApproversRepositoryInterface, { IPPOApproverResult } from './repositoriesInterface/IPPOApprovers';

export default class PPOApproversRepository implements IPPOApproversRepositoryInterface {
  private service: IGenericService;

  constructor(service?: IGenericService) {
    this.service = service ?? genericService;
  }

  public setService(service: IGenericService): void {
    this.service = service;
  }

  public async getApproversForCurrentSite(context: WebPartContext): Promise<IPPOApproverResult> {
    if (!context) {
      throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_PPOAPPROVERS);
    }

    const candidateProjectNames = this.buildCandidateProjectNames(context);
    if (candidateProjectNames.length === 0) {
      return {
        approver: null,
        currentUserRole: Current_User_Role.None,
        currentUserRoles: []
      };
    }

    try {
      const filterParts: string[] = [];
      for (let i = 0; i < candidateProjectNames.length; i++) {
        filterParts.push(`InternalProjectName eq '${this.escapeODataValue(candidateProjectNames[i])}'`);
      }
      const filter = filterParts.join(' or ');

      const items = await this.service.fetchAllItems<any>({
        context,
        listTitle: ParentListNames.PPOApprovers,
        select: [
          'Id',
          'ID',
          'Title',
          'InternalProjectName',
          'Reviewer/Id',
          'Reviewer/ID',
          'Reviewer/Title',
          'Reviewer/EMail',
          'Reviewer/Name',
          'BUH/Id',
          'BUH/ID',
          'BUH/Title',
          'BUH/EMail',
          'BUH/Name',
          'ProjectManager/Id',
          'ProjectManager/ID',
          'ProjectManager/Title',
          'ProjectManager/EMail',
          'ProjectManager/Name'
        ],
        expand: ['Reviewer', 'BUH', 'ProjectManager'],
        filter,
        pageSize: 5
      });
      if (!items || items.length === 0) {
        return {
          approver: null,
          currentUserRole: Current_User_Role.None,
          currentUserRoles: []
        };
      }

      const matchedItem = this.pickBestMatch(items, candidateProjectNames);
      const approver = this.mapApprover(matchedItem);
      const roleResolution = this.resolveCurrentUserRoles(context, approver);
      return {
        approver,
        currentUserRole: roleResolution.primaryRole,
        currentUserRoles: roleResolution.roles
      };
    } catch (error: any) {
      console.error(ErrorMessages.FAILED_TO_FETCH_PPOAPPROVERS, error);
      return {
        approver: null,
        currentUserRole: Current_User_Role.None,
        currentUserRoles: []
      };
    }
  }

  private mapApprover(item: any): IPPOApprovers | null {
    if (!item) {
      return null;
    }

    const approver: IPPOApprovers = {
      ID: Number(item.Id || item.ID || 0),
      Title: item.Title || '',
      InternalProjectName: item.InternalProjectName || '',
      Reviewer: this.mapUsers(item.Reviewer),
      BUH: this.mapUsers(item.BUH),
      ProjectManager: this.mapUsers(item.ProjectManager)
    };

    return approver;
  }

  private mapUsers(value: any): IPPOApproverUser[] | null {
    const users: IPPOApproverUser[] = [];

    if (!value) {
      return null;
    }

    let source: any[] = [];

    if (Array.isArray(value)) {
      source = value;
    } else if (value.results && Array.isArray(value.results)) {
      source = value.results;
    } else {
      source = [value];
    }

    for (let i = 0; i < source.length; i++) {
      const mapped = this.mapSingleUser(source[i]);
      if (mapped) {
        users.push(mapped);
      }
    }

    return users.length > 0 ? users : null;
  }

  private mapSingleUser(user: any): IPPOApproverUser | null {
    if (!user) {
      return null;
    }

    const email = (user.EMail || user.Email || '').trim();
    if (!email) {
      return null;
    }

    return {
      id: Number(user.Id || user.ID || 0) || 0,
      title: user.Title || user.DisplayName || '',
      email: email,
      loginName: user.LoginName || user.Name || undefined
    };
  }

  private resolveCurrentUserRoles(context: WebPartContext, approver: IPPOApprovers | null): { roles: Current_User_Role[]; primaryRole: Current_User_Role } {
    if (!approver) {
      return { roles: [], primaryRole: Current_User_Role.None };
    }

    const currentUserEmail = (context.pageContext?.user?.email || '').trim().toLowerCase();
    if (!currentUserEmail) {
      return { roles: [], primaryRole: Current_User_Role.None };
    }

    const roles: Current_User_Role[] = [];

    if (this.hasMatchingUser(approver.ProjectManager, currentUserEmail)) {
      roles.push(Current_User_Role.ProjectManager);
    }

    if (this.hasMatchingUser(approver.BUH, currentUserEmail)) {
      roles.push(Current_User_Role.BUH);
    }

    if (this.hasMatchingUser(approver.Reviewer, currentUserEmail)) {
      roles.push(Current_User_Role.Reviewer);
    }

    let primaryRole = Current_User_Role.None;
    if (roles.length > 0) {
      if (roles.indexOf(Current_User_Role.ProjectManager) !== -1) {
        primaryRole = Current_User_Role.ProjectManager;
      } else {
        primaryRole = roles[0];
      }
    }

    return { roles, primaryRole };
  }

  private hasMatchingUser(users: IPPOApproverUser[] | null, targetEmail: string): boolean {
    if (!users || users.length === 0 || !targetEmail) {
      return false;
    }

    for (let i = 0; i < users.length; i++) {
      const candidate = users[i];
      if (!candidate || !candidate.email) {
        continue;
      }

      if (candidate.email.toLowerCase() === targetEmail) {
        return true;
      }
    }

    return false;
  }

  private pickBestMatch(items: any[], candidates: string[]): any {
    if (!items || items.length === 0) {
      return null;
    }

    if (items.length === 1 || candidates.length === 0) {
      return items[0];
    }

    const normalizedCandidates: string[] = [];
    for (let i = 0; i < candidates.length; i++) {
      normalizedCandidates.push(candidates[i].toLowerCase());
    }

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const listValue = (item.InternalProjectName || '').toLowerCase();
      if (normalizedCandidates.indexOf(listValue) !== -1) {
        return item;
      }
    }

    return items[0];
  }

  private buildCandidateProjectNames(context: WebPartContext): string[] {
    const candidates: string[] = [];

    const webTitle = (context.pageContext && context.pageContext.web && context.pageContext.web.title)
      ? context.pageContext.web.title.trim()
      : '';
    this.addUniqueCandidate(webTitle, candidates);

    const webServerRelative = context.pageContext && context.pageContext.web
      ? (context.pageContext.web.serverRelativeUrl || '')
      : '';
    const webAbsolute = context.pageContext && context.pageContext.web
      ? (context.pageContext.web.absoluteUrl || '')
      : '';
    const siteAbsolute = context.pageContext && context.pageContext.site
      ? (context.pageContext.site.absoluteUrl || '')
      : '';

    this.addPathSegments(webServerRelative, candidates);
    this.addPathSegments(webAbsolute, candidates);
    this.addPathSegments(siteAbsolute, candidates);

    const filtered: string[] = [];
    for (let i = 0; i < candidates.length; i++) {
      const value = candidates[i];
      if (value && filtered.indexOf(value) === -1) {
        filtered.push(value);
      }
    }

    return filtered;
  }

  private addPathSegments(path: string, candidates: string[]): void {
    if (!path) {
      return;
    }

    const segments = path.split('/');
    const trimmed: string[] = [];

    for (let i = 0; i < segments.length; i++) {
      const part = decodeURIComponent(segments[i].trim());
      if (part.length > 0) {
        trimmed.push(part);
      }
    }

    if (trimmed.length > 0) {
      const last = trimmed[trimmed.length - 1];
      this.addUniqueCandidate(last, candidates);
    }
  }

  private addUniqueCandidate(value: string, candidates: string[]): void {
    if (!value) {
      return;
    }

    if (candidates.indexOf(value) === -1) {
      candidates.push(value);
    }
  }

  private escapeODataValue(value: string): string {
    return value.replace(/'/g, "''");
  }
}
