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
        currentUserRole: Current_User_Role.None
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
          'BUH/Id',
          'BUH/ID',
          'BUH/Title',
          'BUH/EMail',
          'ProjectManager/Id',
          'ProjectManager/ID',
          'ProjectManager/Title',
          'ProjectManager/EMail'
        ],
        expand: ['Reviewer', 'BUH', 'ProjectManager'],
        filter,
        pageSize: 5
      });

      if (!items || items.length === 0) {
        return {
          approver: null,
          currentUserRole: Current_User_Role.None
        };
      }

      const matchedItem = this.pickBestMatch(items, candidateProjectNames);
      const approver = this.mapApprover(matchedItem);
      const currentUserRole = this.resolveCurrentUserRole(context, approver);

      return { approver, currentUserRole };
    } catch (error: any) {
      console.error(ErrorMessages.FAILED_TO_FETCH_PPOAPPROVERS, error);
      return {
        approver: null,
        currentUserRole: Current_User_Role.None
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
      Reviewer: this.mapUser(item.Reviewer),
      BUH: this.mapUser(item.BUH),
      ProjectManager: this.mapUser(item.ProjectManager)
    };

    return approver;
  }

  private mapUser(user: any): IPPOApproverUser | null {
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

  private resolveCurrentUserRole(context: WebPartContext, approver: IPPOApprovers | null): Current_User_Role {
    if (!approver) {
      return Current_User_Role.None;
    }

    const currentUserEmail = (context.pageContext?.user?.email || '').trim().toLowerCase();
    if (!currentUserEmail) {
      return Current_User_Role.None;
    }

    const buhEmail = approver.BUH && approver.BUH.email
      ? approver.BUH.email.toLowerCase()
      : '';
    if (buhEmail === currentUserEmail) {
      return Current_User_Role.BUH;
    }

    const projectManagerEmail = approver.ProjectManager && approver.ProjectManager.email
      ? approver.ProjectManager.email.toLowerCase()
      : '';
    if (projectManagerEmail === currentUserEmail) {
      return Current_User_Role.ProjectManager;
    }

    const reviewerEmail = approver.Reviewer && approver.Reviewer.email
      ? approver.Reviewer.email.toLowerCase()
      : '';
    if (reviewerEmail === currentUserEmail) {
      return Current_User_Role.Reviewer;
    }

    return Current_User_Role.None;
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
