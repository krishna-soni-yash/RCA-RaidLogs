import { WebPartContext } from '@microsoft/sp-webpart-base';
import ParentListNames from '../common/Constants';
import ErrorMessages from '../common/ErrorMessages';
import GenericServiceInstance from './GenericServices';
import { IGenericService } from './IGenericServices';
import { IRCAList } from '../models/IRCAList';

/* eslint-disable @typescript-eslint/no-explicit-any */

export interface IRCAEmailTrigger {
  Id?: number;
  Title: string;
  To?: string;
  CC?: string;
  ItemLink?: string;
  SubSiteURL?: string;
  Remarks?: string;
  Description?: string;
  Priority?: string;
}

export class RCAEmailTriggerService {
  private genericService: IGenericService;
  private context: WebPartContext;

  constructor(context: WebPartContext, service?: IGenericService) {
    this.context = context;
    this.genericService = service ?? GenericServiceInstance;
    this.genericService.init(undefined, context);
  }

  public async createEmailTrigger(rcaItem: IRCAList | any): Promise<boolean> {
    if (!this.context) {
      console.error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_PPOAPPROVERS);
      return false;
    }

    try {
      const emailTriggerItems = await this.mapRcaToEmailTriggers(rcaItem);

      // Save one entry per email trigger payload. Best-effort: continue on errors.
      let allOk = true;
      for (let i = 0; i < emailTriggerItems.length; i++) {
        try {
          const payload = emailTriggerItems[i];
          const res = await this.genericService.saveItem({
            context: this.context,
            listTitle: ParentListNames.RCAEmailTrigger,
            item: payload,
            forceSiteUrl: this.context.pageContext.site.absoluteUrl
          });
          if (!res || !res.success) {
            allOk = false;
            console.error('Failed to create RCA email trigger for payload:', payload, res && res.error);
          }
        } catch (err) {
          allOk = false;
          console.error('Error creating RCA email trigger for one of the payloads:', err);
        }
      }

      if (allOk) console.log('RCA email trigger(s) created successfully for item:', rcaItem?.ID || rcaItem?.Id);
      return allOk;
    } catch (error: any) {
      console.error('Error creating RCA email trigger:', error);
      return false;
    }
  }

  private async mapRcaToEmailTriggers(rcaItem: IRCAList | any): Promise<IRCAEmailTrigger[]> {
    const subsiteUrl = this.context.pageContext.web.absoluteUrl;
    const itemLink = this.generateItemLink(rcaItem);

    const description = rcaItem?.ProblemStatement || rcaItem?.Title || rcaItem?.LinkTitle || '';
    const remarks = rcaItem?.Remarks || '';
    const priority = rcaItem?.RCAPriority || '';

    const subject = this.generateEmailSubject(description);

    // Helper: get array of emails from a responsibility field
    const emailsFromField = (fieldVal: any): string[] => {
      const list = this.getUserEmails(fieldVal);
      if (!list) return [];
      return list.split(/;|,/).map((s: string) => s.trim()).filter(Boolean);
    };

    const a = emailsFromField(rcaItem?.ResponsibilityCorrection);
    const b = emailsFromField(rcaItem?.ResponsibilityCorrective);
    const c = emailsFromField(rcaItem?.ResponsibilityPreventive);

    

    const arrayEqualsAsSet = (x: string[], y: string[]) => {
      if (!x || !y) return false;
      if (x.length !== y.length) return false;
      const sx = x.slice().map(s => s.toLowerCase()).sort();
      const sy = y.slice().map(s => s.toLowerCase()).sort();
      for (let i = 0; i < sx.length; i++) if (sx[i] !== sy[i]) return false;
      return true;
    };

    const payloads: IRCAEmailTrigger[] = [];

    // If all three responsibility fields are present and contain the same users, create a single entry
    if (a.length > 0 && b.length > 0 && c.length > 0 && arrayEqualsAsSet(a, b) && arrayEqualsAsSet(b, c)) {
      payloads.push({
        Title: subject,
        ItemLink: itemLink,
        SubSiteURL: subsiteUrl,
        Description: description ?? '',
        Remarks: remarks ?? '',
        Priority: priority ?? '',
        To: a.join('; ')
      });
      return payloads;
    }

    // Otherwise, create one entry per unique email across fields
    const emailsSet: { [key: string]: boolean } = {};
    [a, b, c].forEach(arr => {
      (arr || []).forEach(e => { emailsSet[e] = true; });
    });

    const uniqueEmails = Object.keys(emailsSet);

    if (uniqueEmails.length === 0) {
      // no responsible users: create single generic entry without To
      payloads.push({
        Title: subject,
        ItemLink: itemLink,
        SubSiteURL: subsiteUrl,
        Description: description ?? '',
        Remarks: remarks ?? '',
        Priority: priority ?? ''
      });
    } else {
      uniqueEmails.forEach(email => {
        payloads.push({
          Title: subject,
          ItemLink: itemLink,
          SubSiteURL: subsiteUrl,
          Description: description ?? '',
          Remarks: remarks ?? '',
          Priority: priority ?? '',
          To: email
        });
      });
    }

    return payloads;
  }

  private generateEmailSubject(description: string): string {
    const projectName = this.extractProjectName();
    const truncated = (description || '').length > 50 ? (description || '').substring(0, 47) + '...' : (description || 'New Item');
    return `New RCA Created - ${projectName}: ${truncated}`;
  }

  private generateItemLink(item: any): string {
    const origin = window.location.origin;
    const pathname = window.location.pathname;
    const baseUrl = `${origin}${pathname}`;
    const id = item?.ID ?? item?.Id ?? item?.itemId ?? item?.ItemId;
    if (id) {
      return `${baseUrl}?RCAId=${id}`;
    }
    return baseUrl;
  }

  private extractProjectName(): string {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const siteUrl = this.context.pageContext.site.absoluteUrl;
    const relativePath = webUrl.replace(siteUrl, '');
    if (relativePath) {
      const segments = relativePath.split('/').filter(s => s);
      return segments.length > 0 ? segments[segments.length - 1] : 'Project';
    }
    return 'Project';
  }

  private getUserEmails(users: any): string {
    if (!users) return '';

    const emails: string[] = [];

    const pushIfEmail = (val: any) => {
      if (!val) return;
      const s = String(val).trim();
      const m = s.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
      if (m) emails.push(m[0]);
    };

    if (Array.isArray(users)) {
      users.forEach(u => {
        if (!u) return;
        if (typeof u === 'string') {
          const parts = u.split(/;|,/).map(p => p.trim()).filter(Boolean);
          parts.forEach(p => {
            const seg = p.split('|').map(s => s.trim());
            if (seg.length > 1) pushIfEmail(seg[1]); else pushIfEmail(seg[0]);
          });
        } else if (typeof u === 'object') {
          pushIfEmail(u.email || u.Email || u.EMail || u.LoginName || u.loginName || u.Login);
        } else {
          pushIfEmail(u);
        }
      });
    } else if (typeof users === 'string') {
      const parts = users.split(/;|,/).map(p => p.trim()).filter(Boolean);
      parts.forEach(p => {
        const seg = p.split('|').map(s => s.trim());
        if (seg.length > 1) pushIfEmail(seg[1]); else pushIfEmail(seg[0]);
      });
    } else if (typeof users === 'object') {
      pushIfEmail(users.email || users.Email || users.EMail || users.LoginName || users.loginName || users.Login);
    } else {
      pushIfEmail(users);
    }

    const seen: { [key: string]: boolean } = {};
    const unique: string[] = [];
    for (let i = 0; i < emails.length; i++) {
      const e = emails[i];
      if (!seen[e]) {
        seen[e] = true;
        unique.push(e);
      }
    }

    return unique.join('; ');
  }
}

export default RCAEmailTriggerService;
