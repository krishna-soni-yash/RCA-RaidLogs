import { WebPartContext } from '@microsoft/sp-webpart-base';
import ParentListNames from '../common/Constants';
import ErrorMessages from '../common/ErrorMessages';
import GenericServiceInstance from './GenericServices';
import { IGenericService } from './IGenericServices';
import { IRaidItem } from '../webparts/rootCauseAnalysis/components/RaidLogs/interfaces/IRaidItem';

/* eslint-disable @typescript-eslint/no-explicit-any */

/**
 * Interface for RAIDLogEmailTrigger list item
 */
export interface IRaidLogEmailTrigger {
  Id?: number;
  Title: string;  // Subject of email
  To?: string;
  CC?: string;
  ItemLink?: string;
  SubSiteURL?: string;
  Remarks?: string;
  Priority?: string;
  IdentificationDate?: string;
  Description?: string;
  Status?: string;
}

/**
 * Service class for creating email trigger entries when RAID items are created
 */
export class RaidLogEmailTriggerService {
  private genericService: IGenericService;
  private context: WebPartContext;

  constructor(context: WebPartContext, service?: IGenericService) {
    this.context = context;
    this.genericService = service ?? GenericServiceInstance;
    this.genericService.init(undefined, context);
  }

  /**
   * Creates an email trigger entry in the RAIDLogEmailTrigger list
   * @param raidItem The RAID item that was created
   * @returns Promise<boolean> True if successful, false otherwise
   */
  public async createEmailTrigger(raidItem: IRaidItem): Promise<boolean> {
    if (!this.context) {
      console.error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_PPOAPPROVERS);
      return false;
    }

    try {
      const emailTriggerItem = await this.mapRaidItemToEmailTrigger(raidItem);

      const result = await this.genericService.saveItem({
        context: this.context,
        listTitle: ParentListNames.RAIDLogEmailTrigger,
        item: emailTriggerItem,
        forceSiteUrl: this.context.pageContext.site.absoluteUrl // Use root site
      });

      if (result.success) {
        console.log('Email trigger created successfully for RAID item:', raidItem.id);
        return true;
      } else {
        console.error('Failed to create email trigger:', result.error);
        return false;
      }
    } catch (error: any) {
      console.error('Error creating email trigger:', error);
      return false;
    }
  }

  /**
   * Maps a RAID item to an email trigger list item
   * @param raidItem The RAID item to map
   * @returns IRaidLogEmailTrigger The mapped email trigger item
   */
  private async mapRaidItemToEmailTrigger(raidItem: IRaidItem): Promise<IRaidLogEmailTrigger> {
    const subsiteUrl = this.context.pageContext.web.absoluteUrl;
    const itemLink = this.generateItemLink(raidItem);
    const emailSubject = this.generateEmailSubject(raidItem);

    const emailTrigger: IRaidLogEmailTrigger = {
      Title: emailSubject,
      ItemLink: itemLink,
      SubSiteURL: subsiteUrl,
      Priority: raidItem.priority || '',
      IdentificationDate: raidItem.identificationDate || raidItem.date || '',
      Description: raidItem.description || raidItem.details || '',
      Status: raidItem.status || 'New',
      Remarks: raidItem.remarks || ''
    };

    // Keep human-readable emails too (for display), but resolve users to SharePoint user IDs
    // so Person fields are populated correctly. We will set ToId/CCId when IDs are available.
    let toEmails = '';
    let ccEmails = '';
    if (raidItem.responsibility && raidItem.responsibility.length > 0) {
      toEmails = this.getUserEmails(raidItem.responsibility);
    } else if (raidItem.byWhom && raidItem.byWhom.length > 0) {
      toEmails = this.getUserEmails(raidItem.byWhom);
    }

    if (raidItem.byWhom && raidItem.byWhom.length > 0) {
      ccEmails = this.getUserEmails(raidItem.byWhom);
    }

    if (toEmails) emailTrigger.To = toEmails;
    if (ccEmails) emailTrigger.CC = ccEmails;

    // To/CC are single-line text fields in the list — store the first email only.
    if (emailTrigger.To) {
      const first = emailTrigger.To.split(/;|,/)[0].trim();
      emailTrigger.To = first;
    }
    if (emailTrigger.CC) {
      const firstC = emailTrigger.CC.split(/;|,/)[0].trim();
      emailTrigger.CC = firstC;
    }

    // Debug logs to verify people-picker mapping
    try {
      console.log('RaidLogEmailTrigger mapping - responsibility:', raidItem.responsibility);
      console.log('RaidLogEmailTrigger mapping - toEmails (before trim):', toEmails, 'final To:', emailTrigger.To, 'final CC:', emailTrigger.CC);
    } catch {
      // ignore logging failures
    }

    // Fallback: if To was not extracted, try parsing raw responsibility value(s) for "id|email" patterns
    if (!emailTrigger.To && raidItem && raidItem.responsibility) {
      try {
        const raw: any = raidItem.responsibility;
        const candidates: string[] = [];

        if (typeof raw === 'string') {
          // may be "id|email; id|email"
          const parts = String(raw).split(/;|,/).map((p: string) => p.trim()).filter(Boolean);
          for (let i = 0; i < parts.length; i++) {
            const seg = String(parts[i]).split('|').map((s: string) => s.trim());
            if (seg.length > 1 && this.isEmail(seg[1])) {
              candidates.push(seg[1]);
            }
          }
        } else if (Array.isArray(raw)) {
          for (let i = 0; i < raw.length; i++) {
            const el: any = raw[i];
            if (!el) continue;
            if (typeof el === 'string') {
              const seg = String(el).split('|').map((s: string) => s.trim());
              if (seg.length > 1 && this.isEmail(seg[1])) candidates.push(seg[1]);
            } else if (typeof el === 'object') {
              // sometimes object holds the original serialized string in a property
              const possible: any = (el as any).email || (el as any).Email || (el as any).loginName || (el as any).LoginName || (el as any).id;
              if (possible && typeof possible === 'string') {
                const seg = String(possible).split('|').map((s: string) => s.trim());
                if (seg.length > 1 && this.isEmail(seg[1])) candidates.push(seg[1]);
              }
            }
          }
        } else if (typeof raw === 'object') {
          const possible: any = (raw as any).email || (raw as any).Email || (raw as any).loginName || (raw as any).LoginName || (raw as any).id;
          if (possible && typeof possible === 'string') {
            const seg = String(possible).split('|').map((s: string) => s.trim());
            if (seg.length > 1 && this.isEmail(seg[1])) candidates.push(seg[1]);
          }
        }

        if (candidates.length > 0) {
          emailTrigger.To = candidates[0];
        }
      } catch {
        // ignore fallback parsing errors
      }
    }

    // Final debug: show the exact payload that will be saved to the RAIDLogEmailTrigger list
    try {
      console.log('RaidLogEmailTrigger final payload:', emailTrigger);
    } catch {
      // ignore
    }

    return emailTrigger;
  }

  /**
   * Generates email subject based on RAID type and item details
   * @param raidItem The RAID item
   * @returns string The email subject
   */
  private generateEmailSubject(raidItem: IRaidItem): string {
    const projectName = this.extractProjectName();
    const raidType = raidItem.type;
    const description = raidItem.description || raidItem.details || 'New Item';
    
    // Truncate description if too long
    const truncatedDesc = description.length > 50 
      ? description.substring(0, 47) + '...' 
      : description;

    return `New ${raidType} Created - ${projectName}: ${truncatedDesc}`;
  }

  /**
   * Generates a link to the RAID item in SharePoint
   * @param raidItem The RAID item
   * @returns string The item link
   */
  private generateItemLink(raidItem: IRaidItem): string {
    const origin = window.location.origin;
    const pathname = window.location.pathname;
    const baseUrl = `${origin}${pathname}`;
    
    // For Risk items, use RAIDId query parameter
    if (raidItem.type === 'Risk' && raidItem.raidId) {
      return `${baseUrl}?RAIDId=${raidItem.raidId}`;
    }
    
    // For other types (Opportunity, Issue, Assumption, Dependency, Constraints), use worklogId
    if (raidItem.id) {
      return `${baseUrl}?RaidlogId=${raidItem.id}`;
    }
    
    return baseUrl;
  }

  /**
   * Extracts project name from the subsite URL
   * @returns string The project name
   */
  private extractProjectName(): string {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const siteUrl = this.context.pageContext.site.absoluteUrl;
    
    // Extract the relative URL part
    const relativePath = webUrl.replace(siteUrl, '');
    
    if (relativePath) {
      // Get the last segment of the path as project name
      const segments = relativePath.split('/').filter(s => s);
      return segments.length > 0 ? segments[segments.length - 1] : 'Project';
    }
    
    return 'Project';
  }

  /**
   * Extracts email addresses from people-picker values.
   * Supports:
   * - Expanded user objects: { email | Email | EMail }
   * - LoginName values when they contain an email
   * - String formats like "id|email; id|email"
   * - Arrays or single string values
   * @param users any people-picker value
   * @returns string Semicolon-separated unique email addresses
   */
  private getUserEmails(users: any): string {
    if (!users) return '';

    const emails: string[] = [];

    const pushIfEmail = (val: any) => {
      if (!val) return;
      const s = String(val).trim();
      if (this.isEmail(s)) {
        emails.push(s);
      }
    };

    if (Array.isArray(users)) {
      users.forEach(u => {
        if (!u) return;
        if (typeof u === 'string') {
          // possible formats: "id|email" or multiple separated by ; or ,
          const parts = u.split(/;|,/).map(p => p.trim()).filter(Boolean);
          parts.forEach(p => {
            const seg = p.split('|').map(s => s.trim());
            if (seg.length > 1) pushIfEmail(seg[1]);
            else pushIfEmail(seg[0]);
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
        if (seg.length > 1) pushIfEmail(seg[1]);
        else pushIfEmail(seg[0]);
      });
    } else if (typeof users === 'object') {
      pushIfEmail(users.email || users.Email || users.EMail || users.LoginName || users.loginName || users.Login);
    } else {
      pushIfEmail(users);
    }

    // Deduplicate and return (avoid Array.from for older lib targets)
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

  /**
   * Simple email format check
   */
  private isEmail(value: string): boolean {
    if (!value) return false;
    // basic regex for email validation
    // allow loginName in form i:0#.f|membership|user@domain.com as well
    const emailMatch = value.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
    return !!emailMatch;
  }

  // resolveUserIds removed — To/CC are single-line text fields; no user-id resolution required

  // Removed unused helper escapeODataValue to avoid unused symbol errors
}

export default RaidLogEmailTriggerService;
