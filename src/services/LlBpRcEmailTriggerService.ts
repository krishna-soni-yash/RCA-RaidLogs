import { WebPartContext } from '@microsoft/sp-webpart-base';
import ParentListNames from '../common/Constants';
import ErrorMessages from '../common/ErrorMessages';
import GenericServiceInstance from './GenericServices';
import { IGenericService } from './IGenericServices';
import { ILessonsLearnt } from '../models/Ll Bp Rc/LessonsLearnt';
import { IBestPractices } from '../models/Ll Bp Rc/BestPractices';
import { IReusableComponents } from '../models/Ll Bp Rc/ReusableComponents';

/* eslint-disable @typescript-eslint/no-explicit-any */

export interface ILLBPrcEmailTrigger {
  Id?: number;
  Title: string;
  To?: string;
  CC?: string;
  ItemLink?: string;
  SubSiteURL?: string;
  Description?: string;
  Category?: string;
  Remarks?: string;
}

export class LlBpRcEmailTriggerService {
  private genericService: IGenericService;
  private context: WebPartContext;

  constructor(context: WebPartContext, service?: IGenericService) {
    this.context = context;
    this.genericService = service ?? GenericServiceInstance;
    this.genericService.init(undefined, context);
  }

  public async createEmailTrigger(item: ILessonsLearnt | IBestPractices | IReusableComponents | any, dataType?: string): Promise<boolean> {
    if (!this.context) {
      console.error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_PPOAPPROVERS);
      return false;
    }

    try {
      const emailTriggerItem = await this.mapItemToEmailTrigger(item, dataType);

      const result = await this.genericService.saveItem({
        context: this.context,
        listTitle: ParentListNames.LLBPRCEmailTrigger,
        item: emailTriggerItem,
        forceSiteUrl: this.context.pageContext.site.absoluteUrl
      });

      if (result.success) {
        console.log('LLBPrc email trigger created successfully for item:', item?.ID || item?.Id);
        return true;
      } else {
        console.error('Failed to create LLBPRC email trigger:', result.error);
        return false;
      }
    } catch (error: any) {
      console.error('Error creating LLBPRC email trigger:', error);
      return false;
    }
  }

  private async mapItemToEmailTrigger(item: any, dataType?: string): Promise<ILLBPrcEmailTrigger> {
    const subsiteUrl = this.context.pageContext.web.absoluteUrl;
    const itemLink = this.generateItemLink(item);

    // Determine description and category based on data type or available fields
    let description = '';
    let category = '';
    let remarks = '';

    const type = (dataType || item?.DataType || '').toString().trim();

    if (type === 'LessonsLearnt' || item?.LlProblemFacedLearning || item?.LlSolution) {
      description = item?.LlProblemFacedLearning ?? item?.LlSolution ?? '';
      category = item?.LlCategory ?? '';
      remarks = item?.LlRemarks ?? '';
    } else if (type === 'BestPractices' || item?.BpBestPracticesDescription) {
      description = item?.BpBestPracticesDescription ?? '';
      category = item?.BpCategory ?? '';
      remarks = item?.BpRemarks ?? '';
    } else if (type === 'ReusableComponents' || item?.RcComponentName) {
      description = item?.RcComponentName ?? '';
      category = item?.RcLocation ?? '';
      remarks = item?.RcRemarks ?? '';
    } else {
      // Generic fallback
      description = item?.Title ?? '';
      remarks = item?.Remarks ?? '';
    }

    const subject = this.generateEmailSubject(type, description);

    const payload: ILLBPrcEmailTrigger = {
      Title: subject,
      ItemLink: itemLink,
      SubSiteURL: subsiteUrl,
      Description: description ?? '',
      Category: category ?? '',
      Remarks: remarks ?? ''
    };

    return payload;
  }

  private generateEmailSubject(type: string, description: string): string {
    const projectName = this.extractProjectName();
    const truncated = (description || '').length > 50 ? (description || '').substring(0, 47) + '...' : (description || 'New Item');
    const readableType = type || 'LLBPrc Item';
    return `New ${readableType} Created - ${projectName}: ${truncated}`;
  }

  private generateItemLink(item: any): string {
    const origin = window.location.origin;
    const pathname = window.location.pathname;
    const baseUrl = `${origin}${pathname}`;
    if (item?.ID || item?.Id) {
      const id = item?.ID ?? item?.Id;
      return `${baseUrl}?LlBpRcId=${id}`;
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
}

export default LlBpRcEmailTriggerService;
