import genericService, { GenericService } from '../../services/GenericServices';
import IGenericService from '../../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRCAList } from '../../models/IRCAList';
import ErrorMessages from '../../common/ErrorMessages';
import { SubSiteListNames, selectedFields, expandFields } from '../../common/Constants';
import IRCARepository from '../IRCARepository';
//import IObjectivesMasterRepository from './repositoriesInterface/IObjectivesMasterRepository';

/**
 * Repository for ProjectTypes list
 * Implements a simple cached fetch of Id/LinkTitle values
 */
export class RCARepository implements IRCARepository {
    private service: IGenericService;
    private cache: IRCAList[] | null = null;
    private cacheTimestamp = 0;
    private readonly CACHE_DURATION = 5 * 60 * 1000; // 5 minutes

    constructor(service?: IGenericService) {
        this.service = service ?? genericService;

    }

    public setService(service: IGenericService): void {
        this.service = service;
    }
    // private normalizeSiteUrl(value?: string): string {
    //     return (value || '').trim().replace(/\/+$/, '').toLowerCase();
    //   }
    public async getRCAItems(useCache: boolean = true, context?: WebPartContext): Promise<IRCAList[]> {
        const now = Date.now();

        if (useCache && this.cache && (now - this.cacheTimestamp) < this.CACHE_DURATION) {
            return this.cache;
        }

        if (!context) {
            throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
        }

        try {
            const genericServiceInstance: IGenericService = new GenericService(undefined, context);
            genericServiceInstance.init(undefined, context);
            // const selectFields: string[] = ['Id', 'LinkTitle','ProjectType','IsActive'];

            const items = await this.service.fetchAllItems<any>({
                context,
                listTitle: SubSiteListNames.RootCauseAnalysis,
                select: selectedFields,
                pageSize: 2000,
                expand: expandFields
                //filter: 'IsActive eq 1 and ProjectType/Title eq \'' + (selectedProjectType) + '\'',
                // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

            });

            const normalized = (items || []).map((it: any) => ({
                ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
                LinkTitle: it?.Title || it?.LinkTitle || '',
                ProblemStatementNumber: it?.ProblemStatementNumber || '',
                ProblemStatement: it?.ProblemStatement || '',
                CauseCategory: it?.CauseCategory || '',
                RCASource: it?.RCASource || '',
                RCAPriority: it?.RCAPriority || '',
                RelatedMetric: it?.RelatedMetric || '',
                Cause: it?.Cause || '',
                RootCause: it?.RootCause || '',
                RCATechniqueUsedAndReference: it?.RCATechniqueUsedAndReference || '',
                RCATypeOfAction: it?.RCATypeOfAction || '',

                ActionPlanCorrection: it?.ActionPlanCorrection || '',
                ResponsibilityCorrection: it?.ResponsibilityCorrection?.EMail.toString() || '',
                PlannedClosureDateCorrection: it?.PlannedClosureDateCorrection || '',
                ActualClosureDateCorrection: it?.ActualClosureDateCorrection || '',

                ActionPlanCorrective: it?.ActionPlanCorrective || '',
                ResponsibilityCorrective: it?.ResponsibilityCorrective?.EMail.toString() || '',
                PlannedClosureDateCorrective: it?.PlannedClosureDateCorrective || '',
                ActualClosureDateCorrective: it?.ActualClosureDateCorrective || '',

                ActionPlanPreventive: it?.ActionPlanPreventive || '',
                ResponsibilityPreventive: it?.ResponsibilityPreventive?.EMail.toString() || '',
                PlannedClosureDatePreventive: it?.PlannedClosureDatePreventive || '',
                ActualClosureDatePreventive: it?.ActualClosureDatePreventive || '',

                PerformanceBeforeActionPlan: it?.PerformanceBeforeActionPlan || '',
                PerformanceAfterActionPlan: it?.PerformanceAfterActionPlan || '',
                QuantitativeOrStatisticalEffecti: it?.QuantitativeOrStatisticalEffecti || '',
                Remarks: it?.Remarks || '',

            })) as unknown as IRCAList[];

            this.cache = normalized;
            this.cacheTimestamp = now;

            return this.cache;
        } catch (error: any) {
            throw new Error('Failed to fetch ProjectType: ' + (error?.message || error));
        }
    }

    public resolveEmailsToIds = async (raw: any, context?: WebPartContext): Promise<number[]> => {
        const emails: string[] = typeof raw === 'string'
            ? raw.split(/; ?/).map((s: string) => s.trim()).filter(Boolean)
            : (Array.isArray(raw) ? raw.map(String).map(s => s.trim()).filter(Boolean) : []);

        if (emails.length === 0) return [];

        // prefer common ensureUser helpers if implemented on the service
        const svcAny = this.service as any;
        const ensureFn = svcAny?.ensureuser
        if (!ensureFn) {
            // no resolver available; avoid sending person objects that may break REST — return empty
            console.warn('RCARepository.saveRCAItem: no ensureUser function on service; skipping person resolution for', emails);
            return [];
        }

        const ids: number[] = [];
        for (const em of emails) {
            try {
                // some ensureUser implementations accept email/login and return user object or id
                const res = await ensureFn(context,em);
                const id = typeof res === 'number' ? res : (res?.Id || res?.ID || res?.id);
                if (id) ids.push(Number(id));
            } catch (e) {
                console.warn('RCARepository.saveRCAItem: ensureUser failed for', em, e);
            }
        }
        return ids;
    };

    public async saveRCAItem(item: IRCAList, context?: WebPartContext): Promise<any> {
        if (!context) {
            throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
        }
        try {
            const payload: Record<string, unknown> = {};
            if (item.LinkTitle !== undefined) payload.Title = item.LinkTitle;
            if (item.ProblemStatementNumber !== undefined) payload.ProblemStatementNumber = item.ProblemStatementNumber;
            if (item.CauseCategory !== undefined) payload.CauseCategory = item.CauseCategory;
            if (item.RCASource !== undefined) payload.RCASource = item.RCASource;
            if (item.RCAPriority !== undefined) payload.RCAPriority = item.RCAPriority;
            if (item.RelatedMetric !== undefined) payload.RelatedMetric = item.RelatedMetric;
            if (item.Cause !== undefined) payload.Cause = item.Cause;
            if (item.RootCause !== undefined) payload.RootCause = item.RootCause;
            if (item.RCATechniqueUsedAndReference !== undefined) payload.RCATechniqueUsedAndReference = item.RCATechniqueUsedAndReference;
            if (item.RCATypeOfAction !== undefined) payload.RCATypeOfAction = item.RCATypeOfAction;

            // helper: try to resolve emails (or simple strings) to SharePoint user IDs using service helpers


            if (item.ActionPlanCorrection !== undefined) payload.ActionPlanCorrection = item.ActionPlanCorrection;
            if (item.ResponsibilityCorrection !== undefined) {
                const ids = await this.resolveEmailsToIds(item.ResponsibilityCorrection,context);
                if (ids.length === 1) payload.ResponsibilityCorrectionId = ids[0];
                else if (ids.length > 1) payload.ResponsibilityCorrectionId = { results: ids };
                else {
                    // no IDs resolved -> skip adding ResponsibilityCorrection to payload to avoid malformed request
                    console.warn('RCARepository.saveRCAItem: no user IDs resolved for ResponsibilityCorrection; field omitted from payload');
                }
            }
            if (item.PlannedClosureDateCorrection !== undefined) payload.PlannedClosureDateCorrection = item.PlannedClosureDateCorrection;
            if (item.ActualClosureDateCorrection !== undefined) payload.ActualClosureDateCorrection = item.ActualClosureDateCorrection;

            if (item.ActionPlanCorrective !== undefined) payload.ActionPlanCorrective = item.ActionPlanCorrective;
            if (item.ResponsibilityCorrective !== undefined) {
                const ids = await this.resolveEmailsToIds(item.ResponsibilityCorrective,context);
                if (ids.length === 1) payload.ResponsibilityCorrectiveId = ids[0];
                else if (ids.length > 1) payload.ResponsibilityCorrectiveId = { results: ids };
                else {
                    console.warn('RCARepository.saveRCAItem: no user IDs resolved for ResponsibilityCorrective; field omitted from payload');
                }
            }

            if (item.PlannedClosureDateCorrective !== undefined) payload.PlannedClosureDateCorrective = item.PlannedClosureDateCorrective;
            if (item.ActualClosureDateCorrective !== undefined) payload.ActualClosureDateCorrective = item.ActualClosureDateCorrective;
            if (item.ActionPlanPreventive !== undefined) payload.ActionPlanPreventive = item.ActionPlanPreventive;
            if (item.ResponsibilityPreventive !== undefined) {
                const ids = await this.resolveEmailsToIds(item.ResponsibilityPreventive,context);
                if (ids.length === 1) payload.ResponsibilityPreventiveId = ids[0];
                else if (ids.length > 1) payload.ResponsibilityPreventiveId = { results: ids };
                else {
                    console.warn('RCARepository.saveRCAItem: no user IDs resolved for ResponsibilityPreventive; field omitted from payload');
                }
            }
            if (item.PlannedClosureDatePreventive !== undefined) payload.PlannedClosureDatePreventive = item.PlannedClosureDatePreventive;
            if (item.ActualClosureDatePreventive !== undefined) payload.ActualClosureDatePreventive = item.ActualClosureDatePreventive;

            const result = await this.service.saveItem<IRCAList>({
                context,
                listTitle: SubSiteListNames.RootCauseAnalysis,
                item: payload,
                select: selectedFields,
                expand: expandFields
            });
            this.refresh();

            return result;
        } catch (error: any) {
            throw new Error('Failed to save RCA item: ' + (error?.message || error));
        }
    }

    public async updateRCAItem(itemId: number, item: IRCAList, context?: WebPartContext): Promise<any> {
        if (!context) {
            throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
        }
        try {
            const payload: Record<string, unknown> = {};
            if (item.LinkTitle !== undefined) payload.Title = item.LinkTitle;
            if (item.ProblemStatementNumber !== undefined) payload.ProblemStatementNumber = item.ProblemStatementNumber;
            if (item.CauseCategory !== undefined) payload.CauseCategory = item.CauseCategory;
            if (item.RCASource !== undefined) payload.RCASource = item.RCASource;
            if (item.RCAPriority !== undefined) payload.RCAPriority = item.RCAPriority;
            if (item.RelatedMetric !== undefined) payload.RelatedMetric = item.RelatedMetric;
            if (item.Cause !== undefined) payload.Cause = item.Cause;
            if (item.RootCause !== undefined) payload.RootCause = item.RootCause;
            if (item.RCATechniqueUsedAndReference !== undefined) payload.RCATechniqueUsedAndReference = item.RCATechniqueUsedAndReference;
            if (item.RCATypeOfAction !== undefined) payload.RCATypeOfAction = item.RCATypeOfAction;

            // helper: try to resolve emails (or simple strings) to SharePoint user IDs using service helpers


            if (item.ActionPlanCorrection !== undefined) payload.ActionPlanCorrection = item.ActionPlanCorrection;
            if (item.ResponsibilityCorrection !== undefined) {
                const ids = await this.resolveEmailsToIds(item.ResponsibilityCorrection,context);
                if (ids.length === 1) payload.ResponsibilityCorrectionId = ids[0];
                else if (ids.length > 1) payload.ResponsibilityCorrectionId = { results: ids };
                else {
                    // no IDs resolved -> skip adding ResponsibilityCorrection to payload to avoid malformed request
                    console.warn('RCARepository.saveRCAItem: no user IDs resolved for ResponsibilityCorrection; field omitted from payload');
                }
            }
            if (item.PlannedClosureDateCorrection !== undefined) payload.PlannedClosureDateCorrection = item.PlannedClosureDateCorrection;
            if (item.ActualClosureDateCorrection !== undefined) payload.ActualClosureDateCorrection = item.ActualClosureDateCorrection;

            if (item.ActionPlanCorrective !== undefined) payload.ActionPlanCorrective = item.ActionPlanCorrective;
            if (item.ResponsibilityCorrective !== undefined) {
                const ids = await this.resolveEmailsToIds(item.ResponsibilityCorrective,context);
                if (ids.length === 1) payload.ResponsibilityCorrectiveId = ids[0];
                else if (ids.length > 1) payload.ResponsibilityCorrectiveId = { results: ids };
                else {
                    console.warn('RCARepository.saveRCAItem: no user IDs resolved for ResponsibilityCorrective; field omitted from payload');
                }
            }

            if (item.PlannedClosureDateCorrective !== undefined) payload.PlannedClosureDateCorrective = item.PlannedClosureDateCorrective;
            if (item.ActualClosureDateCorrective !== undefined) payload.ActualClosureDateCorrective = item.ActualClosureDateCorrective;
            if (item.ActionPlanPreventive !== undefined) payload.ActionPlanPreventive = item.ActionPlanPreventive;
            if (item.ResponsibilityPreventive !== undefined) {
                const ids = await this.resolveEmailsToIds(item.ResponsibilityPreventive,context);
                if (ids.length === 1) payload.ResponsibilityPreventiveId = ids[0];
                else if (ids.length > 1) payload.ResponsibilityPreventiveId = { results: ids };
                else {
                    console.warn('RCARepository.saveRCAItem: no user IDs resolved for ResponsibilityPreventive; field omitted from payload');
                }
            }
            if (item.PlannedClosureDatePreventive !== undefined) payload.PlannedClosureDatePreventive = item.PlannedClosureDatePreventive;
            if (item.ActualClosureDatePreventive !== undefined) payload.ActualClosureDatePreventive = item.ActualClosureDatePreventive;


            await this.service.updateItem({
                context,
                listTitle: SubSiteListNames.RootCauseAnalysis,
                itemId: itemId,
                item: payload,
                select: selectedFields,
                expand: expandFields
            });
        }
        catch (error: any) {
            throw new Error('Failed to update RCA item: ' + (error?.message || error));
        }
    }

    public refresh(): void {
        this.cache = null;
        this.cacheTimestamp = 0;
    }

    public getCacheStatus(): { cached: boolean; itemCount: number; age: number } {
        const now = Date.now();
        return {
            cached: this.cache !== null,
            itemCount: this.cache?.length || 0,
            age: this.cache ? now - this.cacheTimestamp : 0
        };
    }
}

const defaultInstance = new RCARepository();

export default defaultInstance;
export const MetricsRepo = defaultInstance;
export const getRCAItems = async (useCache: boolean = false, context?: WebPartContext): Promise<IRCAList[]> => defaultInstance.getRCAItems(useCache, context);
export const saveRCAItem = async (item: IRCAList, context?: WebPartContext): Promise<any> => defaultInstance.saveRCAItem(item, context);
export const updateRCAItem = async (itemId: number, item: IRCAList, context?: WebPartContext): Promise<any> => defaultInstance.updateRCAItem(itemId, item, context);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();

