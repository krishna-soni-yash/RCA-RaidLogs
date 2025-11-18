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
            if (item.RCATypeOfAction !== undefined) {
                try {
                    const raw = item.RCATypeOfAction;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(',') : []);

                    const values: string[] = parts
                        .map(p => p.trim())
                        .filter(p => p.length > 0);

                    if (values.length > 0) {
                        payload.RCATypeOfAction = values ;
                    } else {
                        console.warn('RCARepository.saveRCAItem: no valid values parsed for RCATypeOfAction; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.saveRCAItem: failed to parse RCATypeOfAction', e);
                }
            }

            // helper: try to resolve emails (or simple strings) to SharePoint user IDs using service helpers


            if (item.ActionPlanCorrection !== undefined) payload.ActionPlanCorrection = item.ActionPlanCorrection;
            if (item.ResponsibilityCorrection !== undefined) {
                try {
                    const raw = item.ResponsibilityCorrection;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(/; ?/) : []);

                    const ids: number[] = parts
                        .map(p => (p || '').split('|')[0].trim())
                        .map(s => Number(s))
                        .filter((n) => !isNaN(n) && n > 0);

                    if (ids.length === 1) {
                        payload.ResponsibilityCorrectionId = ids[0];
                    } else if (ids.length > 1) {
                        payload.ResponsibilityCorrectionId = { results: ids };
                    } else {
                        console.warn('RCARepository.saveRCAItem: no valid IDs parsed for ResponsibilityCorrection; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.saveRCAItem: failed to parse ResponsibilityCorrection', e);
                }
            }
            if (item.PlannedClosureDateCorrection !== undefined) payload.PlannedClosureDateCorrection = item.PlannedClosureDateCorrection;
            if (item.ActualClosureDateCorrection !== undefined) payload.ActualClosureDateCorrection = item.ActualClosureDateCorrection;

            if (item.ActionPlanCorrective !== undefined) payload.ActionPlanCorrective = item.ActionPlanCorrective;
            if (item.ResponsibilityCorrective !== undefined) {
                  try {
                    const raw = item.ResponsibilityCorrective;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(/; ?/) : []);

                    const ids: number[] = parts
                        .map(p => (p || '').split('|')[0].trim())
                        .map(s => Number(s))
                        .filter((n) => !isNaN(n) && n > 0);

                    if (ids.length === 1) {
                        payload.ResponsibilityCorrectiveId = ids[0];
                    } else if (ids.length > 1) {
                        payload.ResponsibilityCorrectiveId = { results: ids };
                    } else {
                        console.warn('RCARepository.saveRCAItem: no valid IDs parsed for ResponsibilityCorrection; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.saveRCAItem: failed to parse ResponsibilityCorrective', e);
                }
            }

            if (item.PlannedClosureDateCorrective !== undefined) payload.PlannedClosureDateCorrective = item.PlannedClosureDateCorrective;
            if (item.ActualClosureDateCorrective !== undefined) payload.ActualClosureDateCorrective = item.ActualClosureDateCorrective;
            if (item.ActionPlanPreventive !== undefined) payload.ActionPlanPreventive = item.ActionPlanPreventive;
            if (item.ResponsibilityPreventive !== undefined) {
                  try {
                    const raw = item.ResponsibilityPreventive;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(/; ?/) : []);

                    const ids: number[] = parts
                        .map(p => (p || '').split('|')[0].trim())
                        .map(s => Number(s))
                        .filter((n) => !isNaN(n) && n > 0);

                    if (ids.length === 1) {
                        payload.ResponsibilityPreventiveId = ids[0];
                    } else if (ids.length > 1) {
                        payload.ResponsibilityPreventiveId = { results: ids };
                    } else {
                        console.warn('RCARepository.saveRCAItem: no valid IDs parsed for ResponsibilityPreventive; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.saveRCAItem: failed to parse ResponsibilityPreventive', e);
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
            if (item.RCATypeOfAction !== undefined) {
                try {
                    const raw = item.RCATypeOfAction;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(',') : []);

                    const values: string[] = parts
                        .map(p => p.trim())
                        .filter(p => p.length > 0);

                    if (values.length > 0) {
                        payload.RCATypeOfAction = values;
                    } else {
                        console.warn('RCARepository.updateRCAItem: no valid values parsed for RCATypeOfAction; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.updateRCAItem: failed to parse RCATypeOfAction', e);
                }
            }

            // helper: try to resolve emails (or simple strings) to SharePoint user IDs using service helpers


            if (item.ActionPlanCorrection !== undefined) payload.ActionPlanCorrection = item.ActionPlanCorrection;

            if (item.ResponsibilityCorrection !== undefined) {
                // ResponsibilityCorrection may be stored as:
                // - "id|email; id|email"
                // - array of such strings
                // Extract id (left side of '|') from each part and pass IDs to payload
                try {
                    const raw = item.ResponsibilityCorrection;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(/; ?/) : []);

                    const ids: number[] = parts
                        .map(p => (p || '').split('|')[0].trim())
                        .map(s => Number(s))
                        .filter((n) => !isNaN(n) && n > 0);

                    if (ids.length === 1) {
                        payload.ResponsibilityCorrectionId = ids[0];
                    } else if (ids.length > 1) {
                        payload.ResponsibilityCorrectionId = { results: ids };
                    } else {
                        console.warn('RCARepository.updateRCAItem: no valid IDs parsed for ResponsibilityCorrection; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.updateRCAItem: failed to parse ResponsibilityCorrection', e);
                }
            }
            if (item.PlannedClosureDateCorrection !== undefined) payload.PlannedClosureDateCorrection = item.PlannedClosureDateCorrection;
            if (item.ActualClosureDateCorrection !== undefined) payload.ActualClosureDateCorrection = item.ActualClosureDateCorrection;

            if (item.ActionPlanCorrective !== undefined) payload.ActionPlanCorrective = item.ActionPlanCorrective;
            if (item.ResponsibilityCorrective !== undefined) {
                  try {
                    const raw = item.ResponsibilityCorrective;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(/; ?/) : []);

                    const ids: number[] = parts
                        .map(p => (p || '').split('|')[0].trim())
                        .map(s => Number(s))
                        .filter((n) => !isNaN(n) && n > 0);

                    if (ids.length === 1) {
                        payload.ResponsibilityCorrectiveId = ids[0];
                    } else if (ids.length > 1) {
                        payload.ResponsibilityCorrectiveId = { results: ids };
                    } else {
                        console.warn('RCARepository.updateRCAItem: no valid IDs parsed for ResponsibilityCorrection; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.updateRCAItem: failed to parse ResponsibilityCorrective', e);
                }
            }

            if (item.PlannedClosureDateCorrective !== undefined) payload.PlannedClosureDateCorrective = item.PlannedClosureDateCorrective;
            if (item.ActualClosureDateCorrective !== undefined) payload.ActualClosureDateCorrective = item.ActualClosureDateCorrective;
            if (item.ActionPlanPreventive !== undefined) payload.ActionPlanPreventive = item.ActionPlanPreventive;
            if (item.ResponsibilityPreventive !== undefined) {
                  try {
                    const raw = item.ResponsibilityPreventive;
                    const parts: string[] = Array.isArray(raw)
                        ? raw.map((r: any) => String(r))
                        : (typeof raw === 'string' ? raw.split(/; ?/) : []);

                    const ids: number[] = parts
                        .map(p => (p || '').split('|')[0].trim())
                        .map(s => Number(s))
                        .filter((n) => !isNaN(n) && n > 0);

                    if (ids.length === 1) {
                        payload.ResponsibilityPreventiveId = ids[0];
                    } else if (ids.length > 1) {
                        payload.ResponsibilityPreventiveId = { results: ids };
                    } else {
                        console.warn('RCARepository.updateRCAItem: no valid IDs parsed for ResponsibilityPreventive; field omitted');
                    }
                } catch (e) {
                    console.warn('RCARepository.updateRCAItem: failed to parse ResponsibilityPreventive', e);
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

