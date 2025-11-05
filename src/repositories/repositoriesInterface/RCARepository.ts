import genericService, { GenericService } from '../../services/GenericServices';
import IGenericService from '../../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRCAList } from '../../models/IRCAList';
import ErrorMessages from '../../common/ErrorMessages';
import { SubSiteListNames,selectedFields,expanFields } from '../../common/Constants';
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
                select:  selectedFields,
                pageSize: 2000,
                expand: expanFields
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
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();

