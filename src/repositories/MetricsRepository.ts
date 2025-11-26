import genericService, { GenericService } from '../services/GenericServices';
import IGenericService from '../services/IGenericServices';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IMetrics } from '../models/IMetrics';
import ErrorMessages from '../common/ErrorMessages';
import ParentListNames, { SubSiteListNames } from '../common/Constants';
import IProjectMetricsRepository from './repositoriesInterface/IProjectMetricsRepository';
//import { getListConfigurationBasedOnMetricLogs } from '../repositories/ObjectivesMasterRepository';
//import IObjectivesMasterRepository from './repositoriesInterface/IObjectivesMasterRepository';

/**
 * Repository for ProjectTypes list
 * Implements a simple cached fetch of Id/LinkTitle values
 */
export class MetricsRepository implements IProjectMetricsRepository {
  private service: IGenericService;
  private cache: IMetrics[] | null = null;
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

  private ActiveVersion: IMetrics[] | null = null;
  private VersionId: number | undefined = undefined;
  public async getMetricsValues(useCache: boolean = true, context?: WebPartContext, selectedProjectType?: string): Promise<IMetrics[]> {
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
        listTitle: ParentListNames.Metrics,
        //select: selectFields,
        pageSize: 2000,
        filter: 'IsActive eq 1 and ProjectType/Title eq \'' + (selectedProjectType) + '\'',
        // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

      });

      const normalized = (items || []).map((it: any) => ({
        ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
        LinkTitle: it?.LinkTitle ?? it?.Title ?? '',
        ProjectType: it?.ProjectType?.Title ?? '',
        NameOfMetrics: it?.NameOfMetrics ?? '',
        IsActive: typeof it?.IsActive === 'boolean' ? it.IsActive : false
      })) as unknown as IMetrics[];

      this.cache = normalized;
      this.cacheTimestamp = now;

      return this.cache;
    } catch (error: any) {
      throw new Error('Failed to fetch ProjectType: ' + (error?.message || error));
    }
  }

  public async getMetricsMeasureandFormulae(useCache: boolean = true, context?: WebPartContext, selectedMetrics?: string, selectedProjectType?: string): Promise<IMetrics[]> {
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
        listTitle: ParentListNames.Metrics,
        //select: selectFields,
        pageSize: 2000,
        filter: 'IsActive eq 1 and NameOfMetrics eq \'' + (selectedMetrics) + '\' and ProjectType/Title eq \'' + (selectedProjectType) + '\'',
        // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

      });

      const normalized = (items || []).map((it: any) => ({
        ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
        LinkTitle: it?.LinkTitle ?? it?.Title ?? '',
        ProjectType: it?.ProjectType?.Title ?? '',
        NameOfMetrics: it?.NameOfMetrics ?? '',
        //  UnitOfMeasure: it?.UnitOfMeasure ?? '',
        //  MetricFormulae: it?.MetricFormulae ?? '',
        //  AssociatedPPM: it?.AssociatedPPM ?? '',
        IsActive: typeof it?.IsActive === 'boolean' ? it.IsActive : false
      })) as unknown as IMetrics[];

      this.cache = normalized;
      this.cacheTimestamp = now;

      return this.cache;
    } catch (error: any) {
      throw new Error('Failed to fetch ProjectType: ' + (error?.message || error));
    }
  }

  public async getMetricsFromProjectMetrics(useCache: boolean = true, context?: WebPartContext, selectedMetrics?: string, selectedProjectType?: string): Promise<IMetrics[]> {
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
      //const listConfig = await getListConfigurationBasedOnMetricLogs(context);
      
      this.ActiveVersion = await this.getApprovedProjectlogs(true, context);
      if (this.ActiveVersion && this.ActiveVersion.length > 0) {
        this.VersionId = this.ActiveVersion[0].ID;
      }
      const items = await this.service.fetchAllItems<any>({
        context,
        listTitle: SubSiteListNames.ProjectMetrics,
        //select: selectFields,
        pageSize: 2000,
        filter: 'IsActive eq 1 and VersionId eq ' + (this.VersionId) + '',
        // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

      });

      const normalized = (items || []).map((it: any) => ({
        ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
        LinkTitle: it?.LinkTitle ?? it?.Title ?? '',
        ProjectType: it?.ProjectType?.Title ?? '',
        // Metrics: it?.NameOfMetrics ?? '',
        PerformanceGoals: it?.PerformanceGoals ?? '',
        UnitOfMeasure: it?.UnitOfMeasure ?? '',
        MetricFormulae: it?.MetricFormulae ?? '',
        AssociatedPPM: it?.AssociatedPPM ?? '',
        Goal: it?.Goal ?? '',
        USL: it?.USL ?? '',
        LSL: it?.LSL ?? '',
        Priority: it?.Priority ?? '',
        DataInput: it?.DataInput ?? '',
        DataSource: it?.DataSource ?? '',
        DataCollectionFrequency: it?.DataCollectionFrequency ?? '',
        DataAnalysisFrequency: it?.DataAnalysisFrequency ?? '',
        BaselineAndRevisionFrequency: it?.BaselineAndRevisionFrequency ?? '',
        Statistical: it?.Statistical ?? '',
        Quantitative: it?.Quantitative ?? '',
        InterpretationGuidelines: it?.InterpretationGuidelines ?? '',
        CausalAnalysisTrigger: it?.CausalAnalysisTrigger ?? '',
        ProbabilityOfSuccessThreshold: it?.ProbabilityOfSuccessThreshold ?? '',
        Process: it?.Process ?? '',
        BG: it?.BG ?? '',
        PG: it?.PG ?? '',
        MetricsFormulae: it?.MetricsFormulae ?? '',
        Metrics: it?.Metrics ?? '',
      })) as unknown as IMetrics[];

      this.cache = normalized;
      this.cacheTimestamp = now;

      return this.cache;
    } catch (error: any) {
      throw new Error('Failed to fetch ProjectType: ' + (error?.message || error));
    }
  }

  public async getSubMetricsFromProjectMetrics(useCache: boolean = true, context?: WebPartContext, selectedMetrics?: string, selectedProjectType?: string): Promise<IMetrics[]> {
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
      //const listConfig = await getListConfigurationBasedOnMetricLogs(context);
      
      // Ensure we have the active version id available
     

      const items = await this.service.fetchAllItems<any>({
        context,
        listTitle: SubSiteListNames.ProjectMetrics,
        //select: selectFields,
        pageSize: 2000,
        filter: 'IsActive eq 1 and Metrics eq \'' + (selectedMetrics) + '\' and  VersionId eq ' + (this.VersionId ?? 0) + '',
        // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

      });

      const normalized = (items || []).map((it: any) => ({
        ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
        LinkTitle: it?.LinkTitle ?? it?.Title ?? '',
        ProjectType: it?.ProjectType?.Title ?? '',
        BG: it?.BG ?? '',
        PG: it?.PG ?? '',
        Metrics: it?.Metrics ?? '',
        UnitOfMeasure: it?.UnitOfMeasure ?? '',
        Process: it?.Process ?? '',
        SubMetrics: it?.SubMetrics ?? '',
        Subprocess: it?.Subprocess ?? '',
        SubGoal: it?.SubGoal ?? '',
        SubUnitOfMeasure: it?.SubUnitOfMeasure ?? '',
        SubMetricsFormulae: it?.SubMetricsFormulae ?? '',
        SubUSL: it?.SubUSL ?? '',
        SubDataInput: it?.SubDataInput ?? '',
        SubLSL: it?.SubLSL ?? '',
        SubDataSource: it?.SubDataSource ?? '',
        SubDataCollectionFrequency: it?.SubDataCollectionFrequency ?? '',
        SubDataAnalysisFrequency: it?.SubDataAnalysisFrequency ?? '',
        SubBaselineAndRevisionFrequency: it?.SubBaselineAndRevisionFrequency ?? '',
        Applicability: it?.Applicability ?? '',
        HasSubProcess: it?.HasSubProcess ?? '',
        OrgStatistical: it?.OrgStatistical ?? '',
        OrgInterpretationGuidelines: it?.OrgInterpretationGuidelines ?? '',
        OrgCausalAnalysisTrigger: it?.OrgCausalAnalysisTrigger ?? '',
        OrgProbabilityOfSuccessThreshold: it?.OrgProbabilityOfSuccessThreshold ?? '',
        statical: it?.Statistical ?? '',
        InterpretationGuidelines: it?.InterpretationGuidelines ?? '',
        CausalAnalysisTrigger: it?.CausalAnalysisTrigger ?? '',
        ProbabilityOfSuccessThreshold: it?.ProbabilityOfSuccessThreshold ?? '',
      })) as unknown as IMetrics[];

      this.cache = normalized;
      this.cacheTimestamp = now;

      return this.cache;
    } catch (error: any) {
      throw new Error('Failed to fetch ProjectType: ' + (error?.message || error));
    }
  }
  public async getApprovedProjectlogs(useCache: boolean = true, context?: WebPartContext): Promise<IMetrics[]> {
    const now = Date.now();

    if (useCache && this.cache && (now - this.cacheTimestamp) < this.CACHE_DURATION) {
      return this.cache;
    }
    try {
      const genericServiceInstance: IGenericService = new GenericService(undefined, context);
      genericServiceInstance.init(undefined, context);
      // const selectFields: string[] = ['Id', 'LinkTitle','ProjectType','IsActive'];
      //const listConfig = await getListConfigurationBasedOnMetricLogs(context);

      const items = await this.service.fetchAllItems<any>({
        context,
        listTitle: SubSiteListNames.ProjectMetricLogs,
        //select: selectFields,
        pageSize: 2000,
        filter: 'IsActive eq 1',
        // filter: 'IsActive eq true and ProjectType in (' + (selectedProjectTypes?.map(pt => `'${pt}'`).join(',') || '') + ')',

      });

      const normalized = (items || []).map((it: any) => ({
        ID: typeof it?.ID === 'number' ? it.ID : (typeof it?.Id === 'number' ? it.Id : 0),
      })) as unknown as IMetrics[];

      this.cache = normalized;
      this.cacheTimestamp = now;

      return this.cache;
    } catch (error: any) {
      throw new Error('Failed to fetch ProjectType: ' + (error?.message || error));
    }

    if (!context) {
      throw new Error(ErrorMessages.WEBPART_CONTEXT_REQUIRED_OBJECTIVES);
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

const defaultInstance = new MetricsRepository();

export default defaultInstance;
export const MetricsRepo = defaultInstance;
export const getMetricsValues = async (useCache: boolean = false, context?: WebPartContext, selectedProjectType?: string): Promise<IMetrics[]> => defaultInstance.getMetricsValues(useCache, context, selectedProjectType);
export const refresh = (): void => defaultInstance.refresh();
export const getCacheStatus = (): { cached: boolean; itemCount: number; age: number } => defaultInstance.getCacheStatus();
export const getMetricsMeasureandFormulae = async (useCache: boolean = false, context?: WebPartContext, selectedMetrics?: string, selectedProjectType?: string): Promise<IMetrics[]> => defaultInstance.getMetricsMeasureandFormulae(useCache, context, selectedMetrics, selectedProjectType);
export const getMetricsFromProjectMetrics = async (useCache: boolean = false, context?: WebPartContext, selectedMetrics?: string, selectedProjectType?: string): Promise<IMetrics[]> => defaultInstance.getMetricsFromProjectMetrics(useCache, context, selectedMetrics, selectedProjectType);
export const getSubMetricsFromProjectMetrics = async (useCache: boolean = false, context?: WebPartContext, selectedMetrics?: string, selectedProjectType?: string): Promise<IMetrics[]> => defaultInstance.getSubMetricsFromProjectMetrics(useCache, context, selectedMetrics, selectedProjectType);