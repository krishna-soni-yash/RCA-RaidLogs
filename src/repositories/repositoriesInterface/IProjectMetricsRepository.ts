//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IMetrics } from '../../models/IMetrics';
import IGenericService from '../../services/IGenericServices';

export interface IProjectMetricsRepository {
  getMetricsValues(useCache?: boolean, context?: WebPartContext, projectType?: string): Promise<IMetrics[]>;
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
  getMetricsMeasureandFormulae(useCache?: boolean, context?: WebPartContext, selectedMetrics?: string,selectedProjectType?:string): Promise<IMetrics[]>;
  getMetricsFromProjectMetrics(useCache?: boolean, context?: WebPartContext, selectedMetrics?: string,selectedProjectType?:string): Promise<IMetrics[]>;
  getSubMetricsFromProjectMetrics(useCache?: boolean, context?: WebPartContext, selectedMetrics?: string,selectedProjectType?:string): Promise<IMetrics[]>;
}

export default IProjectMetricsRepository;