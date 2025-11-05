//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRCAList } from '../models/IRCAList';
import IGenericService from '../services/IGenericServices';

export interface IRCARepository {
  getRCAItems(useCache?: boolean, context?: WebPartContext): Promise<IRCAList[]>;
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
}

export default IRCARepository;