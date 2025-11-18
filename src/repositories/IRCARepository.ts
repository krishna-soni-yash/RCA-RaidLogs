//IProjectMetricsRepository
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IRCAList } from '../models/IRCAList';
import IGenericService from '../services/IGenericServices';

export interface IRCARepository {
  getRCAItems(useCache?: boolean, context?: WebPartContext): Promise<IRCAList[]>;
  saveRCAItem(item: IRCAList, context?: WebPartContext): Promise<any>;
  updateRCAItem(itemId: number, item: IRCAList, context?: WebPartContext): Promise<any>;
  refresh(): void;
  getCacheStatus(): { cached: boolean; itemCount: number; age: number };
  setService(service: IGenericService): void;
}

export default IRCARepository;