export {
  SharePointService,
  SharePointServiceFactory,
  type ISharePointListItem,
  type IListQueryOptions,
  type ISharePointServiceResponse,
  type IListInfo,
  type IBatchOperation
} from './SharePointService';

export {
  RaidListService,
  RaidServiceFactory,
  type IRaidSharePointItem
} from './RaidListService';

import { SharePointServiceFactory } from './SharePointService';
import { RaidServiceFactory } from './RaidListService';

export {
  type IRaidItem,
  type RaidType,
  type IRaidAction,
  type IPersonPickerUser
} from '../components/RaidLogs/IRaidItem';

export interface IServiceConfiguration {
  raidListName?: string;
  defaultPageSize?: number;
  enableLogging?: boolean;
}

export const DEFAULT_SERVICE_CONFIG: IServiceConfiguration = {
  raidListName: 'RAID_Items',
  defaultPageSize: 50,
  enableLogging: false
};

export class ServiceManager {
  private static config: IServiceConfiguration = DEFAULT_SERVICE_CONFIG;
  
  static configure(config: Partial<IServiceConfiguration>): void {
    ServiceManager.config = { ...DEFAULT_SERVICE_CONFIG, ...config };
  }
  
  static getSharePointService(context: any) {
    return SharePointServiceFactory.getInstance(context);
  }
  
  static getRaidService(context: any) {
    return RaidServiceFactory.getInstance(context, ServiceManager.config.raidListName);
  }
  
  static getConfiguration(): IServiceConfiguration {
    return ServiceManager.config;
  }
}