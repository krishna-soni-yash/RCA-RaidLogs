/**
 * SharePoint Services Index
 * 
 * This file exports all SharePoint-related services and interfaces
 * for easy importing throughout the application.
 */

// Core SharePoint Service
export {
  SharePointService,
  SharePointServiceFactory,
  type ISharePointListItem,
  type IListQueryOptions,
  type ISharePointServiceResponse,
  type IListInfo,
  type IBatchOperation
} from './SharePointService';

// RAID-specific Service
export {
  RaidListService,
  RaidServiceFactory,
  type IRaidSharePointItem
} from './RaidListService';

// Import for internal use
import { SharePointServiceFactory } from './SharePointService';
import { RaidServiceFactory } from './RaidListService';

// Re-export RAID types for convenience
export {
  type IRaidItem,
  type RaidType,
  type IRaidAction,
  type IPersonPickerUser
} from '../components/RaidLogs/IRaidItem';

/**
 * Service Configuration
 */
export interface IServiceConfiguration {
  raidListName?: string;
  defaultPageSize?: number;
  enableLogging?: boolean;
}

/**
 * Default configuration values
 */
export const DEFAULT_SERVICE_CONFIG: IServiceConfiguration = {
  raidListName: 'RAID_Items',
  defaultPageSize: 50,
  enableLogging: false
};

/**
 * Service Initialization Helper
 * 
 * Convenience function to initialize all services with configuration
 */
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