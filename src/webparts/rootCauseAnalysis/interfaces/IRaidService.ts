import { IRaidItem } from '../components/RaidLogs/IRaidItem';

/**
 * Extended RAID Item interface
 */
export interface IExtendedRaidItem extends IRaidItem {
  // Extended properties can be added here if needed
}

/**
 * RAID Service Configuration
 */
export interface IRaidServiceConfig {
  listName: string;
  enableBulkOperations: boolean;
  maxBulkSize: number;
  enableLogging: boolean;
}