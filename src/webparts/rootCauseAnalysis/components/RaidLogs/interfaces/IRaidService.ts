import { IRaidItem } from './IRaidItem';

export interface IExtendedRaidItem extends IRaidItem {
}

export interface IRaidServiceConfig {
  listName: string;
  enableBulkOperations: boolean;
  maxBulkSize: number;
  enableLogging: boolean;
}