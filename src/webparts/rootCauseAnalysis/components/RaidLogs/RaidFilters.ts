import { IDropdownOption } from '@fluentui/react';
import { RaidType } from './interfaces/IRaidItem';
import { IExtendedRaidItem } from './interfaces/IRaidService';
import { DROPDOWN_OPTIONS } from '../../../../common/Constants';

export interface IRaidColumnFilters {
  id: string;
  text: string;
  numericValue: string;
  typeOfAction: string;
  status: string;
  identifiedBy?: string;
  date?: Date;
}

export interface IRaidFilterConfig {
  textField: 'description' | 'details';
  textLabel: string;
  numericField?: { field: 'riskExposure' | 'opportunityValue'; label: string };
  showTypeOfAction: boolean;
  showStatus: boolean;
  showDateRange: boolean;
  showIdentifiedBy: boolean;
}

export const ALL_FILTER_KEY = '__all__';

export const EMPTY_RAID_FILTERS: IRaidColumnFilters = {
  id: '',
  text: '',
  numericValue: '',
  typeOfAction: '',
  status: ''
};

export const getRaidFilterConfig = (tab: RaidType): IRaidFilterConfig => {
  switch (tab) {
    case 'Risk':
      return {
        textField: 'description',
        textLabel: 'Description',
        numericField: { field: 'riskExposure', label: 'Risk Exposure' },
        showTypeOfAction: true,
        showStatus: false,
        showDateRange: false,
        showIdentifiedBy: false
      };
    case 'Opportunity':
      return {
        textField: 'description',
        textLabel: 'Description',
        numericField: { field: 'opportunityValue', label: 'Opportunity Value' },
        showTypeOfAction: false,
        showStatus: true,
        showDateRange: false,
        showIdentifiedBy: false
      };
    default:
      return {
        textField: 'details',
        textLabel: 'Details',
        showTypeOfAction: false,
        showStatus: false,
        showDateRange: true,
        showIdentifiedBy: true
      };
  }
};

export const hasActiveRaidFilters = (filters: IRaidColumnFilters): boolean =>
  !!filters.id.trim() ||
  !!filters.text.trim() ||
  !!filters.numericValue.trim() ||
  !!filters.typeOfAction ||
  !!filters.status ||
  !!filters.identifiedBy?.trim() ||
  !!filters.date;

/** Returns the local calendar day as YYYY-MM-DD, or undefined when the value is not a valid date. */
const getDayKey = (value?: string | Date): string | undefined => {
  if (!value) return undefined;
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return undefined;
  const y = date.getFullYear();
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  const d = ('0' + date.getDate()).slice(-2);
  return `${y}-${m}-${d}`;
};

const getActionTypes = (item: IExtendedRaidItem): string[] => {
  const types = (item.actions || []).map(action => action.type).filter(Boolean);
  if (item.typeOfAction) types.push(item.typeOfAction);
  return types;
};

/** Returns the names shown in the Identified By column, so the search matches what the user sees. */
const getByWhomNames = (byWhom: IExtendedRaidItem['byWhom']): string[] => {
  if (!byWhom) return [];
  if (typeof byWhom === 'string') return [byWhom];
  return byWhom.map(user => (user.displayName || '').toLowerCase()).filter(Boolean);
};

export const matchesRaidFilters = (
  item: IExtendedRaidItem,
  filters: IRaidColumnFilters,
  config: IRaidFilterConfig
): boolean => {
  const searchId = filters.id.trim().toLowerCase();
  if (searchId) {
    const itemId = (item.raidLogId || '').toLowerCase();
    if (itemId.indexOf(searchId) === -1) return false;
  }

  const searchText = filters.text.trim().toLowerCase();
  if (searchText) {
    const value = (item[config.textField] || '').toString().toLowerCase();
    if (value.indexOf(searchText) === -1) return false;
  }

  const searchNumeric = filters.numericValue.trim();
  if (config.numericField && searchNumeric) {
    const value = item[config.numericField.field];
    if (value === undefined || value === null || String(value).indexOf(searchNumeric) === -1) return false;
  }

  if (config.showTypeOfAction && filters.typeOfAction) {
    if (getActionTypes(item).indexOf(filters.typeOfAction) === -1) return false;
  }

  if (config.showStatus && filters.status) {
    if (item.status !== filters.status) return false;
  }

  const searchIdentifiedBy = filters.identifiedBy?.trim().toLowerCase();
  if (config.showIdentifiedBy && searchIdentifiedBy) {
    if (!getByWhomNames(item.byWhom).some(name => name.indexOf(searchIdentifiedBy) !== -1)) return false;
  }

  if (config.showDateRange && filters.date) {
    if (getDayKey(item.date) !== getDayKey(filters.date)) return false;
  }

  return true;
};

export const buildStatusFilterOptions = (): IDropdownOption[] => [
  { key: ALL_FILTER_KEY, text: 'All' },
  ...DROPDOWN_OPTIONS.STATUS
];

export const buildActionTypeFilterOptions = (): IDropdownOption[] => [
  { key: ALL_FILTER_KEY, text: 'All' },
  ...DROPDOWN_OPTIONS.ACTION_TYPE
];

/** Formats a filter date as MM/DD/YYYY using local calendar values. */
export const formatFilterDate = (date?: Date): string => {
  if (!date || isNaN(date.getTime())) return '';
  const m = ('0' + (date.getMonth() + 1)).slice(-2);
  const d = ('0' + date.getDate()).slice(-2);
  return `${m}/${d}/${date.getFullYear()}`;
};
