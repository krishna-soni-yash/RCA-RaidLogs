import { IDropdownOption } from '@fluentui/react';
import { IRCAList } from '../../../../models/IRCAList';
import { formatResponsibilityValue } from '../../../../common/exportHelpers';

export const RCA_ALL_FILTER_KEY = '__all__';

export interface IRcaColumnFilters {
  problem: string;
  typeOfAction: string;
  priority: string;
  responsibility: string;
}

export const EMPTY_RCA_FILTERS: IRcaColumnFilters = {
  problem: '',
  typeOfAction: '',
  priority: '',
  responsibility: ''
};

export const hasActiveRcaFilters = (filters: IRcaColumnFilters): boolean =>
  !!filters.problem.trim() ||
  !!filters.typeOfAction ||
  !!filters.priority ||
  !!filters.responsibility.trim();

export const RCA_TYPE_OF_ACTION_FILTER_OPTIONS: IDropdownOption[] = [
  { key: RCA_ALL_FILTER_KEY, text: 'All' },
  { key: 'Correction', text: 'Correction' },
  { key: 'Corrective Action', text: 'Corrective Action' },
  { key: 'Preventive Action', text: 'Preventive Action' }
];

export const RCA_PRIORITY_FILTER_OPTIONS: IDropdownOption[] = [
  { key: RCA_ALL_FILTER_KEY, text: 'All' },
  { key: 'High', text: 'High' },
  { key: 'Medium', text: 'Medium' },
  { key: 'Low', text: 'Low' }
];

const getTypeOfActionList = (value: unknown): string[] => {
  if (Array.isArray(value)) return value.map(v => String(v).trim()).filter(Boolean);
  if (typeof value === 'string') return value.split(',').map(v => v.trim()).filter(Boolean);
  return [];
};

export const matchesRcaFilters = (item: Partial<IRCAList>, filters: IRcaColumnFilters): boolean => {
  const searchProblem = filters.problem.trim().toLowerCase();
  if (searchProblem && (item.LinkTitle || '').toLowerCase().indexOf(searchProblem) === -1) return false;

  if (filters.typeOfAction && getTypeOfActionList(item.RCATypeOfAction).indexOf(filters.typeOfAction) === -1) return false;

  if (filters.priority && item.RCAPriority !== filters.priority) return false;

  const searchResponsibility = filters.responsibility.trim().toLowerCase();
  if (searchResponsibility) {
    const responsibilities = [
      formatResponsibilityValue(item.ResponsibilityCorrection),
      formatResponsibilityValue(item.ResponsibilityCorrective),
      formatResponsibilityValue(item.ResponsibilityPreventive)
    ].join(', ').toLowerCase();
    if (responsibilities.indexOf(searchResponsibility) === -1) return false;
  }

  return true;
};
