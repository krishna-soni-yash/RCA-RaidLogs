import * as React from 'react';
import { DefaultButton, Dropdown, IDropdownOption, SearchBox } from '@fluentui/react';
import {
  RCA_ALL_FILTER_KEY,
  RCA_PRIORITY_FILTER_OPTIONS,
  RCA_TYPE_OF_ACTION_FILTER_OPTIONS,
  IRcaColumnFilters,
  hasActiveRcaFilters
} from './RcaFilters';

export interface IRcaFilterBarProps {
  filters: IRcaColumnFilters;
  onChange: (patch: Partial<IRcaColumnFilters>) => void;
  onClear: () => void;
}

const toFilterValue = (option?: IDropdownOption): string => {
  if (!option || option.key === RCA_ALL_FILTER_KEY) return '';
  return String(option.key);
};

const RcaFilterBar: React.FC<IRcaFilterBarProps> = ({ filters, onChange, onClear }) => {
  const fieldStyles = { root: { minWidth: 180 } };

  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'flex-end', gap: 12, marginBottom: 16 }}>
      <SearchBox
        placeholder="Search Problem statement"
        ariaLabel="Search Causal Analysis by Problem statement"
        value={filters.problem}
        onChange={(_, value) => onChange({ problem: value || '' })}
        clearButtonProps={{ ariaLabel: 'Clear Problem statement search' }}
        styles={{ root: { width: 260, maxWidth: '100%' } }}
      />

      <Dropdown
        label="Type of Action"
        options={RCA_TYPE_OF_ACTION_FILTER_OPTIONS}
        selectedKey={filters.typeOfAction || RCA_ALL_FILTER_KEY}
        onChange={(_, option) => onChange({ typeOfAction: toFilterValue(option) })}
        styles={fieldStyles}
      />

      <Dropdown
        label="Priority"
        options={RCA_PRIORITY_FILTER_OPTIONS}
        selectedKey={filters.priority || RCA_ALL_FILTER_KEY}
        onChange={(_, option) => onChange({ priority: toFilterValue(option) })}
        styles={fieldStyles}
      />

      <SearchBox
        placeholder="Search Responsibility"
        ariaLabel="Search Causal Analysis by Responsibility"
        value={filters.responsibility}
        onChange={(_, value) => onChange({ responsibility: value || '' })}
        clearButtonProps={{ ariaLabel: 'Clear Responsibility search' }}
        styles={{ root: { width: 260, maxWidth: '100%' } }}
      />

      <DefaultButton
        text="Clear filters"
        iconProps={{ iconName: 'ClearFilter' }}
        onClick={onClear}
        disabled={!hasActiveRcaFilters(filters)}
      />
    </div>
  );
};

export default RcaFilterBar;
