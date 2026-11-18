import * as React from 'react';
import { DefaultButton, DatePicker, Dropdown, IDropdownOption, SearchBox } from '@fluentui/react';
import { RaidType } from './interfaces/IRaidItem';
import {
  ALL_FILTER_KEY,
  IRaidColumnFilters,
  buildActionTypeFilterOptions,
  buildStatusFilterOptions,
  formatFilterDate,
  getRaidFilterConfig,
  hasActiveRaidFilters
} from './RaidFilters';

export interface IRaidFilterBarProps {
  currentTab: RaidType;
  filters: IRaidColumnFilters;
  onChange: (patch: Partial<IRaidColumnFilters>) => void;
  onClear: () => void;
}

const toFilterValue = (option?: IDropdownOption): string => {
  if (!option || option.key === ALL_FILTER_KEY) return '';
  return String(option.key);
};


const RaidFilterBar: React.FC<IRaidFilterBarProps> = ({ currentTab, filters, onChange, onClear }) => {
  const config = getRaidFilterConfig(currentTab);

  const statusOptions = React.useMemo(() => buildStatusFilterOptions(), []);
  const actionTypeOptions = React.useMemo(() => buildActionTypeFilterOptions(), []);

  const fieldStyles = { root: { minWidth: 180 } };

  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'flex-end', gap: 12, marginBottom: 16 }}>
      <SearchBox
        placeholder="Filter by ID"
        ariaLabel={`Filter ${currentTab} by ID`}
        value={filters.id}
        onChange={(_, value) => onChange({ id: value || '' })}
        clearButtonProps={{ ariaLabel: 'Clear ID filter' }}
        styles={{ root: { width: 260, maxWidth: '100%' } }}
      />

      <SearchBox
        placeholder={`Search ${config.textLabel}`}
        ariaLabel={`Search ${currentTab} by ${config.textLabel}`}
        value={filters.text}
        onChange={(_, value) => onChange({ text: value || '' })}
        clearButtonProps={{ ariaLabel: `Clear ${config.textLabel} search` }}
        styles={{ root: { width: 260, maxWidth: '100%' } }}
      />

      {config.numericField && (
        <SearchBox
          placeholder={`Search ${config.numericField.label}`}
          ariaLabel={`Search ${currentTab} by ${config.numericField.label}`}
          value={filters.numericValue}
          onChange={(_, value) => onChange({ numericValue: value || '' })}
          clearButtonProps={{ ariaLabel: `Clear ${config.numericField.label} search` }}
          styles={{ root: { width: 260, maxWidth: '100%' } }}
        />
      )}

      {config.showTypeOfAction && (
        <Dropdown
          label="Type of Action"
          options={actionTypeOptions}
          selectedKey={filters.typeOfAction || ALL_FILTER_KEY}
          onChange={(_, option) => onChange({ typeOfAction: toFilterValue(option) })}
          styles={fieldStyles}
        />
      )}

      {config.showStatus && (
        <Dropdown
          label="Status"
          options={statusOptions}
          selectedKey={filters.status || ALL_FILTER_KEY}
          onChange={(_, option) => onChange({ status: toFilterValue(option) })}
          styles={fieldStyles}
        />
      )}

      {config.showDateRange && (
        <DatePicker
          placeholder="Select a date"
          textField={{
            ariaLabel: `Filter ${currentTab} by Date`,
            styles: {
              root: { width: '100%', marginBottom: '-5px' },
              fieldGroup: { height: 32 },
              field: { paddingLeft: 36, lineHeight: '30px' }
            }
          }}
          value={filters.date}
          formatDate={formatFilterDate}
          onSelectDate={(date) => onChange({ date: date || undefined })}
          styles={{ root: { width: 260, maxWidth: '100%', marginTop: 0} }}
        />
      )}

      {config.showIdentifiedBy && (
        <SearchBox
          placeholder="Search Identified By"
          ariaLabel={`Search ${currentTab} by Identified By`}
          value={filters.identifiedBy || ''}
          onChange={(_, value) => onChange({ identifiedBy: value || '' })}
          clearButtonProps={{ ariaLabel: 'Clear Identified By search' }}
          styles={{ root: { width: 260, maxWidth: '100%' } }}
        />
      )}

      <DefaultButton
        text="Clear filters"
        iconProps={{ iconName: 'ClearFilter' }}
        onClick={onClear}
        disabled={!hasActiveRaidFilters(filters)}
      />
    </div>
  );
};

export default RaidFilterBar;
