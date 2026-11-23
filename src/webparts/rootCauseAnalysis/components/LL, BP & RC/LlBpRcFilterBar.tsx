import * as React from 'react';
import { DefaultButton, SearchBox } from '@fluentui/react';

export interface ILlBpRcFilterField {
  key: string;
  label: string;
}

/** Case-insensitive contains match; an empty search matches everything. */
export const matchesTextFilter = (value: string | undefined, search: string | undefined): boolean => {
  const term = (search ?? '').trim().toLowerCase();
  if (!term) return true;
  return (value ?? '').toString().toLowerCase().indexOf(term) !== -1;
};


export interface ILlBpRcFilterBarProps {
  fields: ILlBpRcFilterField[];
  values: Record<string, string>;
  onChange: (key: string, value: string) => void;
  onClear: () => void;
}

const LlBpRcFilterBar: React.FC<ILlBpRcFilterBarProps> = ({ fields, values, onChange, onClear }) => {
  const hasActiveFilters = fields.some(field => !!values[field.key]?.trim());

  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'flex-end', gap: 12, marginTop: 8, marginBottom: 8 }}>
      {fields.map(field => (
        <SearchBox
          key={field.key}
          placeholder={`Search ${field.label}`}
          ariaLabel={`Search by ${field.label}`}
          value={values[field.key] ?? ''}
          onChange={(_, value) => onChange(field.key, value || '')}
          clearButtonProps={{ ariaLabel: `Clear ${field.label} search` }}
          styles={{ root: { width: 260, maxWidth: '100%' } }}
        />
      ))}

      <DefaultButton
        text="Clear filters"
        iconProps={{ iconName: 'ClearFilter' }}
        onClick={onClear}
        disabled={!hasActiveFilters}
      />
    </div>
  );
};

export default LlBpRcFilterBar;
