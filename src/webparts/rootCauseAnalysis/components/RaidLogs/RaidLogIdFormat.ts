import { RaidType } from './interfaces/IRaidItem';

const ID_PREFIXES: Record<RaidType, string> = {
  Issue: 'I',
  Assumption: 'A',
  Dependency: 'D',
  Risk: 'R',
  Constraints: 'C',
  Opportunity: 'O'
};

/** Uses the same prefix and minimum padding for storage and display. */
export function formatRaidLogNumber(type: RaidType, value: string | number): string {
  const digits = String(value).replace(/^0+(?=\d)/, '');
  return `${ID_PREFIXES[type]}-${digits.length < 2 ? `0${digits}` : digits}`;
}

/** Formats legacy numeric or prefixed IDs without rewriting existing records. */
export function formatRaidLogId(type: RaidType, value: string | number | undefined): string | undefined {
  const text = String(value ?? '');
  if (!text) return undefined;

  const match = /^(?:[A-Za-z][A-Za-z _-]*)?(\d+)$/.exec(text.trim());
  if (!match) return text;
  return formatRaidLogNumber(type, match[1]);
}
