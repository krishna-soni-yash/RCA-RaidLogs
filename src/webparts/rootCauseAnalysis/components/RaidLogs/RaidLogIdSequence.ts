import { RaidType } from './interfaces/IRaidItem';
import { formatRaidLogNumber } from './RaidLogIdFormat';

/** Calculates a numeric sequence without relying on text ordering or table state. */
export class RaidLogIdSequence {
  private highest = 0;

  public constructor(private readonly type: RaidType) {}

  public include(value: string | number | null | undefined): void {
    const text = String(value ?? '').trim();
    if (!text) return;

    // Permit existing prefixes (e.g. RISK-0099), but reject decimals/negative IDs.
    const match = /^(.*?)(\d+)$/.exec(text);
    if (!match || (match[1] && !/^[A-Za-z][A-Za-z _-]*$/.test(match[1]))) {
      throw new Error(`Invalid RaidLogID "${text}". Correct this value before creating another item of this type.`);
    }
    const number = Number(match[2]);
    if (!isFinite(number) || number > 9007199254740991) {
      throw new Error('RaidLogID exceeds the supported integer range.');
    }
    // Legacy numeric and prefixed values can coexist as new items are added.
    this.highest = Math.max(this.highest, number);
  }

  public next(): string {
    if (this.highest >= 9007199254740991) {
      throw new Error('No further RaidLogID can be allocated within the supported integer range.');
    }
    return formatRaidLogNumber(this.type, this.highest + 1);
  }
}
