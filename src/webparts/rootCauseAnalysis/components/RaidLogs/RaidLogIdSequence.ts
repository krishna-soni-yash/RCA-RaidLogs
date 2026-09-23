/** Calculates a numeric sequence without relying on text ordering or table state. */
export class RaidLogIdSequence {
  private highest = 0;
  private prefix: string | undefined;
  private width = 0;

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
    if (this.prefix !== undefined && this.prefix !== match[1]) {
      throw new Error('RaidLogID values for this type have inconsistent prefixes. Correct them before creating another item.');
    }
    this.prefix = match[1];
    this.highest = Math.max(this.highest, number);
    if (match[2].length > 1 && match[2].charAt(0) === '0') {
      this.width = Math.max(this.width, match[2].length);
    }
  }

  public next(): string {
    if (this.highest >= 9007199254740991) {
      throw new Error('No further RaidLogID can be allocated within the supported integer range.');
    }
    let digits = String(this.highest + 1);
    while (digits.length < this.width) digits = `0${digits}`;
    return `${this.prefix || ''}${digits}`;
  }
}
