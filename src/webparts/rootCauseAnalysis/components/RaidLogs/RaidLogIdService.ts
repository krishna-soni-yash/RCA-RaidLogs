import '@pnp/sp/fields';
import { IList } from '@pnp/sp/lists';
import { RaidType } from './interfaces/IRaidItem';
import { RaidLogIdSequence } from './RaidLogIdSequence';

export class RaidLogIdService {
  public constructor(private readonly list: IList) {}

  public async toFieldValue(value: string | number): Promise<string | number> {
    const field = await this.list.fields.getByInternalNameOrTitle('RaidLogID')
      .select('TypeAsString', 'ReadOnlyField')();
    if (field.ReadOnlyField || (field.TypeAsString !== 'Number' && field.TypeAsString !== 'Text')) {
      throw new Error('RaidLogID must be a writable Number or Single line of text column.');
    }
    if (field.TypeAsString === 'Number') {
      if (!/^\d+$/.test(String(value))) throw new Error('The numeric RaidLogID column contains an invalid ID.');
      return Number(value);
    }
    return String(value);
  }

  public async next(type: RaidType): Promise<string | number> {
    const sequence = new RaidLogIdSequence();
    const query = this.list.items.select('RaidLogID')
      .filter(`SelectType eq '${type.replace(/'/g, "''")}'`).top(2000);
    // PnP v4's iterator follows every continuation link. A failed page must
    // abort creation: a partial maximum could allocate an existing ID.
    for await (const page of query) {
      for (const item of page) sequence.include(item.RaidLogID);
    }
    const next = sequence.next();
    return this.toFieldValue(next);
  }
}
