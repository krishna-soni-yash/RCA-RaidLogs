import '@pnp/sp/fields';
import { IList } from '@pnp/sp/lists';
import { RaidType } from './interfaces/IRaidItem';
import { RaidLogIdSequence } from './RaidLogIdSequence';

export class RaidLogIdService {
  public constructor(private readonly list: IList) {}

  public async toFieldValue(value: string | number): Promise<string> {
    const field = await this.list.fields.getByInternalNameOrTitle('RaidLogID')
      .select('TypeAsString', 'ReadOnlyField')();
    if (field.ReadOnlyField || field.TypeAsString !== 'Text') {
      throw new Error('RaidLogID must be a writable Single line of text column to store prefixed IDs. Update the column type before creating new RAID items.');
    }
    return String(value);
  }

  public async next(type: RaidType): Promise<string> {
    const sequence = new RaidLogIdSequence(type);
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
