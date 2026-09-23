const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');
const ts = require('typescript');

const root = path.resolve(__dirname, '../src/webparts/rootCauseAnalysis/components/RaidLogs');

function load(name, mocks = {}) {
  const filename = path.join(root, `${name}.ts`);
  const javascript = ts.transpileModule(fs.readFileSync(filename, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2018 }
  }).outputText;
  const module = { exports: {} };
  const localRequire = id => {
    if (Object.hasOwn(mocks, id)) return mocks[id];
    if (id === '@pnp/sp/fields') return {};
    if (id.startsWith('./')) return load(id.substring(2), mocks);
    throw new Error(`Unexpected dependency: ${id}`);
  };
  vm.runInThisContext(`(function(require, module, exports) { ${javascript}\n})`, { filename })(localRequire, module, module.exports);
  return module.exports;
}

const { RaidLogIdSequence } = load('RaidLogIdSequence');
const { RaidLogIdService } = load('RaidLogIdService');
const types = ['Risk', 'Opportunity', 'Issue', 'Assumption', 'Dependency', 'Constraints'];

function fixture(initial = [], fieldType = 'Number') {
  const rows = initial.map((row, index) => ({ Id: index + 1, ...row }));
  const writes = [];
  let failPage = false;
  const list = {
    fields: { getByInternalNameOrTitle: () => ({ select: () => async () => ({ TypeAsString: fieldType, ReadOnlyField: false }) }) },
    items: {
      getById: id => ({ select: () => ({ expand: () => async () => rows.find(row => row.Id === id) }) }),
      select: () => ({ filter: filter => ({ top: () => ({
        async *[Symbol.asyncIterator]() {
          const type = /SelectType eq '([^']+)'/.exec(filter)[1];
          const matches = rows.filter(row => row.SelectType === type);
          yield matches.slice(0, 1);
          if (failPage) throw new Error('Page request failed');
          yield matches.slice(1);
        }
      }) }) }),
      add: async item => {
        const created = { Id: rows.length + 1, ...item };
        rows.push(created);
        writes.push(created);
        return created;
      }
    }
  };
  const generic = {
    init() {},
    getSiteUrlForList: () => 'https://example.test/site',
    getSpInstanceForSite: () => ({ web: { lists: { getByTitle: () => list } } }),
    cleanItemForRaidSave: item => item,
    fetchAllItems: async options => {
      const id = /^Id eq (\d+)$/.exec(options.filter || '');
      const risk = /^RAIDId eq '([^']+)'$/.exec(options.filter || '');
      return rows.filter(row => id ? row.Id === Number(id[1]) : risk ? row.RAIDId === risk[1] : true);
    },
    updateItem: async options => {
      const row = rows.find(item => item.Id === options.itemId);
      Object.assign(row, options.item);
      return { success: true, item: row };
    },
    deleteItem: async () => ({ success: true })
  };
  const { RaidListService } = load('RaidListService', {
    '../../../../services/GenericServices': { default: generic },
    '../../../../common/Constants': { LIST_NAMES: { RAID_LOGS: 'RAIDLogs' } },
    '../../../../services/RaidLogEmailTriggerService': { default: class { async createEmailTrigger() {} } }
  });
  return { list, rows, writes, service: new RaidListService({}), failPage: () => { failPage = true; } };
}

test('numeric maximum, padding, empty IDs, and prefixes', () => {
  for (const [values, expected] of [
    [[], '1'], [[undefined, null, ''], '1'], [[9, 100, 99], '101'],
    [['0009', '0099'], '0100'], [['RISK-0099', 'RISK-0001'], 'RISK-0100'],
    [[1, 3, 8], '9']
  ]) {
    const sequence = new RaidLogIdSequence();
    values.forEach(value => sequence.include(value));
    assert.equal(sequence.next(), expected);
  }
});

test('invalid data, inconsistent prefixes, and overflow fail closed', () => {
  for (const value of ['abc', '-1', '1.5', Infinity, '9007199254740992']) {
    assert.throws(() => new RaidLogIdSequence().include(value));
  }
  const sequence = new RaidLogIdSequence();
  sequence.include('RISK-1');
  assert.throws(() => sequence.include('ISSUE-2'));
  const overflow = new RaidLogIdSequence();
  overflow.include('9007199254740991');
  assert.throws(() => overflow.next());
});

test('each of the six types gets its own numeric maximum across all pages', async () => {
  const f = fixture(types.flatMap((type, index) => [
    { SelectType: type, RaidLogID: 1 },
    { SelectType: type, RaidLogID: 99 + index }
  ]));
  for (const [index, type] of types.entries()) {
    const created = await f.service.createRaidItem({ type });
    assert.equal(created.raidLogId, String(100 + index));
    assert.equal(f.writes[index].RaidLogID, 100 + index);
  }
});

test('text IDs preserve formatting and empty types start at one', async () => {
  const f = fixture([{ SelectType: 'Issue', RaidLogID: 'ISSUE-0099' }], 'Text');
  assert.equal((await f.service.createRaidItem({ type: 'Issue' })).raidLogId, 'ISSUE-0100');
  assert.equal((await f.service.createRaidItem({ type: 'Dependency' })).raidLogId, '1');
});

test('a failed later page prevents creation', async () => {
  const f = fixture([{ SelectType: 'Issue', RaidLogID: 99 }]);
  f.failPage();
  await assert.rejects(f.service.createRaidItem({ type: 'Issue' }), /Page request failed/);
  assert.equal(f.writes.length, 0);
});

test('unsupported columns prevent allocation', async () => {
  await assert.rejects(new RaidLogIdService(fixture([], 'Calculated').list).next('Risk'), /writable/);
});

const action = { type: 'Mitigation', plan: '', responsibility: [], targetDate: '', actualDate: '', status: '' };

test('Risk action rows share one ID and the next Risk increments once', async () => {
  const f = fixture([{ SelectType: 'Risk', RaidLogID: 105, RAIDId: 'old' }]);
  const created = await f.service.createRiskItemWithActions({ type: 'Risk', raidId: 'new' }, action, { ...action, type: 'Contingency' });
  assert.equal(created.length, 2);
  assert.deepEqual(f.writes.map(row => row.RaidLogID), [106, 106]);
  await f.service.createRiskItemWithActions({ type: 'Risk', raidId: 'next' }, action, null);
  assert.equal(f.writes[2].RaidLogID, 107);
});

test('editing cannot replace the existing ID', async () => {
  const f = fixture([{ SelectType: 'Issue', RaidLogID: 73 }]);
  const updated = await f.service.updateRaidItem(1, { raidLogId: '999', details: 'Changed' });
  assert.equal(updated.raidLogId, '73');
  assert.equal(f.rows[0].RaidLogID, 73);
});

test('adding a Risk action reuses the existing ID without consuming a number', async () => {
  const f = fixture([{ SelectType: 'Risk', RaidLogID: 105, RAIDId: 'old', TypeOfAction: 'Mitigation' }]);
  assert.equal(await f.service.updateRiskItemsByRaidId('old', {}, action, { ...action, type: 'Contingency' }), true);
  assert.equal(f.writes[0].RaidLogID, 105);
  assert.equal(f.writes[0].RAIDId, 'old');
});
