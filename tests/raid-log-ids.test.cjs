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
const displayPrefixes = { Risk: 'R', Opportunity: 'O', Issue: 'I', Assumption: 'A', Dependency: 'D', Constraints: 'C' };

function fixture(initial = [], fieldType = 'Text', readOnly = false) {
  const rows = initial.map((row, index) => ({ Id: index + 1, ...row }));
  const writes = [];
  let failPage = false;
  const list = {
    fields: { getByInternalNameOrTitle: () => ({ select: () => async () => ({ TypeAsString: fieldType, ReadOnlyField: readOnly }) }) },
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

test('numeric maximum accepts legacy prefixes, mixed formats, empty IDs, and gaps', () => {
  for (const [values, expected] of [
    [[], 'R-01'], [[undefined, null, ''], 'R-01'], [[9, 100, 99], 'R-101'],
    [['0009', '0099'], 'R-100'], [['RISK-0099', 'RISK-0001'], 'R-100'],
    [[1, 'R-03', 'RISK-0008'], 'R-09']
  ]) {
    const sequence = new RaidLogIdSequence('Risk');
    values.forEach(value => sequence.include(value));
    assert.equal(sequence.next(), expected);
  }
});

test('invalid data and overflow fail closed', () => {
  for (const value of ['abc', '-1', '1.5', 'A-1.5', Infinity, '9007199254740992']) {
    assert.throws(() => new RaidLogIdSequence('Assumption').include(value));
  }
  const overflow = new RaidLogIdSequence('Assumption');
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
    const prefix = displayPrefixes[type];
    assert.equal(created.raidLogId, `${prefix}-${100 + index}`);
    assert.equal(f.writes[index].RaidLogID, `${prefix}-${100 + index}`);
  }
});

test('legacy long prefixes use canonical formatting and empty types start at one', async () => {
  const f = fixture([{ SelectType: 'Issue', RaidLogID: 'ISSUE-0099' }], 'Text');
  assert.equal((await f.service.createRaidItem({ type: 'Issue' })).raidLogId, 'I-100');
  assert.equal(f.writes[0].RaidLogID, 'I-100');
  assert.equal((await f.service.createRaidItem({ type: 'Dependency' })).raidLogId, 'D-01');
});

test('all six types store padded prefixed IDs with independent sequences', async () => {
  const f = fixture();
  for (const [type, prefix] of Object.entries(displayPrefixes)) {
    assert.equal((await f.service.createRaidItem({ type })).raidLogId, `${prefix}-01`);
    assert.equal((await f.service.createRaidItem({ type })).raidLogId, `${prefix}-02`);
  }
  assert.deepEqual(f.writes.map(row => row.RaidLogID), types.flatMap(type => [
    `${displayPrefixes[type]}-01`, `${displayPrefixes[type]}-02`
  ]));
});

test('all six types continue numeric, prefixed, and mixed sequences and display IDs consistently', async () => {
  for (const [type, prefix] of Object.entries(displayPrefixes)) {
    for (const values of [[1, 2, 3], ['1', '2', '3'], [`${prefix}-01`, `${prefix}-02`, `${prefix}-03`], [1, `${prefix}-02`, '3']]) {
      const f = fixture(values.map(value => ({ SelectType: type, RaidLogID: value })));
      for (let index = 0; index < values.length; index++) {
        assert.equal((await f.service.getRaidItemById(index + 1)).raidLogId, `${prefix}-0${index + 1}`);
        assert.equal(f.rows[index].RaidLogID, values[index]);
      }
      assert.equal((await f.service.createRaidItem({ type, raidLogId: '999' })).raidLogId, `${prefix}-04`);
      assert.equal(f.writes[0].RaidLogID, `${prefix}-04`);
      assert.equal((await f.service.createRaidItem({ type })).raidLogId, `${prefix}-05`);
      assert.equal(f.writes[1].RaidLogID, `${prefix}-05`);
    }
  }
});

test('existing numeric and prefixed IDs use the same display format without rewriting stored IDs', async () => {
  for (const [type, prefix] of Object.entries(displayPrefixes)) {
    for (const stored of [1, '01', `${prefix}-01`, `${type.toUpperCase()}-0001`]) {
      const f = fixture([{ SelectType: type, RaidLogID: stored }], 'Text');
      assert.equal((await f.service.getRaidItemById(1)).raidLogId, `${prefix}-01`);
      const updated = await f.service.updateRaidItem(1, { raidLogId: '999', details: 'Changed' });
      assert.equal(updated.raidLogId, `${prefix}-01`);
      assert.equal(f.rows[0].RaidLogID, stored);
      assert.equal((await f.service.createRaidItem({ type })).raidLogId, `${prefix}-02`);
      assert.equal(f.writes[0].RaidLogID, `${prefix}-02`);
    }
    const f = fixture([{ SelectType: type, RaidLogID: `${prefix}-99` }], 'Text');
    assert.equal((await f.service.createRaidItem({ type })).raidLogId, `${prefix}-100`);
    assert.equal(f.writes[0].RaidLogID, `${prefix}-100`);
  }
});

test('a failed later page prevents creation', async () => {
  const f = fixture([{ SelectType: 'Issue', RaidLogID: 99 }]);
  f.failPage();
  await assert.rejects(f.service.createRaidItem({ type: 'Issue' }), /Page request failed/);
  assert.equal(f.writes.length, 0);
});

test('numeric, unsupported, and read-only columns prevent creation with an actionable error', async () => {
  for (const [fieldType, readOnly] of [['Number', false], ['Calculated', false], ['Text', true]]) {
    const f = fixture([], fieldType, readOnly);
    await assert.rejects(f.service.createRaidItem({ type: 'Risk' }), /writable Single line of text/);
    assert.equal(f.writes.length, 0);
    await assert.rejects(new RaidLogIdService(f.list).toFieldValue('R-01'), /Update the column type/);
  }
});

test('malformed or overflowing stored IDs prevent writes', async () => {
  for (const value of ['A-1.5', 'invalid', '9007199254740992', 'A-9007199254740991']) {
    const f = fixture([{ SelectType: 'Assumption', RaidLogID: value }]);
    await assert.rejects(f.service.createRaidItem({ type: 'Assumption' }));
    assert.equal(f.writes.length, 0);
  }
});

const action = { type: 'Mitigation', plan: '', responsibility: [], targetDate: '', actualDate: '', status: '' };

test('Risk action rows share one ID and the next Risk increments once', async () => {
  const f = fixture([{ SelectType: 'Risk', RaidLogID: 105, RAIDId: 'old' }]);
  const created = await f.service.createRiskItemWithActions({ type: 'Risk', raidId: 'new' }, action, { ...action, type: 'Contingency' });
  assert.equal(created.length, 2);
  assert.deepEqual(f.writes.map(row => row.RaidLogID), ['R-106', 'R-106']);
  await f.service.createRiskItemWithActions({ type: 'Risk', raidId: 'next' }, action, null);
  assert.equal(f.writes[2].RaidLogID, 'R-107');
});

test('editing cannot replace the existing ID', async () => {
  const f = fixture([{ SelectType: 'Issue', RaidLogID: 73 }]);
  const updated = await f.service.updateRaidItem(1, { raidLogId: '999', details: 'Changed' });
  assert.equal(updated.raidLogId, 'I-73');
  assert.equal(f.rows[0].RaidLogID, 73);
});

test('adding a Risk action stores the same logical ID without consuming a number', async () => {
  for (const stored of [105, 'R-105']) {
    const f = fixture([{ SelectType: 'Risk', RaidLogID: stored, RAIDId: 'old', TypeOfAction: 'Mitigation' }]);
    assert.equal(await f.service.updateRiskItemsByRaidId('old', {}, action, { ...action, type: 'Contingency' }), true);
    assert.equal(f.rows[0].RaidLogID, stored);
    assert.equal(f.writes[0].RaidLogID, 'R-105');
    assert.equal(f.writes[0].RAIDId, 'old');
    await f.service.createRiskItemWithActions({ type: 'Risk', raidId: 'new' }, action, null);
    assert.equal(f.writes[1].RaidLogID, 'R-106');
  }
});
