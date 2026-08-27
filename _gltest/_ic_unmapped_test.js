const ic = require('./server/intercompany.js');

// Minimal db stub: prepare(sql).all(...args) / .get(...) / .run(...), plus exec.
const ENTITIES = [
  { id: 39, name: 'CLR Silsbee Property Owner LLC', code: 'CLRSILSB2' },
  { id: 40, name: 'CLIP Property Owner LLC',        code: 'CLIPPROP'  },
  { id: 52, name: 'County Line Rail Silsbee LLC',   code: 'COUNTYLI5' },
];
const ACCOUNTS = [ // entity 39's chart
  { code: '18378', name: 'Due from CLIP Property Owner',        type: 'Asset' },
  { code: '23375', name: 'Due to CLIP Property Owner',          type: 'Liability' },
  { code: '17001', name: 'Investment - CLR Silsbee Property Owner LLC', type: 'Asset' },
  { code: '11010', name: 'Loan receivable - CLIP',              type: 'Asset' },   // parser can't read it
  { code: '30100', name: 'Contributed Capital - John H. Grayson Jr.', type: 'Equity' }, // a person
  { code: '60100', name: 'Repairs and maintenance',             type: 'Expense' },
  { code: '19999', name: 'Due from Nowhere Holdings LLC',        type: 'Asset' },  // no entity, no company
];
const MAPPED = [{ account_code: '18378' }];   // only one is already mapped
const BALANCES = { '23375': -202420.88, '17001': 11760052.36, '11010': 415000, '19999': 0, '60100': 12.5 };

const db = {
  exec() {},
  prepare(sql) {
    const s = sql.replace(/\s+/g, ' ');
    return {
      all: (...a) => {
        if (/FROM entities/.test(s)) return ENTITIES;
        if (/FROM accounts/.test(s)) return ACCOUNTS;
        if (/FROM intercompany_accounts WHERE entity_id/.test(s)) return MAPPED;
        if (/PRAGMA table_info/.test(s)) return [{ name: 'counterparty_node_id' }];
        if (/FROM org_nodes/.test(s)) return [];
        return [];
      },
      get: () => { if (/FROM org_nodes/.test(s)) throw new Error('no org_nodes'); return undefined; },
      run: () => ({}),
    };
  },
};
const computeBalances = () => Object.entries(BALANCES).map(([code, balance]) => ({ code, balance }));

const r = ic.listUnmappedAccounts(db, 39, { computeBalances, as_of: '2026-06-30' });
const by = c => r.accounts.find(a => a.account_code === c);
let fail = 0;
const ok = (cond, label) => { console.log((cond ? '  OK   ' : '  FAIL ') + label); if (!cond) fail++; };

ok(r.count === 6, 'the one already-mapped account (18378) is excluded — count ' + r.count);
ok(!by('18378'), '18378 does not appear');
ok(by('11010') && by('11010').looks_ic === false,
   '"Loan receivable - CLIP" is listed even though the parser cannot read it');
ok(by('60100') && by('60100').looks_ic === false, 'a plain expense account is listed too');
ok(by('23375') && by('23375').looks_ic === true && by('23375').ic_type === 'due_to'
   && by('23375').counterparty_entity_id === 40,
   '23375 reads as due_to and resolves to CLIP Property Owner');
ok(by('17001') && by('17001').counterparty_entity_id === 39,
   '17001 still resolves to the entity itself (the known self-investment)');
ok(by('30100') && by('30100').individual === true, 'the capital account naming a person is flagged individual');
ok(by('19999') && by('19999').can_register === true,
   'an unresolvable company offers "register this company"');
ok(by('23375').balance === -202420.88, 'balances come through: 23375 = -202,420.88');
ok(r.accounts[0].looks_ic === true, 'recognised accounts sort first');
ok(r.accounts.findIndex(a => a.account_code === '17001') <
   r.accounts.findIndex(a => a.account_code === '23375'),
   'inside those, the bigger balance sorts first (17001 before 23375)');
ok(r.ic_like === 4, 'ic_like counts recognised non-individual accounts — got ' + r.ic_like);

console.log(fail ? '\n' + fail + ' FAILED' : '\nall passed');
process.exit(fail ? 1 : 0);
