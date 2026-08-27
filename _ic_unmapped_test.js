// Fixture test for listUnmappedAccounts. The module takes `db` and
// `computeBalances` injected, so this runs without building better-sqlite3.
//
// Run twice: once with the selected entity a normal 'accounting' entity, once
// with it marked entity_type 'shell'.
const ic = require('./server/intercompany.js');

const ENTITIES = [
  { id: 39, name: 'CLR Silsbee Property Owner LLC', code: 'CLRSILSB2', entity_type: 'accounting' },
  { id: 40, name: 'CLIP Property Owner LLC',        code: 'CLIPPROP',  entity_type: 'accounting' },
  { id: 52, name: 'County Line Rail Silsbee LLC',   code: 'COUNTYLI5', entity_type: 'accounting' },
];
const ACCOUNTS = [ // entity 39's chart of accounts
  { code: '11010', name: 'Loan receivable - CLIP',                      type: 'Asset' },     // not one of the four kinds
  { code: '13000', name: 'Prepaid insurance',                           type: 'Asset' },     // ordinary balance sheet
  { code: '17001', name: 'Investment - CLR Silsbee Property Owner LLC', type: 'Asset' },
  { code: '18378', name: 'Due from CLIP Property Owner',                type: 'Asset' },     // already mapped
  { code: '18400', name: 'Due from outside vendor',                     type: 'Asset' },     // matcher guesses external
  { code: '18500', name: 'Due to Third Party Co',                       type: 'Liability' }, // ALREADY mapped as external
  { code: '19999', name: 'Due from Nowhere Holdings LLC',               type: 'Asset' },     // no entity, no company
  { code: '23375', name: 'Due to CLIP Property Owner',                  type: 'Liability' },
  { code: '30100', name: 'Contributed Capital - John H. Grayson Jr.',   type: 'Equity' },    // a person
  { code: '30200', name: 'Contributed Capital - County Line Rail Silsbee LLC', type: 'Equity' },
  { code: '30300', name: 'Contributed Capital - Charing Cross Partners LP',    type: 'Equity' },   // investor, not in CL
  { code: '41500', name: 'Management fee income - CLIP',                type: 'Revenue' },   // P&L
  { code: '52200', name: 'Interest expense - Due to Buna',              type: '' },          // P&L by code range
  { code: '60100', name: 'Repairs and maintenance',                     type: 'Expense' },   // P&L
];
// 18500 is mapped WITH is_external = 1. The unmapped list must not care about
// the flag: a mapped account is mapped, and marking one external is exactly how
// a person takes it off this list for good.
const MAPPED = [{ account_code: '18378', is_external: 0 },
                { account_code: '18500', is_external: 1 }];
const BALANCES = { '11010': 415000, '13000': 8200, '17001': 11760052.36, '23375': -202420.88,
                   '18400': 61250, '18500': -33000, '19999': 0, '30100': -500000,
                   '30200': -11760052.36, '30300': -2000000, '41500': 90000, '52200': 4300, '60100': 12.5 };

function makeDb(selectedEntityType) {
  const ents = ENTITIES.map(e => e.id === 39 ? { ...e, entity_type: selectedEntityType } : e);
  return {
    exec() {},
    prepare(sql) {
      const s = sql.replace(/\s+/g, ' ');
      return {
        all: () => {
          if (/FROM entities/.test(s)) return ents;
          if (/FROM accounts/.test(s)) return ACCOUNTS;
          if (/FROM intercompany_accounts WHERE entity_id/.test(s)) return MAPPED;
          if (/PRAGMA table_info/.test(s)) return [{ name: 'counterparty_node_id' }];
          return [];
        },
        get: (id) => {
          // org_nodes is absent here, which is the degrade-not-throw path.
          if (/FROM org_nodes/.test(s)) throw new Error('no such table: org_nodes');
          if (/FROM entities WHERE id/.test(s)) return ents.find(e => e.id === Number(id));
          return undefined;
        },
        run: () => ({}),
      };
    },
  };
}
const computeBalances = () => Object.entries(BALANCES).map(([code, balance]) => ({ code, balance }));

let fail = 0;
const ok = (cond, label) => { console.log((cond ? '  OK   ' : '  FAIL ') + label); if (!cond) fail++; };

// ───────────────────────── normal entity ─────────────────────────
console.log('\nentity_type = accounting');
const r = ic.listUnmappedAccounts(makeDb('accounting'), 39, { computeBalances, as_of: '2026-06-30' });
const by = c => r.accounts.find(a => a.account_code === c);

ok(r.count === 6, 'zero-balance and non-reconcilable kinds dropped — count ' + r.count);
ok(!by('18378'), '18378 does not appear (already mapped)');

// External counterparties. Every due from / due to account with a balance stays
// AVAILABLE for mapping, whatever the matcher guesses; marking one external is
// how it leaves this list, and it leaves for good.
ok(!!by('18400') && by('18400').ic_type === 'due_from',
   '"Due from outside vendor" IS listed — a due-from with a balance is always mappable');
ok(by('18400').is_external === 1 && by('18400').confidence === 'external',
   '...and arrives with external pre-selected, so it is one click');
ok(by('18400').balance === 61250, '...carrying its balance');
ok(!by('18500'), '18500, already mapped as external, is NOT listed');
ok(r.skipped_other === 5, 'and it was not counted as "some other kind" either');

// Not one of the four kinds. These used to be listed; they are the change.
ok(!by('11010'), '"Loan receivable - CLIP" is NOT listed — not a kind the recon matches');
ok(!by('13000'), '"Prepaid insurance" is NOT listed');
ok(!by('60100') && !by('41500') && !by('52200'), 'no P&L account is listed');
ok(r.skipped_other === 5, 'skipped_other counts all five — got ' + r.skipped_other);
ok(r.shell_entity === false, 'not flagged a shell entity');
ok(r.skipped_shell_capital === 0, 'nothing skipped as shell capital');

// The four kinds, all present.
ok(!!by('17001') && by('17001').ic_type === 'investment', '17001 investment listed');
ok(!by('19999'), '19999 due_from carries a ZERO balance, so it is NOT listed');
ok(r.skipped_zero_balance === 1, 'skipped_zero_balance reports it — got ' + r.skipped_zero_balance);
ok(!!by('23375') && by('23375').ic_type === 'due_to', '23375 due_to listed');
ok(!!by('30100') && by('30100').ic_type === 'contributed_capital', '30100 contributed capital listed');
ok(!!by('30200') && by('30200').ic_type === 'contributed_capital', '30200 contributed capital listed');

ok(by('30100').individual === true, 'the capital account naming a person is flagged individual');
ok(by('23375').counterparty_entity_id === 40, '23375 resolves to CLIP Property Owner');
ok(by('17001').counterparty_entity_id === 39, '17001 still resolves to the entity itself');
ok(by('30200').counterparty_entity_id === 52, '30200 resolves to County Line Rail Silsbee');
ok(by('18400').can_register === false, 'an external-looking label is not offered for registration');
ok(by('23375').balance === -202420.88, 'balances come through: 23375 = -202,420.88');
ok(r.accounts[0].account_code === '17001' && r.accounts[1].account_code === '30200',
   'sorted by the money each account carries');
ok(!('looks_ic' in r.accounts[0]), 'looks_ic is gone — every row reads as intercompany now');

// Pre-answered rows. In this fixture the counterparty's balances carry no type,
// so no existing account qualifies and every entity-facing row gets a DRAFT.
ok(!!by('23375').suggested_new && by('23375').suggested_new.ic_type === 'due_from'
   && by('23375').suggested_new.type === 'Asset'
   && by('23375').suggested_new.account_name === 'Due from CLR Silsbee Property Owner LLC',
   'our due_to drafts their mirrored due_from, named after US');
ok(!!by('30200').suggested_new && by('30200').suggested_new.ic_type === 'investment'
   && by('30200').suggested_new.account_name === 'Investment in CLR Silsbee Property Owner LLC',
   'our contributed capital drafts their investment');
ok(!by('18400').suggested_new && !by('18400').suggested_existing,
   'an external row is never pre-answered — nothing to map to');
ok(by('23375').source === 'account' && by('23375').mapping_id === null,
   'rows carry their source');

// Investor capital: the contributor is not set up in CL, so no account is
// needed and no suggestion is drafted.
ok(!!by('30300') && by('30300').investor_capital === true,
   '"Contributed Capital - Charing Cross Partners LP" is investor capital');
ok(!by('30300').suggested_new && !by('30300').suggested_existing,
   '...and carries no account suggestion');
ok(!by('30100').investor_capital,
   'an INDIVIDUAL capital account is not investor capital — it keeps its own handling');
ok(by('30200').suggested_new && by('30200').suggested_new.ic_type === 'investment',
   'capital whose contributor IS a CL entity still drafts/pre-fills the investment account');

// ───────────────────────── shell entity ─────────────────────────
console.log('\nentity_type = shell');
const rs = ic.listUnmappedAccounts(makeDb('shell'), 39, { computeBalances, as_of: '2026-06-30' });
const bys = c => rs.accounts.find(a => a.account_code === c);

ok(rs.shell_entity === true, 'flagged a shell entity');
ok(!bys('30100') && !bys('30200'), 'neither contributed-capital account is listed on a shell');
ok(rs.skipped_shell_capital === 3, 'skipped_shell_capital reports all three — got ' + rs.skipped_shell_capital);
ok(rs.count === 3, 'the other three kinds survive — count ' + rs.count);
ok(!!bys('17001') && !!bys('23375') && !!bys('18400'),
   'due from / due to / investment are still listed on a shell');
ok(ic.isShellEntity({ entity_type: 'shell' }) && ic.isShellEntity({ entity_type: ' Shell ' })
   && !ic.isShellEntity({ entity_type: 'accounting' }) && !ic.isShellEntity({}) && !ic.isShellEntity(null),
   'isShellEntity: case/space tolerant, and a missing column is not a shell');

console.log(fail ? '\n' + fail + ' FAILED' : '\nall passed');
process.exit(fail ? 1 : 0);
