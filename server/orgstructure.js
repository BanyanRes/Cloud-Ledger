// ─────────────────────── Org structure (ownership tree) ───────────────────────
//
// Records who owns whom, from the legal org charts, and uses it to tie a fund's
// investment balance to the contributed capital of the operating company it
// actually funded.
//
// ── Why this needs its own node table instead of a parent column on `entities` ──
//
// The ownership chain runs through holding companies that hold the investment
// balances but keep no ledger in CloudLedger. For County Line Rail Fund I the
// legal chain (org chart, slide "County Line Rail Fund I, LP", 4/13/2026) is:
//
//   County Line Rail Fund I, LP        (entity 40 — has a ledger)
//     └── CLRFI Midco I, LLC           (entity 70 — has a ledger, but empty)
//           ├── County Line SRN, LLC          SHELL, no ledger
//           │     └── Sabine River & Northern Railroad LLC   (entity 37)
//           ├── CLRFI CLIP Sponsor, LLC       SHELL, no ledger
//           │     └── 76.51% County Line Industrial Park LLC (entity 42)
//           │           └── CLIP Property Owner LLC          (entity 54)
//           ├── CLRFI Silsbee Sponsor, LLC    SHELL, no ledger
//           │     └── 54.53% County Line Rail Silsbee LLC    (entity 52)
//           │           └── CLR Silsbee Property Owner LLC   (entity 39)
//           └── CLR Buna Property Owner LLC   (entity 38)
//     └── CLRFI Midco II, LLC          SHELL, no ledger (RRIF borrower)
//
// A `parent_entity_id` column could not express the shells at all, and dropping
// them would make CLRF look like the direct parent of four operating companies —
// losing exactly the 76.51% / 54.53% steps that create the non-controlling
// interest. So a node is either a CloudLedger entity (`entity_id` set) or a
// ledger-less shell (`entity_id` NULL), and edges carry the ownership percent.
//
// ── Look-through ──
//
// CLRF carries one investment account per operating company (121011 CLIP,
// 121021 Buna, 121031 SRN, 121041 Silsbee), and the operating company carries
// contributed capital naming a shell partway up the chain ("Contributed Capital
// – CLRFI – Midco 1", "…– CLRF I CLIP Sponsor LLC"). Neither side names the
// other directly, so a naive investor-to-counterparty match finds nothing. The
// reconciliation here walks the tree: capital contributed by ANY node on the
// path between the investor and the subsidiary counts as capital from that
// investor. Validated at 6/30/2026 — CLRF 121031 (60,408,356.37) against SRN's
// total contributed capital (60,408,356.26) ties to $0.11.

const { parseAccountName, listMappings } = require('./intercompany');

const NODE_TYPES = ['fund', 'holdco', 'company', 'property_owner', 'shell'];
const DEFAULT_TOLERANCE = 0.005;
// A tie-out that lands inside a dollar is rounding, not a break. Without this,
// CLRF's SRN investment reports an 11-cent "difference" on $60.4M and reads
// exactly like the $9.5M one two rows above it.
const DEFAULT_MATERIALITY = 1.00;

// ══════════════════════════════ Schema ══════════════════════════════

function ensureSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS org_nodes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      entity_id INTEGER,              -- NULL = shell with no CloudLedger ledger
      node_type TEXT NOT NULL DEFAULT 'company',
      notes TEXT,
      sort_order INTEGER NOT NULL DEFAULT 0,
      created_at TEXT, created_by TEXT, updated_at TEXT, updated_by TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_org_node_entity ON org_nodes(entity_id);

    CREATE TABLE IF NOT EXISTS org_edges (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      parent_node_id INTEGER NOT NULL,
      child_node_id INTEGER NOT NULL,
      ownership_pct REAL NOT NULL DEFAULT 100,
      notes TEXT,
      created_at TEXT, created_by TEXT,
      UNIQUE(parent_node_id, child_node_id)
    );
    CREATE INDEX IF NOT EXISTS idx_org_edge_parent ON org_edges(parent_node_id);
    CREATE INDEX IF NOT EXISTS idx_org_edge_child ON org_edges(child_node_id);
  `);
}

// ══════════════════════════════ Read ══════════════════════════════

function listStructure(db) {
  ensureSchema(db);
  const nodes = db.prepare(`
    SELECT n.*, e.name AS entity_name, e.code AS entity_code, e.entity_type
    FROM org_nodes n LEFT JOIN entities e ON e.id = n.entity_id
    ORDER BY n.sort_order, n.name COLLATE NOCASE`).all();
  const edges = db.prepare('SELECT * FROM org_edges').all();
  return { nodes, edges };
}

// Build the ownership tree under one root.
//
// `effective_pct` is the product of the ownership percentages along the path,
// so a 76.51% step anywhere above a node reduces everything below it. A node
// reachable by more than one path is visited once per path (that is what an
// org chart means), but a path that revisits a node already on it is cut — a
// cycle in the data must not hang the request.
function buildTree(nodes, edges, rootId) {
  const byId = new Map(nodes.map(n => [n.id, n]));
  const childrenOf = new Map();
  for (const e of edges) {
    if (!childrenOf.has(e.parent_node_id)) childrenOf.set(e.parent_node_id, []);
    childrenOf.get(e.parent_node_id).push(e);
  }
  const root = byId.get(rootId);
  if (!root) { const err = new Error('Root node not found'); err.status = 404; throw err; }

  const cycles = [];
  const walk = (nodeId, effPct, ancestry, depth) => {
    const n = byId.get(nodeId);
    if (!n) return null;
    const out = {
      ...n,
      effective_pct: effPct,
      nci_pct: 100 - effPct,
      depth,
      // Every node from the root down to (and including) this one — the chain
      // the look-through reconciliation walks.
      path: [...ancestry, nodeId],
      children: [],
    };
    const kids = (childrenOf.get(nodeId) || []).slice()
      .sort((a, b) => {
        const na = byId.get(a.child_node_id), nb = byId.get(b.child_node_id);
        return (na ? na.sort_order : 0) - (nb ? nb.sort_order : 0)
          || String(na && na.name).localeCompare(String(nb && nb.name));
      });
    for (const e of kids) {
      if (ancestry.includes(e.child_node_id) || e.child_node_id === nodeId) {
        cycles.push({ parent_node_id: nodeId, child_node_id: e.child_node_id });
        continue;
      }
      const pct = Number(e.ownership_pct);
      const childPct = effPct * (isFinite(pct) ? pct : 100) / 100;
      const sub = walk(e.child_node_id, childPct, [...ancestry, nodeId], depth + 1);
      if (sub) { sub.ownership_pct = pct; sub.edge_id = e.id; sub.edge_notes = e.notes || null; out.children.push(sub); }
    }
    return out;
  };
  const tree = walk(rootId, 100, [], 0);
  return { tree, cycles };
}

// Depth-first list of every node in a tree, root first.
function flattenTree(tree) {
  const out = [];
  const rec = n => { out.push(n); (n.children || []).forEach(rec); };
  if (tree) rec(tree);
  return out;
}

// Roots = nodes that are nobody's child. Used to render the page without the
// caller having to know which node is the top of a structure.
function listRoots(db) {
  const { nodes, edges } = listStructure(db);
  const isChild = new Set(edges.map(e => e.child_node_id));
  return nodes.filter(n => !isChild.has(n.id));
}

// ══════════════════════════════ Write ══════════════════════════════

function saveNode(db, body, who) {
  ensureSchema(db);
  const now = new Date().toISOString();
  const type = NODE_TYPES.includes(body.node_type) ? body.node_type : 'company';
  const eid = body.entity_id ? Number(body.entity_id) : null;
  if (body.id) {
    db.prepare(`UPDATE org_nodes SET name=?, entity_id=?, node_type=?, notes=?, sort_order=?, updated_at=?, updated_by=? WHERE id=?`)
      .run(body.name, eid, type, body.notes || null, Number(body.sort_order) || 0, now, who || null, Number(body.id));
    return Number(body.id);
  }
  return db.prepare(`INSERT INTO org_nodes (name, entity_id, node_type, notes, sort_order, created_at, created_by, updated_at, updated_by)
    VALUES (?,?,?,?,?,?,?,?,?)`)
    .run(body.name, eid, type, body.notes || null, Number(body.sort_order) || 0, now, who || null, now, who || null).lastInsertRowid;
}

function deleteNode(db, id) {
  db.prepare('DELETE FROM org_edges WHERE parent_node_id = ? OR child_node_id = ?').run(id, id);
  db.prepare('DELETE FROM org_nodes WHERE id = ?').run(id);
}

function saveEdge(db, body, who) {
  ensureSchema(db);
  const p = Number(body.parent_node_id), c = Number(body.child_node_id);
  if (!p || !c) { const e = new Error('parent_node_id and child_node_id are required'); e.status = 400; throw e; }
  if (p === c) { const e = new Error('A node cannot own itself'); e.status = 400; throw e; }
  // Refuse an edge that would close a loop, rather than storing it and relying
  // on the renderer's cycle guard to paper over it.
  if (isAncestor(db, c, p)) {
    const e = new Error('That would create a circular ownership chain'); e.status = 400; throw e;
  }
  const pct = body.ownership_pct == null ? 100 : Number(body.ownership_pct);
  if (!isFinite(pct) || pct < 0 || pct > 100) { const e = new Error('Ownership % must be between 0 and 100'); e.status = 400; throw e; }
  const now = new Date().toISOString();
  if (body.id) {
    db.prepare('UPDATE org_edges SET parent_node_id=?, child_node_id=?, ownership_pct=?, notes=? WHERE id=?')
      .run(p, c, pct, body.notes || null, Number(body.id));
    return Number(body.id);
  }
  return db.prepare('INSERT INTO org_edges (parent_node_id, child_node_id, ownership_pct, notes, created_at, created_by) VALUES (?,?,?,?,?,?)')
    .run(p, c, pct, body.notes || null, now, who || null).lastInsertRowid;
}

function isAncestor(db, maybeAncestorId, nodeId) {
  const edges = db.prepare('SELECT parent_node_id, child_node_id FROM org_edges').all();
  const parentsOf = new Map();
  for (const e of edges) {
    if (!parentsOf.has(e.child_node_id)) parentsOf.set(e.child_node_id, []);
    parentsOf.get(e.child_node_id).push(e.parent_node_id);
  }
  const seen = new Set(); const stack = [nodeId];
  while (stack.length) {
    const cur = stack.pop();
    if (seen.has(cur)) continue;
    seen.add(cur);
    for (const p of (parentsOf.get(cur) || [])) {
      if (p === maybeAncestorId) return true;
      stack.push(p);
    }
  }
  return false;
}

function deleteEdge(db, id) { db.prepare('DELETE FROM org_edges WHERE id = ?').run(id); }

// ═════════════ Fund investment ↔ subsidiary capital (look-through) ═════════════
//
// For each investment account on a ledger entity in the tree, find the
// subsidiary it names and compare it to that subsidiary's contributed capital.
// Two comparisons are reported side by side, because in the real data they
// disagree and the difference is informative:
//
//   chain_capital  — capital contributed by the investor OR by any shell on the
//                    path between the investor and the subsidiary. This is the
//                    strict answer: it counts only capital that names this
//                    ownership chain.
//   total_capital  — every contributed-capital account on the subsidiary,
//                    whoever it names. This is what ties for SRN, because the
//                    sponsor's waived development fee was booked to capital
//                    without naming the fund.

function reconcileFundInvestments(db, rootNodeId, opts) {
  ensureSchema(db);
  const tol = Number(opts.tolerance) > 0 ? Number(opts.tolerance) : DEFAULT_TOLERANCE;
  const mat = Number(opts.materiality) > 0 ? Number(opts.materiality) : DEFAULT_MATERIALITY;
  const { computeBalances, as_of } = opts;
  const { nodes, edges } = listStructure(db);
  const { tree, cycles } = buildTree(nodes, edges, rootNodeId);
  const all = flattenTree(tree);
  const byNodeId = new Map(all.map(n => [n.id, n]));

  // Ledger entities present in this tree, and where each sits.
  const nodeForEntity = new Map();
  for (const n of all) if (n.entity_id != null && !nodeForEntity.has(n.entity_id)) nodeForEntity.set(n.entity_id, n);
  const ledgerEntityIds = [...nodeForEntity.keys()];

  // Balances for every ledger entity in the tree, once.
  const balances = new Map();
  for (const eid of ledgerEntityIds) {
    const rows = computeBalances(eid, as_of ? { as_of } : {});
    balances.set(eid, rows);
  }
  const balMap = new Map([...balances].map(([eid, rows]) => [eid, new Map(rows.map(r => [String(r.code), r]))]));
  const amountOf = (eid, code) => {
    const r = (balMap.get(eid) || new Map()).get(String(code));
    return r ? Number(r.balance) || 0 : 0;
  };

  // IC mappings for those entities (the investment side is mapped; the capital
  // side is read straight off the chart of accounts so an unmapped capital
  // account still counts).
  const mappings = [];
  for (const eid of ledgerEntityIds) mappings.push(...listMappings(db, { entity_id: eid }));

  // Every contributed-capital account on a ledger entity, with the counterparty
  // its mapping names (or, when unmapped, the label parsed from its name).
  const capitalByEntity = new Map();
  for (const eid of ledgerEntityIds) {
    const rows = (balances.get(eid) || []).filter(r => {
      const p = parseAccountName(r.name);
      return p && p.ic_type === 'contributed_capital';
    });
    const mapped = new Map(mappings.filter(m => m.entity_id === eid).map(m => [String(m.account_code), m]));
    capitalByEntity.set(eid, rows.map(r => {
      const m = mapped.get(String(r.code));
      const parsed = parseAccountName(r.name) || {};
      return {
        entity_id: eid, account_code: String(r.code), account_name: r.name,
        amount: Number(r.balance) || 0,
        counterparty_entity_id: m ? m.counterparty_entity_id : null,
        counterparty_label: parsed.label || null,
        mapped: !!m,
      };
    }));
  }

  const hasCapital = nodeId => {
    const n = byNodeId.get(nodeId);
    if (!n || n.entity_id == null) return false;
    return (capitalByEntity.get(n.entity_id) || []).some(c => Math.abs(c.amount) > tol);
  };

  // An investment account names the operating company at the BOTTOM of the
  // chain ("CLIP - Investment Purchase"), but the balance it holds faces the
  // company one accounting layer down — the first entity below the investor
  // that actually carries contributed capital. Skipping over ledger entities
  // with no capital matters: CLRFI Midco I is a real CloudLedger entity with a
  // completely empty ledger, so stopping there would compare $60M of investment
  // against nothing. Comparing at the wrong layer double-counts, because each
  // layer pushes the layer above it down again — CLR Silsbee Property Owner
  // carries BOTH the Midco capital (6,387,181.23) and the CLR Silsbee LLC
  // capital (11,760,052.36) that represents the same money.
  const comparisonNodeFor = (investorNode, targetNode) => {
    const i = targetNode.path.indexOf(investorNode.id);
    if (i < 0) return targetNode;
    const below = targetNode.path.slice(i + 1);
    for (const id of below) if (hasCapital(id)) return byNodeId.get(id);
    return targetNode;
  };

  const rows = [];
  for (const m of mappings) {
    if (m.ic_type !== 'investment') continue;
    const investorNode = nodeForEntity.get(m.entity_id);
    if (!investorNode) continue;
    const amount = amountOf(m.entity_id, m.account_code);
    const targetNode = m.counterparty_entity_id != null ? nodeForEntity.get(Number(m.counterparty_entity_id)) : null;

    // Self-referential investments are the IC Reconciliation's job, not this
    // report's — skip them here so they aren't double-reported.
    if (m.counterparty_entity_id != null && Number(m.counterparty_entity_id) === Number(m.entity_id)) continue;

    if (!targetNode) {
      rows.push({
        investor_entity_id: m.entity_id, investor_name: investorNode.name,
        account_code: m.account_code, account_name: m.account_name, investment: amount,
        named_entity_id: m.counterparty_entity_id || null, named_name: null,
        subsidiary_entity_id: null, subsidiary_name: null, status: 'not_in_tree',
        message: m.counterparty_entity_id
          ? 'The subsidiary this account names is not in this ownership structure.'
          : 'This investment account has no counterparty mapped.',
        chain: [], chain_capital: 0, total_capital: 0,
        chain_difference: amount, total_difference: amount,
        chain_legs: [], other_legs: [], effective_pct: null, nci_pct: null, retargeted: false,
      });
      continue;
    }

    const cmpNode = comparisonNodeFor(investorNode, targetNode);
    const retargeted = cmpNode.id !== targetNode.id;

    // The chain: every node from the investor down to the comparison company.
    // Capital naming any node ABOVE the comparison company on that chain is
    // this investor's capital, looked through the shells.
    const tIdx = cmpNode.path.indexOf(investorNode.id);
    const inChain = tIdx >= 0 ? cmpNode.path.slice(tIdx) : [investorNode.id, cmpNode.id];
    const chainNodes = inChain.map(id => byNodeId.get(id)).filter(Boolean);
    // The comparison company itself is excluded — capital it contributed to
    // itself is a gross-up, not funding from above.
    const chainEntityIds = new Set(
      chainNodes.filter(n => n.entity_id != null && n.id !== cmpNode.id).map(n => n.entity_id));

    const caps = capitalByEntity.get(cmpNode.entity_id) || [];
    const chainLegs = caps.filter(c => c.counterparty_entity_id != null && chainEntityIds.has(Number(c.counterparty_entity_id)));
    const otherLegs = caps.filter(c => !chainLegs.includes(c));
    const chainCapital = chainLegs.reduce((s, c) => s + c.amount, 0);
    const totalCapital = caps.reduce((s, c) => s + c.amount, 0);
    const chainDiff = amount - chainCapital;
    const totalDiff = amount - totalCapital;
    const best = Math.min(Math.abs(chainDiff), Math.abs(totalDiff));

    rows.push({
      investor_entity_id: m.entity_id, investor_name: investorNode.name,
      account_code: m.account_code, account_name: m.account_name, investment: amount,
      // What the account name says vs. the layer actually compared.
      named_entity_id: targetNode.entity_id, named_name: targetNode.name,
      subsidiary_entity_id: cmpNode.entity_id, subsidiary_name: cmpNode.name,
      retargeted,
      retarget_note: retargeted
        ? 'The account names ' + targetNode.name + ', but ' + cmpNode.name +
          ' sits between it and ' + investorNode.name + ' and is the layer that carries the capital. Compared there to avoid counting the same money twice.'
        : null,
      effective_pct: round4(cmpNode.effective_pct), nci_pct: round4(cmpNode.nci_pct),
      // The shells between investor and subsidiary — the reason a direct match fails.
      chain: chainNodes.map(n => ({ id: n.id, name: n.name, entity_id: n.entity_id, is_shell: n.entity_id == null, ownership_pct: n.ownership_pct == null ? 100 : n.ownership_pct })),
      chain_capital: round2(chainCapital), total_capital: round2(totalCapital),
      chain_difference: round2(chainDiff), total_difference: round2(totalDiff),
      chain_legs: chainLegs, other_legs: otherLegs,
      status: best < tol ? (Math.abs(chainDiff) < tol ? 'ties_on_chain' : 'ties_on_total')
        : best < mat ? 'ties_rounding'
        : 'difference',
      tie_basis: Math.abs(chainDiff) <= Math.abs(totalDiff) ? 'chain' : 'total',
      best_difference: round2(Math.abs(chainDiff) <= Math.abs(totalDiff) ? chainDiff : totalDiff),
    });
  }
  rows.sort((a, b) => Math.abs(b.investment) - Math.abs(a.investment));

  // Non-controlling interest, computed from GL equity rather than from an
  // ownership schedule. Reported as indicative — see the caveats below.
  //
  // Only nodes whose OWN edge is below 100% are counted. Every entity beneath a
  // diluted step inherits the same effective %, so counting them all would add
  // the same minority twice: CLIP Property Owner's equity is County Line
  // Industrial Park's equity pushed down, and charging 23.49% to both would
  // report roughly double the real minority.
  const nci = all
    .filter(n => n.entity_id != null && n.ownership_pct != null && n.ownership_pct < 100 - 0.00005)
    .map(n => {
      const equity = (balances.get(n.entity_id) || [])
        .filter(r => r.type === 'Equity').reduce((s, r) => s + (Number(r.balance) || 0), 0);
      return {
        node_id: n.id, entity_id: n.entity_id, name: n.name,
        ownership_pct: round4(n.ownership_pct),
        effective_pct: round4(n.effective_pct), nci_pct: round4(n.nci_pct),
        entity_equity: round2(equity),
        nci_equity: round2(equity * n.nci_pct / 100),
      };
    })
    .sort((a, b) => Math.abs(b.nci_equity) - Math.abs(a.nci_equity));

  return {
    root: { id: tree.id, name: tree.name },
    as_of: as_of || null,
    tolerance: tol,
    materiality: mat,
    tree,
    cycles,
    rows,
    nci,
    totals: {
      investment_total: round2(rows.reduce((s, r) => s + r.investment, 0)),
      chain_capital_total: round2(rows.reduce((s, r) => s + r.chain_capital, 0)),
      total_capital_total: round2(rows.reduce((s, r) => s + r.total_capital, 0)),
      ties: rows.filter(r => r.status !== 'difference' && r.status !== 'not_in_tree').length,
      rounding_ties: rows.filter(r => r.status === 'ties_rounding').length,
      differences: rows.filter(r => r.status === 'difference' || r.status === 'not_in_tree').length,
      // Sum of each row's BEST difference, so a row that ties on one basis
      // does not inflate the headline via the other.
      abs_difference: round2(rows.reduce((s, r) => s + Math.abs(r.best_difference == null ? r.total_difference : r.best_difference), 0)),
      nci_equity_total: round2(nci.reduce((s, n) => s + n.nci_equity, 0)),
      shell_count: all.filter(n => n.entity_id == null).length,
      ledger_count: ledgerEntityIds.length,
    },
  };
}

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function round4(n) { return Math.round((Number(n) || 0) * 10000) / 10000; }

// ══════════════════════════ Seed: County Line Rail Fund I ══════════════════════════
//
// From the 4/13/2026 org chart. Entities resolve by code, so a missing entity
// becomes a shell node rather than failing the whole seed. Idempotent: it does
// nothing if a node with the root's name already exists.

const CLRF_CHART = {
  root: { name: 'County Line Rail Fund I, LP', code: 'COUNTYLI1', node_type: 'fund' },
  children: [
    {
      name: 'CLRFI Midco I, LLC', code: 'CLRFIMID', node_type: 'holdco', pct: 100,
      notes: 'Bank of Texas borrower',
      children: [
        {
          name: 'County Line SRN, LLC', node_type: 'shell', pct: 100,
          notes: 'Holds the SRN investment; no CloudLedger ledger',
          children: [{ name: 'Sabine River & Northern Railroad LLC', code: 'SABINERI', node_type: 'property_owner', pct: 100 }],
        },
        {
          name: 'CLRFI CLIP Sponsor, LLC', node_type: 'shell', pct: 100,
          notes: 'Holds the CLIP investment; no CloudLedger ledger',
          children: [{
            name: 'County Line Industrial Park LLC', code: 'COUNTYLI2', node_type: 'company', pct: 76.51,
            children: [{ name: 'CLIP Property Owner LLC', code: 'CLIPPROP', node_type: 'property_owner', pct: 100 }],
          }],
        },
        {
          name: 'CLRFI Silsbee Sponsor, LLC', node_type: 'shell', pct: 100,
          notes: 'Holds the Silsbee investment; no CloudLedger ledger',
          children: [{
            name: 'County Line Rail Silsbee LLC', code: 'COUNTYLI5', node_type: 'company', pct: 54.53,
            children: [{ name: 'CLR Silsbee Property Owner LLC', code: 'CLRSILSB2', node_type: 'property_owner', pct: 100 }],
          }],
        },
        { name: 'CLR Buna Property Owner LLC', code: 'CLRBUNAP', node_type: 'property_owner', pct: 100 },
      ],
    },
    { name: 'CLRFI Midco II, LLC', node_type: 'shell', pct: 100, notes: 'RRIF borrower; no CloudLedger ledger' },
  ],
};

function seedChart(db, chart, who) {
  ensureSchema(db);
  const existing = db.prepare('SELECT id FROM org_nodes WHERE name = ?').get(chart.root.name);
  if (existing) return { created: false, root_node_id: existing.id, message: 'Already seeded' };

  const entByCode = new Map(db.prepare('SELECT id, code FROM entities').all().map(e => [e.code, e.id]));
  let order = 0;
  const created = [];

  const add = (spec, parentId, pct) => {
    const eid = spec.code ? (entByCode.get(spec.code) || null) : null;
    const notes = spec.notes || (spec.code && !eid ? 'Entity code ' + spec.code + ' not found in CloudLedger' : null);
    const id = saveNode(db, {
      name: spec.name, entity_id: eid,
      node_type: eid ? spec.node_type : 'shell',
      notes, sort_order: order++,
    }, who);
    created.push({ id, name: spec.name, entity_id: eid });
    if (parentId) saveEdge(db, { parent_node_id: parentId, child_node_id: id, ownership_pct: pct == null ? 100 : pct }, who);
    for (const c of (spec.children || [])) add(c, id, c.pct);
    return id;
  };

  // The root's children live on `chart.children`, not inside `chart.root` —
  // graft them on so one recursive walk covers the whole chart.
  const rootId = db.transaction(() => add({ ...chart.root, children: chart.children || [] }, null, null))();
  return { created: true, root_node_id: rootId, nodes: created.length, detail: created };
}

// ══════════════════════════════ Routes ══════════════════════════════

function registerOrgStructureRoutes(app, deps) {
  const { db, auth, requireRole, computeBalances, userHasEntityAccess } = deps;
  ensureSchema(db);

  const who = req => (req.user && (req.user.name || req.user.email)) || null;
  const gate = [auth, requireRole('Admin', 'Accountant')];
  const fail = (res, e) => res.status(e.status || 500).json({ error: e.message });

  // The tree spans many entities, so access is checked per ledger entity in it.
  const assertTreeAccess = (req, nodeList) => {
    for (const n of nodeList) {
      if (n.entity_id != null && !userHasEntityAccess(req.user.id, req.user.role, n.entity_id)) {
        const e = new Error('No access to entity ' + n.entity_id); e.status = 403; throw e;
      }
    }
  };

  app.get('/api/org-structure', ...gate, (req, res) => {
    try {
      const { nodes, edges } = listStructure(db);
      assertTreeAccess(req, nodes);
      res.json({ nodes, edges, roots: listRoots(db).map(r => ({ id: r.id, name: r.name })) });
    } catch (e) { fail(res, e); }
  });

  app.get('/api/org-structure/tree', ...gate, (req, res) => {
    try {
      const rootId = Number(req.query.root_node_id);
      if (!rootId) return res.status(400).json({ error: 'root_node_id is required' });
      const { nodes, edges } = listStructure(db);
      const built = buildTree(nodes, edges, rootId);
      assertTreeAccess(req, flattenTree(built.tree));
      res.json(built);
    } catch (e) { fail(res, e); }
  });

  app.post('/api/org-structure/nodes', ...gate, (req, res) => {
    try {
      if (!req.body || !req.body.name) return res.status(400).json({ error: 'Node name is required' });
      if (req.body.entity_id) assertTreeAccess(req, [{ entity_id: Number(req.body.entity_id) }]);
      res.json({ id: saveNode(db, req.body, who(req)), success: true });
    } catch (e) { fail(res, e); }
  });

  app.put('/api/org-structure/nodes/:id', ...gate, (req, res) => {
    try {
      if (!req.body || !req.body.name) return res.status(400).json({ error: 'Node name is required' });
      if (req.body.entity_id) assertTreeAccess(req, [{ entity_id: Number(req.body.entity_id) }]);
      saveNode(db, { ...req.body, id: Number(req.params.id) }, who(req));
      res.json({ success: true });
    } catch (e) { fail(res, e); }
  });

  app.delete('/api/org-structure/nodes/:id', ...gate, (req, res) => {
    try { deleteNode(db, Number(req.params.id)); res.json({ success: true }); }
    catch (e) { fail(res, e); }
  });

  app.post('/api/org-structure/edges', ...gate, (req, res) => {
    try { res.json({ id: saveEdge(db, req.body || {}, who(req)), success: true }); }
    catch (e) { fail(res, e); }
  });

  app.put('/api/org-structure/edges/:id', ...gate, (req, res) => {
    try { saveEdge(db, { ...(req.body || {}), id: Number(req.params.id) }, who(req)); res.json({ success: true }); }
    catch (e) { fail(res, e); }
  });

  app.delete('/api/org-structure/edges/:id', ...gate, (req, res) => {
    try { deleteEdge(db, Number(req.params.id)); res.json({ success: true }); }
    catch (e) { fail(res, e); }
  });

  // Load the County Line Rail Fund I chart. Idempotent.
  app.post('/api/org-structure/seed/clrf', ...gate, (req, res) => {
    try { res.json(seedChart(db, CLRF_CHART, who(req))); }
    catch (e) { fail(res, e); }
  });

  app.get('/api/org-structure/reconcile/investments', ...gate, (req, res) => {
    try {
      const rootId = Number(req.query.root_node_id);
      if (!rootId) return res.status(400).json({ error: 'root_node_id is required' });
      const as_of = /^\d{4}-\d{2}-\d{2}$/.test(String(req.query.as_of || '')) ? String(req.query.as_of) : null;
      const { nodes, edges } = listStructure(db);
      assertTreeAccess(req, flattenTree(buildTree(nodes, edges, rootId).tree));
      res.json(reconcileFundInvestments(db, rootId, { computeBalances, as_of, tolerance: req.query.tolerance, materiality: req.query.materiality }));
    } catch (e) { fail(res, e); }
  });
}

module.exports = {
  registerOrgStructureRoutes,
  ensureSchema,
  listStructure,
  listRoots,
  buildTree,
  flattenTree,
  reconcileFundInvestments,
  seedChart,
  CLRF_CHART,
  NODE_TYPES,
};
