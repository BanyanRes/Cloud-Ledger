import sys
p = 'C:/Users/JimmyYun/Cloud-Ledger/server/index.js'
src = open(p, encoding='utf-8').read()

old = """// Link Turnkey project to CL entity + seed POC chart of accounts
app.post('/api/turnkey/projects/link', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  try {
    const { turnkey_project_id, project_code, project_name } = req.body;
    if (!turnkey_project_id || !project_code || !project_name) {
      return res.status(400).json({ error: 'turnkey_project_id, project_code, project_name required' });
    }
    const map = turnkey.linkProject(db, { turnkey_project_id, project_code, project_name });
    res.json({ ok: true, project_map: map });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});"""

new = """// Register a Turnkey project as a job dimension on the company entity.
// Body: { turnkey_project_id, project_code, project_name,
//         contract_amount?, total_estimated_costs? }
// Update-on-conflict so this can also be called to refresh contract/estimate.
app.post('/api/turnkey/projects/link', turnkeyAuth, turnkey.requireScope('turnkey:sync'), (req, res) => {
  try {
    const { turnkey_project_id, project_code, project_name,
            contract_amount, total_estimated_costs } = req.body;
    if (!turnkey_project_id || !project_code || !project_name) {
      return res.status(400).json({ error: 'turnkey_project_id, project_code, project_name required' });
    }
    const map = turnkey.linkProject(db, {
      turnkey_project_id, project_code, project_name,
      contract_amount, total_estimated_costs,
    });
    res.json({ ok: true, project_map: map });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});"""

assert old in src
src = src.replace(old, new, 1)
open(p, 'w', encoding='utf-8').write(src)
print('OK')
