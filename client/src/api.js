const API_BASE = '/api';
function getToken() { return localStorage.getItem('cl_token'); }
function setToken(t) { localStorage.setItem('cl_token', t); }
function clearToken() { localStorage.removeItem('cl_token'); }

async function request(path, options = {}) {
  const token = getToken();
  const headers = { ...(token ? { Authorization: 'Bearer ' + token } : {}), ...options.headers };
  if (!(options.body instanceof FormData)) headers['Content-Type'] = 'application/json';
  const res = await fetch(API_BASE + path, {
    ...options, headers,
    body: options.body instanceof FormData ? options.body : (options.body ? JSON.stringify(options.body) : undefined),
  });
  if (res.status === 401) { clearToken(); window.location.reload(); return null; }
  const data = await res.json();
  if (!res.ok) { const err = new Error(data.error || 'Request failed'); err.detail = data; throw err; }
  return data;
}

export const api = {
  // Auth
  login: (email, pw) => request('/auth/login', { method: 'POST', body: { email, password: pw } }),
  signup: (name, email, pw, role) => request('/auth/signup', { method: 'POST', body: { name, email, password: pw, role } }),
  me: () => request('/auth/me'),
  updateProfile: (name, email) => request('/auth/profile', { method: 'PUT', body: { name, email } }),
  changePassword: (cur, nw) => request('/auth/change-password', { method: 'POST', body: { current_password: cur, new_password: nw } }),
  forgotPassword: (email) => request('/auth/forgot-password', { method: 'POST', body: { email } }),
  resetPassword: (token, new_password) => request('/auth/reset-password', { method: 'POST', body: { token, new_password } }),
  adminResetPassword: (uid, pw) => request('/auth/admin-reset-password', { method: 'POST', body: { user_id: uid, new_password: pw } }),

  // Users
  getUsers: () => request('/users'),
  deleteUser: (id) => request('/users/' + id, { method: 'DELETE' }),
  updateUser: (id, data) => request('/users/' + id, { method: 'PUT', body: data }),
  getUserEntityAccess: (id) => request('/users/' + id + '/entity-access'),
  setUserEntityAccess: (id, entity_ids) => request('/users/' + id + '/entity-access', { method: 'PUT', body: { entity_ids } }),

  // User groups (bundle users + grant entity access to all at once, e.g. CLA)
  getGroups: () => request('/groups'),
  getGroup: (id) => request('/groups/' + id),
  setGroupMembers: (id, user_ids) => request('/groups/' + id + '/members', { method: 'PUT', body: { user_ids } }),
  setGroupEntities: (id, entity_ids) => request('/groups/' + id + '/entities', { method: 'PUT', body: { entity_ids } }),

  // Entities
  getEntities: () => request('/entities'),
  // Turnkey Rail WIP schedule (in-app report). JWT-authed admin endpoint.
  getTurnkeyWip: (asOf) => request('/admin/turnkey/wip-schedule' + (asOf ? ('?as_of=' + asOf) : '')),
  getTurnkeyProjects: () => request('/admin/turnkey/projects'),
  createEntity: (name, entity_type, display_id) => request('/entities', { method: 'POST', body: { name, ...(entity_type ? { entity_type } : {}), ...(display_id ? { display_id } : {}) } }),
  updateEntity: (id, data) => request('/entities/' + id, { method: 'PUT', body: data }),
  importTrialBalance: (eid, file, asOfDate) => {
    const fd = new FormData();
    fd.append('file', file);
    if (asOfDate) fd.append('as_of_date', asOfDate);
    return request('/entities/' + eid + '/import-tb', { method: 'POST', body: fd });
  },
  importGLPreview: (eid, file) => {
    const fd = new FormData();
    fd.append('file', file);
    return request('/entities/' + eid + '/import-gl/preview', { method: 'POST', body: fd });
  },
  importGL: (eid, file, mapping, asOfDate) => {
    const fd = new FormData();
    fd.append('file', file);
    fd.append('mapping', JSON.stringify(mapping));
    if (asOfDate) fd.append('as_of_date', asOfDate);
    return request('/entities/' + eid + '/import-gl', { method: 'POST', body: fd });
  },
  bulkCreateEntities: (ents) => request('/entities/bulk', { method: 'POST', body: { entities: ents } }),
  deleteEntity: (id) => request('/entities/' + id, { method: 'DELETE' }),

  // Accounts
  getAccounts: (eid) => request('/entities/' + eid + '/accounts'),
  createAccount: (eid, data) => request('/entities/' + eid + '/accounts', { method: 'POST', body: data }),
  updateAccount: (eid, code, data) => request('/entities/' + eid + '/accounts/' + encodeURIComponent(code), { method: 'PUT', body: data }),
  deleteAccount: (eid, code) => request('/entities/' + eid + '/accounts/' + encodeURIComponent(code), { method: 'DELETE' }),

  // Journal Entries
  getEntries: (eid, from, to) => {
    let q = '/entities/' + eid + '/entries'; const p = [];
    if (from) p.push('from=' + from); if (to) p.push('to=' + to);
    return request(q + (p.length ? '?' + p.join('&') : ''));
  },
  getEntry: (eid, id) => request('/entities/' + eid + '/entries/' + id),
  createEntry: (eid, data) => request('/entities/' + eid + '/entries', { method: 'POST', body: data }),
  updateEntry: (eid, id, data) => request('/entities/' + eid + '/entries/' + id, { method: 'PUT', body: data }),
  deleteEntry: (eid, id) => request('/entities/' + eid + '/entries/' + id, { method: 'DELETE' }),
  bulkEntriesPreview: (eid, file) => {
    const fd = new FormData(); fd.append('file', file);
    return request('/entities/' + eid + '/entries/bulk/preview', { method: 'POST', body: fd });
  },
  bulkEntriesCommit: (eid, entries) => request('/entities/' + eid + '/entries/bulk', { method: 'POST', body: { entries } }),

  // Attachments
  uploadAttachments: (eid, entryId, files) => {
    const fd = new FormData();
    for (const f of files) fd.append('files', f);
    return request('/entities/' + eid + '/entries/' + entryId + '/attachments', { method: 'POST', body: fd });
  },
  downloadAttachment: (id) => API_BASE + '/attachments/' + id + '/download?token=' + encodeURIComponent(getToken() || ''),
  deleteAttachment: (id) => request('/attachments/' + id, { method: 'DELETE' }),

  // Balances
  getBalances: (eid, opts = {}) => {
    const p = [];
    if (opts.as_of) p.push('as_of=' + opts.as_of);
    if (opts.from) p.push('from=' + opts.from);
    if (opts.to) p.push('to=' + opts.to);
    if (opts.close_pl_before) p.push('close_pl_before=' + opts.close_pl_before);
    if (opts.location_id) p.push('location_id=' + opts.location_id);
    if (opts.class_id) p.push('class_id=' + opts.class_id);
    return request('/entities/' + eid + '/balances' + (p.length ? '?' + p.join('&') : ''));
  },
  // Fund reporting (CLRF-style LP package)
  getFundInvestments: (eid) => request('/entities/' + eid + '/fund-investments'),
  createFundInvestment: (eid, data) => request('/entities/' + eid + '/fund-investments', { method: 'POST', body: data }),
  updateFundInvestment: (eid, id, data) => request('/entities/' + eid + '/fund-investments/' + id, { method: 'PATCH', body: data }),
  deleteFundInvestment: (eid, id) => request('/entities/' + eid + '/fund-investments/' + id, { method: 'DELETE' }),
  setClassPartnerType: (eid, id, partner_type) => request('/entities/' + eid + '/classes/' + id, { method: 'PATCH', body: { partner_type } }),
  setClassCommitment: (eid, classId, commitment_amount) => request('/entities/' + eid + '/commitments/by-class/' + classId, { method: 'PUT', body: { commitment_amount } }),
  getFundAllocation: (eid, asOf) => request('/entities/' + eid + '/fund-allocation' + (asOf ? ('?as_of=' + asOf) : '')),
  getFundStatementsPdf: async (eid, asOf) => {
    const token = getToken();
    const res = await fetch(API_BASE + '/entities/' + eid + '/fund-statements.pdf?as_of=' + asOf, { headers: token ? { Authorization: 'Bearer ' + token } : {} });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const ctype = res.headers.get('content-type') || '';
    if (!res.ok || ctype.includes('application/json')) { let d = {}; try { d = await res.json(); } catch {} throw new Error(d.error || 'Generate failed'); }
    const cd = res.headers.get('content-disposition') || '';
    const mm = cd.match(/filename="?([^"]+)"?/);
    const filename = mm ? mm[1] : 'Fund_Financial_Statements.pdf';
    const blob = await res.blob();
    return { blob, filename };
  },
  getTtmPL: (eid, asOf) => request('/entities/' + eid + '/ttm-pl?as_of=' + asOf),
  analyzeTtmPL: async (eid, asOf) => {
    const token = getToken();
    const res = await fetch(API_BASE + '/entities/' + eid + '/ttm-pl/analyze', {
      method: 'POST',
      headers: { 'content-type': 'application/json', ...(token ? { Authorization: 'Bearer ' + token } : {}) },
      body: JSON.stringify({ as_of: asOf }),
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Analysis failed');
    return data;
  },
  getTtmPLXlsx: async (eid, asOf, analysis) => {
    const token = getToken();
    const res = await fetch(API_BASE + '/entities/' + eid + '/ttm-pl.xlsx', {
      method: 'POST',
      headers: { 'content-type': 'application/json', ...(token ? { Authorization: 'Bearer ' + token } : {}) },
      body: JSON.stringify({ as_of: asOf, analysis: analysis || null }),
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const ctype = res.headers.get('content-type') || '';
    if (!res.ok || ctype.includes('application/json')) { let data = {}; try { data = await res.json(); } catch {} throw new Error(data.error || 'Export failed'); }
    const cd = res.headers.get('content-disposition') || '';
    const m = cd.match(/filename="?([^"]+)"?/);
    const filename = m ? m[1] : 'Trailing_12_Months.xlsx';
    const blob = await res.blob();
    return { blob, filename };
  },
  getGLDetail: (eid, opts = {}) => {
    const p = [];
    if (opts.from) p.push('from=' + opts.from);
    if (opts.to) p.push('to=' + opts.to);
    if (opts.location_id) p.push('location_id=' + opts.location_id);
    if (opts.class_id) p.push('class_id=' + opts.class_id);
    if (opts.project_id) p.push('project_id=' + opts.project_id);
    if (opts.account_code) p.push('account_code=' + encodeURIComponent(opts.account_code));
    return request('/entities/' + eid + '/gl-detail' + (p.length ? '?' + p.join('&') : ''));
  },
  getClasses: (eid) => request('/entities/' + eid + '/classes'),
  getCommitments: (eid) => request('/entities/' + eid + '/commitments'),
  createCommitment: (eid, body) => request('/entities/' + eid + '/commitments', { method: 'POST', body }),
  updateCommitment: (eid, id, body) => request('/entities/' + eid + '/commitments/' + id, { method: 'PATCH', body }),
  deleteCommitment: (eid, id) => request('/entities/' + eid + '/commitments/' + id, { method: 'DELETE' }),
  getMemorizedReports: (eid) => request('/entities/' + eid + '/memorized-reports'),
  createMemorizedReport: (eid, body) => request('/entities/' + eid + '/memorized-reports', { method: 'POST', body }),
  updateMemorizedReport: (eid, id, body) => request('/entities/' + eid + '/memorized-reports/' + id, { method: 'PATCH', body }),
  deleteMemorizedReport: (eid, id) => request('/entities/' + eid + '/memorized-reports/' + id, { method: 'DELETE' }),
  getLocations: (eid) => request('/entities/' + eid + '/locations'),
  getProjects: (eid) => request('/entities/' + eid + '/projects'),
  createLocation: (eid, data) => request('/entities/' + eid + '/locations', { method: 'POST', body: data }),
  updateLocation: (eid, id, data) => request('/entities/' + eid + '/locations/' + id, { method: 'PATCH', body: data }),
  deleteLocation: (eid, id) => request('/entities/' + eid + '/locations/' + id, { method: 'DELETE' }),
  createClass: (eid, data) => request('/entities/' + eid + '/classes', { method: 'POST', body: data }),
  updateClass: (eid, id, data) => request('/entities/' + eid + '/classes/' + id, { method: 'PATCH', body: data }),
  deleteClass: (eid, id) => request('/entities/' + eid + '/classes/' + id, { method: 'DELETE' }),
  createProject: (eid, data) => request('/entities/' + eid + '/projects', { method: 'POST', body: data }),
  updateProject: (eid, id, data) => request('/entities/' + eid + '/projects/' + id, { method: 'PATCH', body: data }),
  deleteProject: (eid, id) => request('/entities/' + eid + '/projects/' + id, { method: 'DELETE' }),
  bulkProjects: (eid, projects, applyAll) => request('/entities/' + eid + '/projects/bulk', { method: 'POST', body: { projects, apply_all: !!applyAll } }),
  setLocationKind: (eid, id, kind) => request('/entities/' + eid + '/locations/' + id, { method: 'PATCH', body: { kind } }),
  setClassKind: (eid, id, kind) => request('/entities/' + eid + '/classes/' + id, { method: 'PATCH', body: { kind } }),
  // ── Accounts Receivable: customers ──
  getArCustomers: (eid) => request('/entities/' + eid + '/ar/customers'),
  createArCustomer: (eid, data) => request('/entities/' + eid + '/ar/customers', { method: 'POST', body: data }),
  updateArCustomer: (eid, id, data) => request('/entities/' + eid + '/ar/customers/' + id, { method: 'PATCH', body: data }),
  deleteArCustomer: (eid, id) => request('/entities/' + eid + '/ar/customers/' + id, { method: 'DELETE' }),
  getDimensionBalances: (eid, opts = {}) => {
    const p = [];
    if (opts.dim) p.push('dim=' + opts.dim);
    if (opts.accounts) p.push('accounts=' + encodeURIComponent(opts.accounts));
    if (opts.account_prefix) p.push('account_prefix=' + encodeURIComponent(opts.account_prefix));
    if (opts.kind) p.push('kind=' + encodeURIComponent(opts.kind));
    if (opts.as_of) p.push('as_of=' + opts.as_of);
    return request('/entities/' + eid + '/dimension-balances' + (p.length ? '?' + p.join('&') : ''));
  },
  getPivot: (eid, opts = {}) => {
    const p = [];
    if (opts.dim) p.push('dim=' + opts.dim);
    if (opts.accounts) p.push('accounts=' + encodeURIComponent(opts.accounts));
    if (opts.account_prefix) p.push('account_prefix=' + encodeURIComponent(opts.account_prefix));
    if (opts.from) p.push('from=' + opts.from);
    if (opts.to) p.push('to=' + opts.to);
    if (opts.as_of) p.push('as_of=' + opts.as_of);
    return request('/entities/' + eid + '/pivot' + (p.length ? '?' + p.join('&') : ''));
  },
  getSummary: () => request('/summary'),

  // Bank Transactions
  getBankTransactions: (eid, bankAcct, status) => {
    const p = []; if (bankAcct) p.push('bank_account=' + bankAcct); if (status) p.push('status=' + status);
    return request('/entities/' + eid + '/bank-transactions' + (p.length ? '?' + p.join('&') : ''));
  },
  uploadBankTransactions: (eid, bankAcct, file) => {
    const fd = new FormData(); fd.append('file', file); fd.append('bank_account', bankAcct);
    return request('/entities/' + eid + '/bank-transactions/upload', { method: 'POST', body: fd });
  },
  codeBankTransaction: (eid, id, account_code, memo, dims) => request('/entities/' + eid + '/bank-transactions/' + id, { method: 'PUT', body: { account_code, memo, ...(dims || {}) } }),
  splitBankTransaction: (eid, id, splits) => request('/entities/' + eid + '/bank-transactions/' + id + '/splits', { method: 'PUT', body: { splits } }),
  postBankTransactions: (eid, ids) => request('/entities/' + eid + '/bank-transactions/post', { method: 'POST', body: { transaction_ids: ids } }),
  getBankMatchCandidates: (eid, id) => request('/entities/' + eid + '/bank-transactions/' + id + '/match-candidates'),
  matchBankTransaction: (eid, id, je_id) => request('/entities/' + eid + '/bank-transactions/' + id + '/match', { method: 'POST', body: { je_id } }),
  unmatchBankTransaction: (eid, id) => request('/entities/' + eid + '/bank-transactions/' + id + '/unmatch', { method: 'POST' }),
  deleteBankTransaction: (eid, id) => request('/entities/' + eid + '/bank-transactions/' + id, { method: 'DELETE' }),
  deleteBankBatch: (eid, batchId) => request('/entities/' + eid + '/bank-transactions/batch/' + batchId, { method: 'DELETE' }),

  // Bank Rec
  getReconciliations: (eid) => request('/entities/' + eid + '/reconciliations'),
  getCleared: (eid, code) => request('/entities/' + eid + '/cleared/' + code),
  createReconciliation: (eid, data) => request('/entities/' + eid + '/reconciliations', { method: 'POST', body: data }),
  getReconciliationReport: (eid, id) => request('/entities/' + eid + '/reconciliations/' + id + '/report'),
  deleteReconciliation: (eid, id) => request('/entities/' + eid + '/reconciliations/' + id, { method: 'DELETE' }),

  // Entity Workpapers
  getEntityFiles: (eid) => request('/entities/' + eid + '/files'),
  uploadEntityFiles: (eid, files, folderPath) => {
    const fd = new FormData();
    for (const f of files) fd.append('files', f);
    fd.append('folder_path', folderPath || '');
    return request('/entities/' + eid + '/files', { method: 'POST', body: fd });
  },
  downloadEntityFile: (id) => API_BASE + '/entity-files/' + id + '/download?token=' + encodeURIComponent(getToken() || ''),
  deleteEntityFile: (id) => request('/entity-files/' + id, { method: 'DELETE' }),
  replaceEntityFile: (id, file) => {
    const fd = new FormData();
    fd.append('file', file);
    return request('/entity-files/' + id, { method: 'PUT', body: fd });
  },
  createEntityFolder: (eid, folderPath) => request('/entities/' + eid + '/folders', { method: 'POST', body: { folder_path: folderPath } }),
  deleteEntityFolder: (eid, folderPath) => request('/entities/' + eid + '/folders?folder_path=' + encodeURIComponent(folderPath), { method: 'DELETE' }),
  renameEntityFolder: (eid, oldPath, newPath) => request('/entities/' + eid + '/folders/rename', { method: 'PUT', body: { old_path: oldPath, new_path: newPath } }),
  moveEntityFile: (id, folderPath) => request('/entity-files/' + id + '/move', { method: 'PUT', body: { folder_path: folderPath } }),

  // Bill.com integration
  getBillcomConfig: (entityId) => request('/billcom/config/' + entityId),
  saveBillcomConfig: (entityId, body) => request('/billcom/config/' + entityId, { method: 'PUT', body }),
  deleteBillcomConfig: (entityId) => request('/billcom/config/' + entityId, { method: 'DELETE' }),
  testBillcomConnection: (entityId) => request('/billcom/config/' + entityId + '/test', { method: 'POST' }),
  getBillcomAccounts: (entityId) => request('/billcom/accounts/' + entityId),
  getBillcomMappings: (entityId) => request('/billcom/mappings/' + entityId),
  saveBillcomMappings: (entityId, mappings) => request('/billcom/mappings/' + entityId, { method: 'PUT', body: { mappings } }),
  syncBillcom: (entityId) => request('/billcom/sync/' + entityId, { method: 'POST' }),
  pushBillcomCoa: (entityId, body) => request('/billcom/push-coa/' + entityId, { method: 'POST', body }),
  getBillcomSyncLog: (entityId, limit) => request('/billcom/sync-log/' + entityId + (limit ? '?limit=' + limit : '')),
  getApAging: (entityId, asOf) => request('/billcom/ap-aging/' + entityId + (asOf ? '?as_of=' + asOf : '')),

  // Requisition (development-project coding engine)
  seedRequisitionHistory: (eid, body) => request('/requisition/' + eid + '/seed-history', { method: 'POST', body }),
  // Cost-code -> {cost_code_name,...} catalog used to auto-fill the Cost Code Name field.
  getRequisitionCoaMap: (eid) => request('/requisition/' + eid + '/coa-map'),
  predictRequisitionCoding: (eid, lines) => request('/requisition/' + eid + '/predict', { method: 'POST', body: { lines } }),
  // Download a stored invoice's original bytes (PDF/image) by its saved id.
  downloadRequisitionInvoice: (id) => API_BASE + '/requisition/invoice/' + id + '/download?token=' + encodeURIComponent(getToken() || ''),
  // Read one invoice PDF/image with Claude → pre-filled fields + cost-code suggestion.
  readRequisitionInvoice: (eid, file) => {
    const fd = new FormData();
    fd.append('invoice', file);
    return request('/requisition/' + eid + '/read-invoice', { method: 'POST', body: fd });
  },
  // Roll-forward: upload Req#N workbook + new-period invoices, get back the
  // rolled-forward Req#N+1 .xlsx (blob) on success, or a thrown Error carrying
  // the reconciliation detail on a 422 failure. Returns { blob, filename, summary }.
  rollForwardRequisition: async (eid, workbookFile, newCurrent, meta = {}) => {
    const fd = new FormData();
    fd.append('workbook', workbookFile);
    fd.append('newCurrent', JSON.stringify(newCurrent || []));
    if (meta.invoices && meta.invoices.length) fd.append('invoices', JSON.stringify(meta.invoices));
    if (meta.reqNumber != null && meta.reqNumber !== '') fd.append('reqNumber', String(meta.reqNumber));
    if (meta.asOfDate) fd.append('asOfDate', meta.asOfDate);
    // force=true tells the server to download the rolled-forward file even if a
    // required reconciliation check failed (the user opted in after seeing it).
    if (meta.force) fd.append('force', 'true');
    const token = getToken();
    let res;
    try {
      res = await fetch(API_BASE + '/requisition/' + eid + '/rollforward', {
        method: 'POST',
        headers: token ? { Authorization: 'Bearer ' + token } : {},
        body: fd,
      });
    } catch (netErr) {
      // The request never got an HTTP response (browser "Failed to fetch"). There
      // is no server message behind this, so classify it: probe a lightweight
      // endpoint to tell "server restarting/unreachable" apart from "the upload
      // was too big or the roll-forward timed out", and note the upload size.
      let serverUp = false;
      try {
        const h = await fetch(API_BASE + '/turnkey/health', { method: 'GET', cache: 'no-store' });
        serverUp = h.ok;
      } catch (_) { serverUp = false; }
      let mb = 0;
      try {
        const b64 = (meta.invoices || []).reduce((n, i) => n + ((i && i.file_b64 ? i.file_b64.length : 0)), 0);
        mb = Math.round((b64 * 0.75) / 1e5) / 10; // base64 chars -> bytes -> MB (1 dp)
      } catch (_) {}
      const sizeNote = mb >= 1
        ? (' The upload was about ' + mb + ' MB of invoice files, which may have exceeded the size/time limit — try rolling forward with fewer or smaller invoice PDFs at once.')
        : '';
      const cause = serverUp
        ? ('The server is reachable, so the roll-forward request itself failed to complete — usually the upload was too large or it took too long to return.' + sizeNote)
        : 'Could not reach the server. It may be restarting after a recent update, or briefly offline. Wait about 30 seconds and try again.';
      const raw = (netErr && netErr.message) ? netErr.message : 'network error';
      const err = new Error('Roll-forward could not be sent (' + raw + '). ' + cause);
      err.network = true;
      err.detail = { networkError: raw, serverReachable: serverUp, approxUploadMB: mb };
      throw err;
    }
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const ctype = res.headers.get('content-type') || '';
    if (!res.ok || ctype.includes('application/json')) {
      let data = {}; try { data = await res.json(); } catch {}
      const err = new Error(data.error || 'Roll-forward failed');
      err.detail = data;
      throw err;
    }
    let summary = {}; try { summary = JSON.parse(res.headers.get('x-reconcile-summary') || '{}'); } catch {}
    let failedChecks = []; try { failedChecks = JSON.parse(res.headers.get('x-reconcile-failed') || '[]'); } catch {}
    const workpaperFolder = res.headers.get('x-workpaper-folder') || '';
    let workpaperSaved = {}; try { workpaperSaved = JSON.parse(res.headers.get('x-workpaper-saved') || '{}'); } catch {}
    // Invoice-packet PDF saved to Workpapers; its entity-file id lets the client
    // also download the packet into the user's Downloads folder.
    const packetFileId = res.headers.get('x-packet-file-id') || '';
    const packetFileName = res.headers.get('x-packet-file-name') || '';
    // '1' when the server downloaded despite a failed required check (force).
    const forced = res.headers.get('x-reconcile-forced') === '1';
    // How the development fee was determined this period (or that it needs manual
    // entry): { amount, base, rate_text, source, needs_review, note, prior, validated }.
    let devFee = null; try { devFee = JSON.parse(res.headers.get('x-dev-fee') || 'null'); } catch {}
    const cd = res.headers.get('content-disposition') || '';
    const m = cd.match(/filename="?([^"]+)"?/);
    const filename = m ? m[1] : 'Requisition_Report.xlsx';
    const blob = await res.blob();
    return { blob, filename, summary, failedChecks, workpaperFolder, workpaperSaved, packetFileId, packetFileName, forced, devFee };
  },

  // Workpapers › Management Fee: analyze a prior-quarter workbook, then generate
  // the next quarter as a downloadable .xlsx.
  mgmtFeeAnalyze: async (eid, file) => {
    const fd = new FormData();
    fd.append('workbook', file);
    const token = getToken();
    const res = await fetch(API_BASE + '/workpapers/mgmt-fee/' + eid + '/analyze', {
      method: 'POST', headers: token ? { Authorization: 'Bearer ' + token } : {}, body: fd,
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Analyze failed');
    return data;
  },
  mgmtFeeGenerate: async (eid, file, changes, quarterStart) => {
    const fd = new FormData();
    fd.append('workbook', file);
    fd.append('changes', JSON.stringify(changes || []));
    if (quarterStart) fd.append('quarter_start', quarterStart);
    const token = getToken();
    const res = await fetch(API_BASE + '/workpapers/mgmt-fee/' + eid + '/generate', {
      method: 'POST', headers: token ? { Authorization: 'Bearer ' + token } : {}, body: fd,
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const ctype = res.headers.get('content-type') || '';
    if (!res.ok || ctype.includes('application/json')) {
      let data = {}; try { data = await res.json(); } catch {}
      throw new Error(data.error || 'Generate failed');
    }
    let summary = {}; try { summary = JSON.parse(res.headers.get('x-mgmt-fee-summary') || '{}'); } catch {}
    const cd = res.headers.get('content-disposition') || '';
    const m = cd.match(/filename="?([^"]+)"?/);
    const filename = m ? m[1] : 'Mgmt_Fee_Calc.xlsx';
    const blob = await res.blob();
    return { blob, filename, summary };
  },

  // Workpapers › Financial Statements: preview tie-outs, then generate the
  // merged PDF package (cover + exec summary + GL statements + requisition).
  financialStatementsPreview: async (eid, asOf, period) => {
    const token = getToken();
    const res = await fetch(API_BASE + '/workpapers/financial-statements/' + eid + '/preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...(token ? { Authorization: 'Bearer ' + token } : {}) },
      body: JSON.stringify({ as_of: asOf, period: period || 'monthly' }),
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Preview failed');
    return data;
  },
  financialStatementsGenerate: async (eid, asOf, period, execSummaryFile, reqReportFile) => {
    const fd = new FormData();
    fd.append('as_of', asOf);
    fd.append('period', period || 'monthly');
    if (execSummaryFile) fd.append('execSummary', execSummaryFile);
    if (reqReportFile) fd.append('reqReport', reqReportFile);
    const token = getToken();
    const res = await fetch(API_BASE + '/workpapers/financial-statements/' + eid + '/generate', {
      method: 'POST', headers: token ? { Authorization: 'Bearer ' + token } : {}, body: fd,
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const ctype = res.headers.get('content-type') || '';
    if (!res.ok || ctype.includes('application/json')) {
      let data = {}; try { data = await res.json(); } catch {}
      throw new Error(data.error || 'Generate failed');
    }
    let summary = {}; try { summary = JSON.parse(res.headers.get('x-financials-summary') || '{}'); } catch {}
    const cd = res.headers.get('content-disposition') || '';
    const m = cd.match(/filename="?([^"]+)"?/);
    const filename = m ? m[1] : 'Financial_Statements.pdf';
    const blob = await res.blob();
    return { blob, filename, summary };
  },

  setToken, getToken, clearToken,
};
