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
  // Parse JSON only when the response actually is JSON. A non-JSON body (an HTML
  // error page, an SPA shell served by mistake, a proxy 502) would otherwise make
  // res.json() throw "Unexpected token '<'" and hide the real status. Surface a
  // readable error with the HTTP status instead.
  const ctype = res.headers.get('content-type') || '';
  if (!ctype.includes('application/json')) {
    const text = await res.text().catch(() => '');
    if (res.ok) { const err = new Error('Unexpected non-JSON response from server (HTTP ' + res.status + ')'); err.detail = { status: res.status, body: text.slice(0, 500) }; throw err; }
    const err = new Error('Server error (HTTP ' + res.status + ')' + (res.status === 502 || res.status === 504 ? ' — the request may have timed out. For large scanned PDFs, try again.' : ''));
    err.detail = { status: res.status, body: text.slice(0, 500) }; throw err;
  }
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
  // Per-user UI preferences (sidebar category item order, etc.)
  getMyPrefs: () => request('/me/prefs'),
  saveMyPrefs: (patch) => request('/me/prefs', { method: 'PUT', body: patch }),
  forgotPassword: (email) => request('/auth/forgot-password', { method: 'POST', body: { email } }),
  resetPassword: (token, new_password) => request('/auth/reset-password', { method: 'POST', body: { token, new_password } }),
  adminResetPassword: (uid, pw) => request('/auth/admin-reset-password', { method: 'POST', body: { user_id: uid, new_password: pw } }),

  // Users
  getUsers: () => request('/users'),
  deleteUser: (id) => request('/users/' + id, { method: 'DELETE' }),
  updateUser: (id, data) => request('/users/' + id, { method: 'PUT', body: data }),
  getUserEntityAccess: (id) => request('/users/' + id + '/entity-access'),
  setUserEntityAccess: (id, entity_ids, exclusions, levels) => request('/users/' + id + '/entity-access', { method: 'PUT', body: { entity_ids, exclusions, levels } }),

  // User groups (bundle users + grant entity access to all at once, e.g. CLA)
  getGroups: () => request('/groups'),
  getGroup: (id) => request('/groups/' + id),
  setGroupMembers: (id, user_ids) => request('/groups/' + id + '/members', { method: 'PUT', body: { user_ids } }),
  setGroupEntities: (id, entity_ids, levels) => request('/groups/' + id + '/entities', { method: 'PUT', body: { entity_ids, levels } }),

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

  // Period locking
  getPeriods: (eid) => request('/entities/' + eid + '/periods'),
  softClosePeriod: (eid, month, reason) => request('/entities/' + eid + '/periods/soft-close', { method: 'POST', body: { month, reason } }),
  reopenSoftPeriod: (eid, month) => request('/entities/' + eid + '/periods/reopen-soft', { method: 'POST', body: { month } }),
  hardCloseYear: (eid, year, reason) => request('/entities/' + eid + '/periods/hard-close-year', { method: 'POST', body: { year, reason } }),
  reopenYear: (eid, year) => request('/entities/' + eid + '/periods/reopen-year', { method: 'POST', body: { year } }),

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
    if (opts.project_id) p.push('project_id=' + opts.project_id);
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
  // The server generates the "Items Needing Attention" analysis itself and
  // appends it to the workbook — no client-side analysis payload is sent.
  getTtmPLXlsx: async (eid, asOf) => {
    const token = getToken();
    const res = await fetch(API_BASE + '/entities/' + eid + '/ttm-pl.xlsx', {
      method: 'POST',
      headers: { 'content-type': 'application/json', ...(token ? { Authorization: 'Bearer ' + token } : {}) },
      body: JSON.stringify({ as_of: asOf }),
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
  // Styled workbook build (ExcelJS, server-side). The client still assembles the
  // rows and live formulas; `style` names the header / subtotal / grand-total rows
  // so the server can draw the underlines SheetJS cannot.
  postStyledXlsx: async (spec) => {
    const token = getToken();
    const res = await fetch(API_BASE + '/xlsx-styled', {
      method: 'POST',
      headers: { 'content-type': 'application/json', ...(token ? { Authorization: 'Bearer ' + token } : {}) },
      body: JSON.stringify(spec),
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const ctype = res.headers.get('content-type') || '';
    if (!res.ok || ctype.includes('application/json')) { let d = {}; try { d = await res.json(); } catch {} throw new Error(d.error || 'Export failed'); }
    const cd = res.headers.get('content-disposition') || '';
    const m = cd.match(/filename="?([^"]+)"?/);
    return { blob: await res.blob(), filename: m ? m[1] : (spec.filename || 'Report.xlsx') };
  },
  backfillDocNumbers: (eid, dryRun) => request('/entities/' + eid + '/backfill-doc-numbers', { method: 'POST', body: { dry_run: !!dryRun } }),
  getBillcomDimensionMaps: (eid) => request('/billcom/dimension-maps/' + eid),
  autoBillcomDimensionMaps: (eid) => request('/billcom/dimension-maps/' + eid + '/auto', { method: 'POST', body: {} }),
  saveBillcomDimensionMaps: (eid, body) => request('/billcom/dimension-maps/' + eid, { method: 'PUT', body }),
  retagBillcomProjects: (eid, body) => request('/billcom/retag-projects/' + eid, { method: 'POST', body: body || {} }),
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
  // ── Accounts Receivable: settings, recurring templates, invoices, receipts, aging ──
  getArSettings: (eid) => request('/entities/' + eid + '/ar/settings'),
  saveArSettings: (eid, data) => request('/entities/' + eid + '/ar/settings', { method: 'PUT', body: data }),
  getArTemplates: (eid) => request('/entities/' + eid + '/ar/templates'),
  createArTemplate: (eid, data) => request('/entities/' + eid + '/ar/templates', { method: 'POST', body: data }),
  updateArTemplate: (eid, id, data) => request('/entities/' + eid + '/ar/templates/' + id, { method: 'PATCH', body: data }),
  deleteArTemplate: (eid, id) => request('/entities/' + eid + '/ar/templates/' + id, { method: 'DELETE' }),
  generateArInvoice: (eid, id, invoice_date) => request('/entities/' + eid + '/ar/templates/' + id + '/generate', { method: 'POST', body: { invoice_date } }),
  getArInvoices: (eid, opts = {}) => {
    const p = [];
    if (opts.status) p.push('status=' + encodeURIComponent(opts.status));
    if (opts.from) p.push('from=' + opts.from);
    if (opts.to) p.push('to=' + opts.to);
    return request('/entities/' + eid + '/ar/invoices' + (p.length ? '?' + p.join('&') : ''));
  },
  getArInvoice: (eid, id) => request('/entities/' + eid + '/ar/invoices/' + id),
  createArInvoice: (eid, data) => request('/entities/' + eid + '/ar/invoices', { method: 'POST', body: data }),
  updateArInvoice: (eid, id, data) => request('/entities/' + eid + '/ar/invoices/' + id, { method: 'PATCH', body: data }),
  deleteArInvoice: (eid, id) => request('/entities/' + eid + '/ar/invoices/' + id, { method: 'DELETE' }),
  voidArInvoice: (eid, id, date) => request('/entities/' + eid + '/ar/invoices/' + id + '/void', { method: 'POST', body: { date } }),
  sendArInvoice: (eid, id, body = {}) => request('/entities/' + eid + '/ar/invoices/' + id + '/send', { method: 'POST', body }),
  markArInvoiceSent: (eid, id) => request('/entities/' + eid + '/ar/invoices/' + id + '/mark-sent', { method: 'POST', body: {} }),
  saveArInvoicePdf: (eid, id) => request('/entities/' + eid + '/ar/invoices/' + id + '/save-pdf', { method: 'POST', body: {} }),
  arInvoicePdfUrl: (eid, id) => API_BASE + '/entities/' + eid + '/ar/invoices/' + id + '/pdf?token=' + encodeURIComponent(getToken() || ''),
  addArReceipt: (eid, id, data) => request('/entities/' + eid + '/ar/invoices/' + id + '/receipts', { method: 'POST', body: data }),
  deleteArReceipt: (eid, id, rid) => request('/entities/' + eid + '/ar/invoices/' + id + '/receipts/' + rid, { method: 'DELETE' }),
  getArAging: (eid, asOf) => request('/entities/' + eid + '/ar/aging' + (asOf ? '?as_of=' + asOf : '')),
  getArOpenInvoices: (eid, amount, excludeTxn) => { const q = []; if (amount != null) q.push('amount=' + amount); if (excludeTxn != null) q.push('exclude_txn=' + excludeTxn); return request('/entities/' + eid + '/ar/open-invoices' + (q.length ? '?' + q.join('&') : '')); },
  createCreditMemo: (eid, data) => request('/entities/' + eid + '/ar/credit-memos', { method: 'POST', body: data }),
  getCreditMemos: (eid, openOnly) => request('/entities/' + eid + '/ar/credit-memos' + (openOnly ? '?open=1' : '')),
  applyCreditMemo: (eid, id, data) => request('/entities/' + eid + '/ar/credit-memos/' + id + '/apply', { method: 'POST', body: data }),
  deleteCreditApplication: (eid, rid) => request('/entities/' + eid + '/ar/credit-applications/' + rid, { method: 'DELETE' }),
  arOpeningImport: (eid, items, force, opts = {}) => request('/entities/' + eid + '/ar/opening-import', { method: 'POST', body: { items, force: !!force, as_of: opts.as_of || null, allow_over_gl: !!opts.allow_over_gl } }),
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

  // Wire coding notes (auto-populate coding on statement upload)
  getBankCodingNotes: (eid) => request('/entities/' + eid + '/bank-coding-notes'),
  createBankCodingNote: (eid, note) => request('/entities/' + eid + '/bank-coding-notes', { method: 'POST', body: note }),
  updateBankCodingNote: (eid, id, note) => request('/entities/' + eid + '/bank-coding-notes/' + id, { method: 'PUT', body: note }),
  deleteBankCodingNote: (eid, id) => request('/entities/' + eid + '/bank-coding-notes/' + id, { method: 'DELETE' }),
  uploadBankCodingNoteFiles: (eid, id, files) => {
    const fd = new FormData();
    for (const f of files) fd.append('files', f);
    return request('/entities/' + eid + '/bank-coding-notes/' + id + '/attachments', { method: 'POST', body: fd });
  },
  bankCodingNoteFileUrl: (id) => API_BASE + '/bank-coding-note-attachments/' + id + '/download?token=' + encodeURIComponent(getToken() || ''),
  deleteBankCodingNoteFile: (id) => request('/bank-coding-note-attachments/' + id, { method: 'DELETE' }),

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
  downloadEntityFolder: (eid, folderPath) => API_BASE + '/entities/' + eid + '/folders/download?folder_path=' + encodeURIComponent(folderPath || '') + '&token=' + encodeURIComponent(getToken() || ''),
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
  getBillcomEntities: () => request('/billcom/entities'),
  getBillcomConfig: (entityId) => request('/billcom/config/' + entityId),
  saveBillcomConfig: (entityId, body) => request('/billcom/config/' + entityId, { method: 'PUT', body }),
  setBillcomCutoff: (entityId, syncCutoffDate, lines, asOf) => request('/billcom/config/' + entityId + '/cutoff', { method: 'PUT', body: { sync_cutoff_date: syncCutoffDate || null, ...(lines ? { lines, as_of: asOf || null } : {}) } }),
  checkApAgingOverlap: (entityId) => request('/billcom/ap-aging-check/' + entityId, { method: 'POST', body: {} }),
  deleteBillcomConfig: (entityId) => request('/billcom/config/' + entityId, { method: 'DELETE' }),
  testBillcomConnection: (entityId) => request('/billcom/config/' + entityId + '/test', { method: 'POST' }),
  getBillcomAccounts: (entityId) => request('/billcom/accounts/' + entityId),
  getBillcomMappings: (entityId) => request('/billcom/mappings/' + entityId),
  saveBillcomMappings: (entityId, mappings) => request('/billcom/mappings/' + entityId, { method: 'PUT', body: { mappings } }),
  syncBillcom: (entityId) => request('/billcom/sync/' + entityId, { method: 'POST' }),
  unsyncBillcom: (entityId, dryRun) => request('/billcom/unsync/' + entityId, { method: 'POST', body: { dry_run: !!dryRun } }),
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

  // ── Editable / persistent Requisition draft ──────────────────────────────
  // A `phase` (e.g. '2','2a') selects a rail-assets stream; omit/'' = default.
  // Without a phase, getRequisitionDraft returns { draft, drafts:[...], is_rail }
  // so the UI can list Phase 2 / 2a; with one it returns that stream's draft.
  getRequisitionDraft: (eid, phase) => request('/requisition/' + eid + '/draft' + (phase ? ('?phase=' + encodeURIComponent(phase)) : '')),
  getRequisitionSeedSource: (eid, phase) => request('/requisition/' + eid + '/draft/seed-source' + (phase ? ('?phase=' + encodeURIComponent(phase)) : '')),
  // Create the open draft. Auto-seeds when a prior finalized Req exists; pass a
  // workbook File for the first-time case. baseChoice ('uploaded'|'filed')
  // answers the Option-B upload-guard prompt. `phase` is the rail-assets stream.
  // On a guard conflict the thrown Error carries err.detail.conflict + message.
  createRequisitionDraft: (eid, { workbookFile, reqNumber, asOfDate, baseChoice, phase } = {}) => {
    const fd = new FormData();
    if (workbookFile) fd.append('workbook', workbookFile);
    if (reqNumber != null && reqNumber !== '') fd.append('reqNumber', String(reqNumber));
    if (asOfDate) fd.append('asOfDate', asOfDate);
    if (baseChoice) fd.append('baseChoice', baseChoice);
    if (phase) fd.append('phase', phase);
    return request('/requisition/' + eid + '/draft', { method: 'POST', body: fd });
  },
  addRequisitionDraftInvoice: (eid, inv, phase) => request('/requisition/' + eid + '/draft/invoice', { method: 'POST', body: { ...inv, phase } }),
  updateRequisitionDraftInvoice: (eid, id, inv, phase) => request('/requisition/' + eid + '/draft/invoice/' + id, { method: 'PUT', body: { ...inv, phase } }),
  deleteRequisitionDraftInvoice: (eid, id, phase) => request('/requisition/' + eid + '/draft/invoice/' + id + (phase ? ('?phase=' + encodeURIComponent(phase)) : ''), { method: 'DELETE' }),
  rollRequisitionDraft: (eid, meta = {}) => request('/requisition/' + eid + '/draft/roll', { method: 'POST', body: { reqNumber: meta.reqNumber, asOfDate: meta.asOfDate, phase: meta.phase, force: !!meta.force } }),
  downloadRequisitionDraftUrl: (eid, phase) => API_BASE + '/requisition/' + eid + '/draft/download?token=' + encodeURIComponent(getToken() || '') + (phase ? ('&phase=' + encodeURIComponent(phase)) : ''),
  finalizeRequisitionDraft: (eid, force = false, phase) => request('/requisition/' + eid + '/draft/finalize', { method: 'POST', body: { force, phase } }),
  // Reopen the latest finalized requisition of a stream back into an editable draft.
  reopenRequisitionDraft: (eid, phase) => request('/requisition/' + eid + '/draft/reopen', { method: 'POST', body: { phase } }),
  discardRequisitionDraft: (eid, phase) => request('/requisition/' + eid + '/draft' + (phase ? ('?phase=' + encodeURIComponent(phase)) : ''), { method: 'DELETE' }),

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

  // Workpapers › GP Fees & Expenses (CLRF): build the quarterly schedule from the
  // portfolio-company ledgers. The server also files a copy under the entity's
  // workpaper folder, one per period.
  gpFeesGenerate: async (eid, quarterEnd) => {
    const token = getToken();
    const res = await fetch(API_BASE + '/workpapers/gp-fees/' + eid + '/generate', {
      method: 'POST',
      headers: Object.assign({ 'Content-Type': 'application/json' }, token ? { Authorization: 'Bearer ' + token } : {}),
      body: JSON.stringify({ quarter_end: quarterEnd }),
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const ctype = res.headers.get('content-type') || '';
    if (!res.ok || ctype.includes('application/json')) {
      let data = {}; try { data = await res.json(); } catch {}
      throw new Error(data.error || 'Report failed');
    }
    let summary = {}; try { summary = JSON.parse(res.headers.get('x-gp-fees-summary') || '{}'); } catch {}
    const cd = res.headers.get('content-disposition') || '';
    const m = cd.match(/filename="?([^"]+)"?/);
    return { blob: await res.blob(), filename: m ? m[1] : 'CLRF_GP_Fees.xlsx', summary };
  },

  // Workpapers › Valuation Summary (CLRF): take the prior quarter's valuation
  // workbook and inject the GL-derived figures for the target quarter, saving a
  // copy under Workpapers › Valuation › <quarter>.
  investmentValuationGenerate: async (eid, quarterEnd) => {
    const token = getToken();
    const res = await fetch(API_BASE + '/workpapers/investment-valuation/' + eid + '/generate', {
      method: 'POST',
      headers: Object.assign({ 'Content-Type': 'application/json' }, token ? { Authorization: 'Bearer ' + token } : {}),
      body: JSON.stringify({ quarter_end: quarterEnd }),
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    let data = {}; try { data = await res.json(); } catch {}
    if (!res.ok) throw new Error(data.error || 'Report failed');
    return data;
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
  financialStatementsGenerate: async (eid, asOf, period, execSummaryFile, reqReportFile, wipFile) => {
    const fd = new FormData();
    fd.append('as_of', asOf);
    fd.append('period', period || 'monthly');
    if (execSummaryFile) fd.append('execSummary', execSummaryFile);
    // WIP schedule -> merged as the package's 'Schedule of Contracts' section.
    if (wipFile) fd.append('wip', wipFile);
    // reqReportFile may be a single File or an array of up to two Files.
    const reqFiles = Array.isArray(reqReportFile) ? reqReportFile.filter(Boolean) : (reqReportFile ? [reqReportFile] : []);
    for (const f of reqFiles) fd.append('reqReport', f);
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
  // WIP schedule (Schedule of Contracts). Stored per period, so as_of decides
  // which month's schedule is replaced/read.
  financialStatementsWipUpload: async (eid, asOf, file) => {
    const fd = new FormData();
    fd.append('as_of', asOf);
    fd.append('wip', file);
    const token = getToken();
    const res = await fetch(API_BASE + '/workpapers/financial-statements/' + eid + '/wip', {
      method: 'POST', headers: token ? { Authorization: 'Bearer ' + token } : {}, body: fd,
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'WIP upload failed');
    return data;
  },
  // Operating budget (annual). The workbook is parsed server-side into budget
  // lines on upload; the monthly Budget-to-Actual schedule reads those, not the
  // file. Re-uploading a year creates a new version.
  financialStatementsBudgetUpload: async (eid, file, note) => {
    const fd = new FormData();
    fd.append('budget', file);
    if (note) fd.append('note', note);
    const token = getToken();
    const res = await fetch(API_BASE + '/workpapers/financial-statements/' + eid + '/budget', {
      method: 'POST', headers: token ? { Authorization: 'Bearer ' + token } : {}, body: fd,
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Budget upload failed');
    return data;
  },
  financialStatementsBudgetStatus: async (eid, asOf) => {
    const token = getToken();
    const qs = 'as_of=' + encodeURIComponent(asOf);
    const res = await fetch(API_BASE + '/workpapers/financial-statements/' + eid + '/budget?' + qs, {
      method: 'GET', headers: token ? { Authorization: 'Bearer ' + token } : {},
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'Budget status failed');
    return data;
  },
  financialStatementsWipStatus: async (eid, asOf) => {
    const token = getToken();
    const qs = 'as_of=' + encodeURIComponent(asOf);
    const res = await fetch(API_BASE + '/workpapers/financial-statements/' + eid + '/wip?' + qs, {
      method: 'GET', headers: token ? { Authorization: 'Bearer ' + token } : {},
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || 'WIP status failed');
    return data;
  },
  // Excel version of the statement package — mirrors the PDF formatting, one
  // worksheet per statement. GET (no uploads); returns a downloadable .xlsx blob.
  financialStatementsExcel: async (eid, asOf, period) => {
    const token = getToken();
    const qs = 'as_of=' + encodeURIComponent(asOf) + '&period=' + encodeURIComponent(period || 'monthly');
    const res = await fetch(API_BASE + '/workpapers/financial-statements/' + eid + '/excel?' + qs, {
      method: 'GET', headers: token ? { Authorization: 'Bearer ' + token } : {},
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const ctype = res.headers.get('content-type') || '';
    if (!res.ok || ctype.includes('application/json')) {
      let data = {}; try { data = await res.json(); } catch {}
      throw new Error(data.error || 'Excel export failed');
    }
    const cd = res.headers.get('content-disposition') || '';
    const m = cd.match(/filename=\"?([^\"]+)\"?/);
    const filename = m ? m[1] : 'Financial_Statements.xlsx';
    const blob = await res.blob();
    return { blob, filename };
  },

  // Workpapers › Insurance Allocation: upload the carrier billing invoice + the
  // consolidated billing report; the server computes the entity allocation,
  // files the workpaper under Workpapers, and returns the .xlsx + a summary.
  insuranceAllocationGenerate: async (eid, invoiceFile, consolidatedFile) => {
    const fd = new FormData();
    fd.append('invoice', invoiceFile);
    fd.append('consolidated', consolidatedFile);
    const token = getToken();
    const res = await fetch(API_BASE + '/entities/' + eid + '/insurance-allocation', {
      method: 'POST', headers: token ? { Authorization: 'Bearer ' + token } : {}, body: fd,
    });
    if (res.status === 401) { clearToken(); window.location.reload(); return null; }
    const ctype = res.headers.get('content-type') || '';
    if (!res.ok || ctype.includes('application/json')) {
      let data = {}; try { data = await res.json(); } catch {}
      throw new Error(data.error || 'Allocation failed');
    }
    let summary = {}; try { summary = JSON.parse(decodeURIComponent(res.headers.get('x-alloc-summary') || '')); } catch {}
    const cd = res.headers.get('content-disposition') || '';
    const m = cd.match(/filename=\"?([^\"]+)\"?/);
    const filename = m ? m[1] : 'Insurance Allocation.xlsx';
    const blob = await res.blob();
    return { blob, filename, summary };
  },

  // ── Intercompany ──
  // A mapping says which GL account faces which entity. That is the whole
  // setup: reconciliation runs for one entity and follows the mappings.
  getIcMappings: (q = {}) => {
    const p = new URLSearchParams();
    if (q.entity_id) p.set('entity_id', q.entity_id);
    const s = p.toString();
    return request('/intercompany/mappings' + (s ? '?' + s : ''));
  },
  // Returns { suggestions, hidden_individuals, hidden_examples }. Outside
  // individuals holding capital are dropped by default — they are never
  // intercompany counterparties. Pass includeIndividuals to see them anyway.
  suggestIcMappings: (entity_id, as_of, includeIndividuals) =>
    request('/intercompany/mappings/suggest?entity_id=' + entity_id
      + (as_of ? '&as_of=' + as_of : '')
      + (includeIndividuals ? '&include_individuals=1' : '')),
  // Every account on the entity that carries no mapping yet, INCLUDING the ones
  // whose name the parser cannot read as intercompany. That is the difference
  // from suggestIcMappings, which can only propose what it recognises.
  // The finished pairs: both sides' GL accounts and balances on one row.
  getIcMappedAccounts: (entity_id, as_of) =>
    request('/intercompany/accounts/mapped?entity_id=' + entity_id + (as_of ? '&as_of=' + as_of : '')),
  getIcUnmappedAccounts: (entity_id, as_of) =>
    request('/intercompany/accounts/unmapped?entity_id=' + entity_id + (as_of ? '&as_of=' + as_of : '')),
  // The counterparty's accounts, ranked as answers to one of ours. Names are not
  // trustworthy here (a "Due from" account can carry a credit), so the ranking
  // is: does their name resolve back to us, then how closely the balances offset.
  getIcCounterpartyAccounts: (q) =>
    request('/intercompany/counterparty-accounts?entity_id=' + q.entity_id
      + '&counterparty_entity_id=' + q.counterparty_entity_id
      + '&account_code=' + encodeURIComponent(q.account_code)
      + (q.as_of ? '&as_of=' + q.as_of : '')),
  // Accepts one mapping object or an array of them (the "accept suggestions" path).
  createIcMapping: (b) => request('/intercompany/mappings', { method: 'POST', body: b }),
  // Manual counterparty value for an external mapping (their books are outside
  // CloudLedger). Keyed by as-of date; null balance clears it.
  setIcManualBalance: (mapping_id, as_of, balance) =>
    request('/intercompany/manual-balance', { method: 'PUT', body: { mapping_id, as_of, balance } }),
  // External entity trial balances: uploaded monthly for counterparties with
  // no ledger in CloudLedger, then read by mapping and reconciliation.
  getExternalTbs: () => request('/intercompany/external-tbs'),
  getExternalTbLines: (node_id, as_of) =>
    request('/intercompany/external-tb-lines?node_id=' + node_id + '&as_of=' + as_of),
  deleteExternalTb: (node_id, as_of) =>
    request('/intercompany/external-tbs?node_id=' + node_id + '&as_of=' + as_of, { method: 'DELETE' }),
  uploadExternalTb: (node_id, as_of, file) => {
    const fd = new FormData();
    fd.append('file', file); fd.append('node_id', node_id); fd.append('as_of', as_of);
    return request('/intercompany/external-tbs', { method: 'POST', body: fd });
  },
  updateIcMapping: (id, b) => request('/intercompany/mappings/' + id, { method: 'PUT', body: b }),
  deleteIcMapping: (id) => request('/intercompany/mappings/' + id, { method: 'DELETE' }),
  // Reconciliation is run for ONE entity, so every row can carry the account
  // code and name on both sides.
  reconcileIcEntity: (entity_id, as_of) =>
    request('/intercompany/reconcile/entity?entity_id=' + entity_id + (as_of ? '&as_of=' + as_of : '')),

  // People. Marking a name as an individual stops it being proposed as a
  // counterparty anywhere — the automatic name-shape test only covers capital
  // accounts, so a person on a due-from / due-to account is marked by hand.
  getIcPeople: () => request('/intercompany/people'),
  markIcPerson: (b) => request('/intercompany/people', { method: 'POST', body: b }),
  unmarkIcPerson: (id) => request('/intercompany/people/' + id, { method: 'DELETE' }),

  // Registered companies: counterparties that appear on the org charts but have
  // no ledger here (QOZBs, joint ventures, sponsor holdcos). Stored as org
  // nodes, so one registered here can later be placed in an ownership tree.
  getIcCompanies: () => request('/intercompany/companies'),
  createIcCompany: (b) => request('/intercompany/companies', { method: 'POST', body: b }),

  // -- Org structure (ownership tree) --
  // Nodes are either a CloudLedger entity or a ledger-less shell; edges carry
  // the ownership percent. The investment tie-out is scoped by root node, not
  // by IC group, because it walks the chain rather than a flat member list.
  getOrgStructure: () => request('/org-structure'),
  getOrgTree: (root_node_id) => request('/org-structure/tree?root_node_id=' + root_node_id),
  createOrgNode: (b) => request('/org-structure/nodes', { method: 'POST', body: b }),
  saveOrgNode: (id, b) => request('/org-structure/nodes/' + id, { method: 'PUT', body: b }),
  deleteOrgNode: (id) => request('/org-structure/nodes/' + id, { method: 'DELETE' }),
  createOrgEdge: (b) => request('/org-structure/edges', { method: 'POST', body: b }),
  saveOrgEdge: (id, b) => request('/org-structure/edges/' + id, { method: 'PUT', body: b }),
  deleteOrgEdge: (id) => request('/org-structure/edges/' + id, { method: 'DELETE' }),
  seedOrgClrf: () => request('/org-structure/seed/clrf', { method: 'POST', body: {} }),
  reconcileOrgInvestments: (root_node_id, as_of) =>
    request('/org-structure/reconcile/investments?root_node_id=' + root_node_id + (as_of ? '&as_of=' + as_of : '')),

  // ── Consolidation (Braker / HP) ──
  // The operating column is an uploaded trial balance; every route is scoped to
  // the parent entity and refuses anything outside Braker and HP server-side.
  getConsolGroups: () => request('/consolidation/groups'),
  getConsol: (parentEid) => request('/consolidation/' + parentEid),
  getConsolTb: (parentEid, entityId, asOf) =>
    request('/consolidation/' + parentEid + '/operating-tb'
      + (entityId ? '?entity_id=' + entityId : '') + (asOf ? (entityId ? '&' : '?') + 'as_of=' + asOf : '')),
  uploadConsolTb: (parentEid, entityId, asOf, file) => {
    const fd = new FormData();
    fd.append('file', file); fd.append('entity_id', entityId); fd.append('as_of', asOf);
    return request('/consolidation/' + parentEid + '/operating-tb', { method: 'POST', body: fd });
  },
  deleteConsolTb: (parentEid, entityId, asOf) =>
    request('/consolidation/' + parentEid + '/operating-tb?entity_id=' + entityId + '&as_of=' + asOf, { method: 'DELETE' }),
  getConsolMap: (parentEid, entityId) =>
    request('/consolidation/' + parentEid + '/map' + (entityId ? '?entity_id=' + entityId : '')),
  saveConsolMap: (parentEid, entityId, rows) =>
    request('/consolidation/' + parentEid + '/map', { method: 'PUT', body: { entity_id: entityId, rows } }),
  getConsolFundingAccounts: (parentEid) => request('/consolidation/' + parentEid + '/funding-accounts'),
  saveConsolFundingAccounts: (parentEid, rows) =>
    request('/consolidation/' + parentEid + '/funding-accounts', { method: 'PUT', body: { rows } }),
  getConsolFullEliminations: (parentEid) => request('/consolidation/' + parentEid + '/full-eliminations'),
  saveConsolFullEliminations: (parentEid, rows) =>
    request('/consolidation/' + parentEid + '/full-eliminations', { method: 'PUT', body: { rows } }),
  getConsolEliminations: (parentEid, asOf) =>
    request('/consolidation/' + parentEid + '/eliminations?as_of=' + asOf),
  getConsolSchedules: (parentEid, asOf) =>
    request('/consolidation/' + parentEid + '/schedules?as_of=' + asOf),

  setToken, getToken, clearToken,
};
