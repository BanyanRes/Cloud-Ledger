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
  if (!res.ok) throw new Error(data.error || 'Request failed');
  return data;
}

export const api = {
  // Auth
  login: (email, pw) => request('/auth/login', { method: 'POST', body: { email, password: pw } }),
  signup: (name, email, pw, role) => request('/auth/signup', { method: 'POST', body: { name, email, password: pw, role } }),
  me: () => request('/auth/me'),
  changePassword: (cur, nw) => request('/auth/change-password', { method: 'POST', body: { current_password: cur, new_password: nw } }),
  forgotPassword: (email) => request('/auth/forgot-password', { method: 'POST', body: { email } }),
  adminResetPassword: (uid, pw) => request('/auth/admin-reset-password', { method: 'POST', body: { user_id: uid, new_password: pw } }),

  // Users
  getUsers: () => request('/users'),
  deleteUser: (id) => request('/users/' + id, { method: 'DELETE' }),
  updateUser: (id, data) => request('/users/' + id, { method: 'PUT', body: data }),

  // Entities
  getEntities: () => request('/entities'),
  createEntity: (code, name) => request('/entities', { method: 'POST', body: { code, name } }),
  bulkCreateEntities: (ents) => request('/entities/bulk', { method: 'POST', body: { entities: ents } }),
  deleteEntity: (id) => request('/entities/' + id, { method: 'DELETE' }),

  // Accounts
  getAccounts: (eid) => request('/entities/' + eid + '/accounts'),
  createAccount: (eid, data) => request('/entities/' + eid + '/accounts', { method: 'POST', body: data }),
  deleteAccount: (eid, code) => request('/entities/' + eid + '/accounts/' + code, { method: 'DELETE' }),

  // Journal Entries
  getEntries: (eid, from, to) => {
    let q = '/entities/' + eid + '/entries'; const p = [];
    if (from) p.push('from=' + from); if (to) p.push('to=' + to);
    return request(q + (p.length ? '?' + p.join('&') : ''));
  },
  createEntry: (eid, data) => request('/entities/' + eid + '/entries', { method: 'POST', body: data }),
  updateEntry: (eid, id, data) => request('/entities/' + eid + '/entries/' + id, { method: 'PUT', body: data }),
  deleteEntry: (eid, id) => request('/entities/' + eid + '/entries/' + id, { method: 'DELETE' }),

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
    return request('/entities/' + eid + '/balances' + (p.length ? '?' + p.join('&') : ''));
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
  codeBankTransaction: (eid, id, account_code, memo) => request('/entities/' + eid + '/bank-transactions/' + id, { method: 'PUT', body: { account_code, memo } }),
  postBankTransactions: (eid, ids) => request('/entities/' + eid + '/bank-transactions/post', { method: 'POST', body: { transaction_ids: ids } }),
  deleteBankTransaction: (eid, id) => request('/entities/' + eid + '/bank-transactions/' + id, { method: 'DELETE' }),
  deleteBankBatch: (eid, batchId) => request('/entities/' + eid + '/bank-transactions/batch/' + batchId, { method: 'DELETE' }),

  // Bank Rec
  getReconciliations: (eid) => request('/entities/' + eid + '/reconciliations'),
  getCleared: (eid, code) => request('/entities/' + eid + '/cleared/' + code),
  createReconciliation: (eid, data) => request('/entities/' + eid + '/reconciliations', { method: 'POST', body: data }),

  setToken, getToken, clearToken,
};
