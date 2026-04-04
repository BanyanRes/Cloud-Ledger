const API_BASE = '/api';

function getToken() { return localStorage.getItem('cl_token'); }
function setToken(t) { localStorage.setItem('cl_token', t); }
function clearToken() { localStorage.removeItem('cl_token'); }

async function request(path, options = {}) {
  const token = getToken();
  const res = await fetch(API_BASE + path, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      ...(token ? { Authorization: 'Bearer ' + token } : {}),
      ...options.headers,
    },
    body: options.body ? JSON.stringify(options.body) : undefined,
  });
  if (res.status === 401) { clearToken(); window.location.reload(); return null; }
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || 'Request failed');
  return data;
}

export const api = {
  // Auth
  login: (email, password) => request('/auth/login', { method: 'POST', body: { email, password } }),
  signup: (name, email, password, role) => request('/auth/signup', { method: 'POST', body: { name, email, password, role } }),
  me: () => request('/auth/me'),

  // Users
  getUsers: () => request('/users'),
  deleteUser: (id) => request('/users/' + id, { method: 'DELETE' }),
  updateUser: (id, data) => request('/users/' + id, { method: 'PUT', body: data }),

  // Entities
  getEntities: () => request('/entities'),
  createEntity: (code, name) => request('/entities', { method: 'POST', body: { code, name } }),
  bulkCreateEntities: (entities) => request('/entities/bulk', { method: 'POST', body: { entities } }),
  deleteEntity: (id) => request('/entities/' + id, { method: 'DELETE' }),

  // Accounts
  getAccounts: (eid) => request('/entities/' + eid + '/accounts'),
  createAccount: (eid, data) => request('/entities/' + eid + '/accounts', { method: 'POST', body: data }),
  deleteAccount: (eid, code) => request('/entities/' + eid + '/accounts/' + code, { method: 'DELETE' }),

  // Journal Entries
  getEntries: (eid) => request('/entities/' + eid + '/entries'),
  createEntry: (eid, data) => request('/entities/' + eid + '/entries', { method: 'POST', body: data }),
  deleteEntry: (eid, id) => request('/entities/' + eid + '/entries/' + id, { method: 'DELETE' }),

  // Reports
  getBalances: (eid) => request('/entities/' + eid + '/balances'),
  getSummary: () => request('/summary'),

  // Bank Reconciliation
  getReconciliations: (eid) => request('/entities/' + eid + '/reconciliations'),
  getCleared: (eid, acctCode) => request('/entities/' + eid + '/cleared/' + acctCode),
  createReconciliation: (eid, data) => request('/entities/' + eid + '/reconciliations', { method: 'POST', body: data }),

  // Token management
  setToken, getToken, clearToken,
};
