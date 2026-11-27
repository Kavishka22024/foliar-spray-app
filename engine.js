// The phone's copy of the app (built by tools/build-engine.js; do not edit)
// The Google services the app's server code uses, made of plain JavaScript, so that the SAME code (all the .gs files) can run on the phone over the phone's own
// copy of the tables (see OFFLINE_REDESIGN_PLAN.md). Only what the code really uses is built. Nothing here talks to Google or to the network.
var __DEVICE__ = true; // server code may ask: typeof __DEVICE__ !== 'undefined'

var __LS = (function () {
  function LRange(sh, r, c, nr, nc) {
    this.sh = sh;
    this.r = r;
    this.c = c;
    this.nr = nr;
    this.nc = nc;
  }
  LRange.prototype.getValues = function () {
    var out = [];
    for (var i = 0; i < this.nr; i++) {
      var row = [];
      for (var j = 0; j < this.nc; j++) row.push(this.sh.cell(this.r + i, this.c + j));
      out.push(row);
    }
    return out;
  };
  LRange.prototype.setValues = function (v) {
    for (var i = 0; i < v.length; i++) for (var j = 0; j < v[i].length; j++) this.sh.set(this.r + i, this.c + j, v[i][j]);
    return this;
  };
  LRange.prototype.clearContent = function () {
    for (var i = 0; i < this.nr; i++) for (var j = 0; j < this.nc; j++) this.sh.set(this.r + i, this.c + j, '');
    return this;
  };
  LRange.prototype.setNumberFormat = function () {
    return this;
  };
  LRange.prototype.setFontWeight = function () {
    return this;
  };

  // d[r - 1][c - 1] is the cell in row r, column c. lastRow is the last row with something in it (kept up to date as cells are set).
  function LSheet(name) {
    this.name = name;
    this.d = [];
    this.maxRows = 1000;
    this.maxCols = 26;
    this.lastRow = 0;
    this.stale = false;
    this.version = 0;
  }
  LSheet.prototype.cell = function (r, c) {
    var row = this.d[r - 1];
    if (!row) return '';
    var v = row[c - 1];
    return v === undefined ? '' : v;
  };
  LSheet.prototype.set = function (r, c, x) {
    var row = this.d[r - 1];
    if (!row) {
      row = [];
      this.d[r - 1] = row;
    }
    row[c - 1] = x;
    if (x !== '' && x !== undefined && x !== null) {
      if (r > this.lastRow) this.lastRow = r;
    } else if (r === this.lastRow) this.stale = true;
    this.version++;
  };
  LSheet.prototype.getLastRow = function () {
    if (this.stale) {
      var i = this.d.length;
      while (i > 0) {
        var row = this.d[i - 1];
        if (
          row &&
          row.some(function (v) {
            return v !== '' && v !== undefined && v !== null;
          })
        )
          break;
        i--;
      }
      this.lastRow = i;
      this.stale = false;
    }
    return this.lastRow;
  };
  LSheet.prototype.getMaxRows = function () {
    return this.maxRows;
  };
  LSheet.prototype.getMaxColumns = function () {
    return this.maxCols;
  };
  LSheet.prototype.insertColumnsAfter = function (pos, n) {
    this.maxCols += n;
  };
  LSheet.prototype.getRange = function (r, c, nr, nc) {
    nr = nr === undefined ? 1 : nr;
    nc = nc === undefined ? 1 : nc;
    if (r + nr - 1 > this.maxRows) this.maxRows = r + nr - 1 + 500; // (the copy on the phone has no limit of its own)
    if (c + nc - 1 > this.maxCols) throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
    return new LRange(this, r, c, nr, nc);
  };
  LSheet.prototype.setFrozenRows = function () {};
  LSheet.prototype.deleteRow = function (r) {
    this.d.splice(r - 1, 1);
    this.stale = true;
    this.version++;
  };
  LSheet.prototype.deleteRows = function (from, n) {
    this.d.splice(from - 1, n);
    this.stale = true;
    this.version++;
  };
  // the copy kept on the phone: the header row, then the data rows
  LSheet.prototype.setHeaders = function (h) {
    this.d = [h.slice()];
    this.lastRow = 1;
    this.stale = false;
    this.maxCols = Math.max(this.maxCols, h.length);
    this.version++;
  };
  LSheet.prototype.importRows = function (headers, rows) {
    this.d = [headers.slice()];
    for (var i = 0; i < rows.length; i++) this.d.push(rows[i].slice());
    this.lastRow = this.d.length;
    this.stale = true;
    this.maxRows = Math.max(1000, this.d.length + 500);
    this.maxCols = Math.max(this.maxCols, headers.length);
    this.version++;
  };
  LSheet.prototype.exportRows = function (ncols) {
    var last = this.getLastRow(),
      out = [];
    for (var i = 1; i < last; i++) {
      var row = this.d[i] || [],
        o = new Array(ncols);
      for (var c = 0; c < ncols; c++) o[c] = row[c] === undefined ? '' : row[c];
      out.push(o);
    }
    return out;
  };

  function LSS() {
    this.sheets = [];
  }
  LSS.prototype.getId = function () {
    return 'device-copy';
  };
  LSS.prototype.getSheets = function () {
    return this.sheets;
  };
  LSS.prototype.getSheetByName = function (n) {
    for (var i = 0; i < this.sheets.length; i++) if (this.sheets[i].name === n) return this.sheets[i];
    return null;
  };
  LSS.prototype.insertSheet = function (n) {
    var s = new LSheet(n);
    this.sheets.push(s);
    return s;
  };
  return { LSheet: LSheet, LSS: LSS };
})();

var __SS = new __LS.LSS();
var __props = { SSID: 'device-copy' },
  __cache = {},
  __tz = 'Asia/Colombo';
var SpreadsheetApp = {
  getActiveSpreadsheet: function () {
    return __SS;
  },
  openById: function () {
    return __SS;
  }
};
var PropertiesService = {
  getScriptProperties: function () {
    return {
      getProperty: function (k) {
        return __props[k] === undefined ? null : __props[k];
      },
      setProperty: function (k, v) {
        __props[k] = String(v);
      }
    };
  }
};
var LockService = {
  getScriptLock: function () {
    return { waitLock: function () {}, releaseLock: function () {} };
  }
};
var CacheService = {
  getScriptCache: function () {
    return {
      get: function (k) {
        return __cache[k] === undefined ? null : __cache[k];
      },
      put: function (k, v) {
        __cache[k] = String(v);
      },
      remove: function (k) {
        delete __cache[k];
      }
    };
  }
};
var Session = {
  getScriptTimeZone: function () {
    return __tz;
  },
  getActiveUser: function () {
    return {
      getEmail: function () {
        return '';
      }
    };
  },
  getEffectiveUser: function () {
    return {
      getEmail: function () {
        return '';
      }
    };
  }
};
var __unavailable = function () {
  throw new Error('Not available on the phone.');
};
var __fmtDate = function (date, tz, pattern) {
  var parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: tz,
    hourCycle: 'h23',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  })
    .formatToParts(date)
    .reduce(function (o, p) {
      o[p.type] = p.value;
      return o;
    }, {});
  return pattern
    .replace(/yyyy/g, parts.year)
    .replace(/MM/g, parts.month)
    .replace(/dd/g, parts.day)
    .replace(/HH/g, parts.hour)
    .replace(/mm/g, parts.minute)
    .replace(/ss/g, parts.second)
    .replace(/'/g, '');
};
var Utilities = {
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  getUuid: function () {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      var r = (Math.random() * 16) | 0;
      return (c === 'x' ? r : (r & 3) | 8).toString(16);
    });
  },
  formatDate: __fmtDate,
  computeDigest: __unavailable,
  computeHmacSha256Signature: __unavailable,
  base64Encode: __unavailable,
  base64EncodeWebSafe: __unavailable
};
var Logger = {
  log: function () {}
};
var ContentService = {
  createTextOutput: function (t) {
    return {
      setMimeType: function () {
        return this;
      },
      getContent: function () {
        return t;
      }
    };
  },
  MimeType: { JSON: 'json' }
};

// ---- Auth.gs
/** Authentication, sessions and per-task authorisation. */

// ---------- password hashing (salted, iterated HMAC-SHA256 with a server-side pepper) ----------
function pepper_() {
  const p = PropertiesService.getScriptProperties().getProperty('PEPPER');
  if (!p) fail_('App is not set up. Run setupApp() once from the script editor.', 'NOT_SETUP');
  return p;
}
function hashWith_(pw, salt, iter) {
  const key = pepper_();
  let h = Utilities.base64Encode(Utilities.computeHmacSha256Signature(salt + ':' + pw, key));
  for (let i = 1; i < iter; i++) {
    h = Utilities.base64Encode(Utilities.computeHmacSha256Signature(h + ':' + salt, key));
  }
  return h;
}
function makeHash_(pw) {
  const salt = randomString_(16);
  return 'v1$' + PW_ITERATIONS + '$' + salt + '$' + hashWith_(pw, salt, PW_ITERATIONS);
}
function constEq_(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}
function verifyHash_(pw, stored) {
  const parts = String(stored || '').split('$');
  if (parts.length !== 4 || parts[0] !== 'v1') return false;
  return constEq_(hashWith_(pw, parts[2], parseInt(parts[1], 10)), parts[3]);
}
/** Spend the same time for unknown users so response time does not reveal which usernames exist. */
function dummyHash_() {
  hashWith_('x', 'dummysaltdummysalt', PW_ITERATIONS);
}

function checkPasswordPolicy_(pw, username, isAdminRole) {
  pw = String(pw || '');
  const min = isAdminRole ? 10 : 6;
  if (pw.length < min) fail_('Password must be at least ' + min + ' characters.', 'VALIDATION');
  if (pw.length > 100) fail_('Password is too long.', 'VALIDATION');
  if (username && pw.toLowerCase() === String(username).toLowerCase()) fail_('Password cannot be the same as the username.', 'VALIDATION');
}

// ---------- sessions (random token; only its hash is used as the cache key) ----------
function sessKey_(token) {
  return 'S_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token));
}
function createSession_(userId) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put(sessKey_(token), JSON.stringify({ uid: userId }), SESSION_SECONDS);
  return token;
}
function destroySession_(token) {
  if (typeof token === 'string' && token.length === 64) CacheService.getScriptCache().remove(sessKey_(token));
}

/** Resolve a token to {user, role, tasks[], isAdmin}. Throws when invalid. A phone's long-lived key (dk_...) is accepted only inside the sync doors (Sync.gs sets DEVICE_AUTH_). */
function ctx_(token) {
  if (isDeviceKey_(token)) {
    if (!DEVICE_AUTH_) fail_('Please log in.', 'NO_SESSION');
    const c = ctxOfUser_(deviceUserId_(token), token);
    if (!c) fail_('This phone is no longer signed in. Please log in again.', 'DEVICE_REVOKED');
    return c;
  }
  if (typeof token !== 'string' || token.length !== 64) fail_('Please log in.', 'NO_SESSION');
  const raw = CacheService.getScriptCache().get(sessKey_(token));
  if (!raw) fail_('Your session has expired. Please log in again.', 'SESSION_EXPIRED');
  const c = ctxOfUser_(JSON.parse(raw).uid, token);
  if (!c) {
    destroySession_(token);
    fail_('Account is not active.', 'SESSION_EXPIRED');
  }
  return c;
}
/** The context of an account (no session involved): its row, its job role, the tasks it holds. null when the account is gone or inactive. The phone's copy of the app uses this too. */
function ctxOfUser_(uid, token) {
  const user = dbAll_(SHEETS.USERS).filter(function (u) {
    return u.user_id === uid;
  })[0];
  if (!user || !user.active) return null;
  const role = dbAll_(SHEETS.ROLES).filter(function (r) {
    return r.role_id === user.role_id;
  })[0];
  if (!role || !role.active) fail_('Your job role is not active. Contact the Admin.', 'FORBIDDEN');
  const assignable = TASKS.map(function (t) {
    return t.id;
  });
  const tasks = role.is_admin
    ? assignable
    : expandTasks_(
        String(role.tasks || '')
          .split(',')
          .map(function (s) {
            return s.trim();
          })
      ).filter(function (t) {
        return assignable.indexOf(t) >= 0;
      });
  if (!role.is_admin && (role.field_scope || 'assigned') !== 'all')
    WHOLE_ESTATE_TASKS.forEach(function (w) {
      const i = tasks.indexOf(w);
      if (i >= 0) tasks.splice(i, 1);
    }); // a limited role never sees the whole-estate screens
  if (!role.is_admin && tasks.indexOf('history.view') < 0 && (tasks.indexOf('spray.entry') >= 0 || tasks.indexOf('pluck.entry') >= 0))
    tasks.push('history.view'); // whoever enters data can look back at it
  return { user: user, role: role, tasks: tasks, isAdmin: !!role.is_admin, token: token };
}

/** Require a valid session that holds at least one of the given tasks (any task if none given). */
function authz_(token, anyOf, allowMustChange) {
  const c = ctx_(token);
  if (c.user.must_change && !allowMustChange) fail_('You must change your password first.', 'PASSWORD_CHANGE_REQUIRED');
  if (anyOf && anyOf.length) {
    const ok = anyOf.some(function (t) {
      return c.tasks.indexOf(t) >= 0;
    });
    if (!ok) fail_('You do not have permission for this action.', 'FORBIDDEN');
  }
  return c;
}
function hasTask_(c, t) {
  return c.tasks.indexOf(t) >= 0;
}

function meOf_(c) {
  const estate = (dbAll_(SHEETS.ESTATE)[0] || {}).name || '';
  const metaById = {};
  TASKS.forEach(function (t) {
    metaById[t.id] = t;
  });
  return {
    user_id: c.user.user_id,
    username: c.user.username,
    display_name: c.user.display_name,
    role_name: c.role.role_name,
    is_admin: c.isAdmin,
    must_change: !!c.user.must_change,
    field_scope: c.isAdmin ? 'all' : c.role.field_scope || 'assigned',
    tasks: c.tasks.map(function (id) {
      return { id: id, label: metaById[id].label, group: metaById[id].group };
    }),
    app_name: APP_NAME,
    version: APP_VERSION,
    today: todayStr_(),
    estate: estate
  };
}

// ---------- client API ----------
function api_login(username, password) {
  return wrap_(function () {
    const uname = String(username || '')
      .trim()
      .toLowerCase();
    const pw = String(password || '');
    if (!uname || !pw || uname.length > 50 || pw.length > 100) fail_('Invalid username or password.', 'LOGIN');
    const named = function (list) {
      return list.filter(function (u) {
        return u.username === uname;
      })[0];
    };
    const seen = named(dbAll_(SHEETS.USERS));
    const seenGood = seen ? verifyHash_(pw, seen.pw_hash) : (dummyHash_(), false); // the slow part: done before the lock, so logins do not queue behind each other
    const token = withLock_(function () {
      const user = named(dbAll_(SHEETS.USERS)); // read again under the lock: the counters must not be lost
      if (!user) {
        audit_(uname, 'LOGIN_FAIL', 'user', '', 'unknown username');
        fail_('Invalid username or password.', 'LOGIN');
      }
      const now = nowIso_();
      if (user.locked_until && user.locked_until > now) {
        audit_(uname, 'LOGIN_BLOCKED', 'user', user.user_id, 'locked');
        fail_('This account is temporarily locked after too many failed attempts. Try again later or ask the Admin to unlock it.', 'LOCKED');
      }
      const good = seen && seen.pw_hash === user.pw_hash ? seenGood : verifyHash_(pw, user.pw_hash); // (the password was changed in between: check it again)
      if (!good || !user.active) {
        const fails = (user.failed_count || 0) + 1;
        const patch = { failed_count: fails };
        if (fails >= LOGIN_MAX_FAILS) {
          patch.locked_until = Utilities.formatDate(new Date(Date.now() + LOGIN_LOCK_MINUTES * 60000), tz_(), "yyyy-MM-dd'T'HH:mm:ss");
          patch.failed_count = 0;
        }
        dbUpdate_(SHEETS.USERS, user._row, patch);
        audit_(uname, 'LOGIN_FAIL', 'user', user.user_id, good ? 'inactive' : 'bad password');
        fail_('Invalid username or password.', 'LOGIN');
      }
      dbUpdate_(SHEETS.USERS, user._row, { failed_count: 0, locked_until: '', last_login: now });
      audit_(uname, 'LOGIN', 'user', user.user_id, '');
      return createSession_(user.user_id);
    });
    return { token: token, me: meOf_(ctx_(token)) };
  });
}

function api_logout(token) {
  return wrap_(function () {
    destroySession_(token);
    return true;
  });
}

function api_me(token) {
  return wrap_(function () {
    return meOf_(authz_(token, null, true));
  });
}

function api_changePassword(token, oldPw, newPw) {
  return wrap_(function () {
    const c = authz_(token, null, true);
    return withLock_(function () {
      const user = dbAll_(SHEETS.USERS).filter(function (u) {
        return u.user_id === c.user.user_id;
      })[0];
      if (!verifyHash_(String(oldPw || ''), user.pw_hash)) fail_('Current password is incorrect.', 'VALIDATION');
      checkPasswordPolicy_(newPw, user.username, c.isAdmin);
      if (String(newPw) === String(oldPw)) fail_('New password must be different from the current one.', 'VALIDATION');
      dbUpdate_(SHEETS.USERS, user._row, { pw_hash: makeHash_(String(newPw)), must_change: false, updated_at: nowIso_() });
      audit_(user.username, 'PASSWORD_CHANGE', 'user', user.user_id, '');
      return meOf_(ctx_(token));
    });
  });
}

// ---- Config.gs
/**
 * CAFT Foliar Spray Experiment - configuration: sheet schemas, task registry, parameter registry.
 * Every user-facing feature is a TASK. A job ROLE is a set of tasks plus a field scope. Admin has everything.
 */
const APP_NAME = 'CAFT Foliar Spray Experiment';
const APP_VERSION = '0.2.0';

const SHEETS = {
  USERS: 'Users',
  ROLES: 'Roles',
  ESTATE: 'Estate',
  DIVISIONS: 'Divisions',
  FIELDS: 'Fields',
  PARAMS: 'Parameters',
  SPRAY: 'SprayRecords',
  SPRAY_EVENTS: 'SprayEvents',
  LEAF: 'LeafRecords',
  AUDIT: 'Audit',
  PLANS: 'FoliarPlans',
  TEMPLATE: 'FoliarTemplate',
  STAMPS: 'FoliarStamps',
  CHANGES: 'Changes',
  DEVICES: 'Devices',
  OPLOG: 'OpLog'
};

const SCHEMA = {
  Users: [
    'user_id',
    'username',
    'display_name',
    'role_id',
    'pw_hash',
    'active',
    'must_change',
    'failed_count',
    'locked_until',
    'last_login',
    'created_at',
    'created_by',
    'updated_at',
    'default_username',
    'prev_usernames',
    'division_ids'
  ], // division_ids: whole divisions this account is responsible for (comma list)
  Roles: ['role_id', 'role_name', 'description', 'tasks', 'field_scope', 'is_admin', 'active', 'created_at', 'updated_at'],
  Estate: ['estate_id', 'name', 'created_at'],
  Divisions: ['division_id', 'name', 'seq', 'active', 'created_at'],
  Fields: [
    'field_id',
    'division_id',
    'seq',
    'code',
    'name',
    'area_ha',
    'last_prune_date',
    'tea_type',
    'method',
    'afo_user_id',
    'active',
    'created_at',
    'updated_at',
    'drone_ha'
  ],
  Parameters: ['param_id', 'key', 'value', 'effective_from', 'note', 'set_by', 'set_at'],
  SprayRecords: [
    'record_id',
    'date',
    'field_id',
    'method',
    'prune_date',
    'start_time',
    'end_time',
    'hours',
    'units',
    'area_ha',
    'flags',
    'status',
    'void_reason',
    'entered_by',
    'entered_at',
    'edited_by',
    'edited_at',
    'event_id',
    'kg_applied',
    'pct_total',
    'zinc_kg',
    'epsom_kg',
    'wuxal_kg',
    'fourth_name',
    'fourth_kg'
  ], // the last five: this field's share of the event's four compounds (blank on an older event that recorded one total kg)
  // one row per spray EVENT (one mix of foliar applied by one crew); its fields are the SprayRecords rows with the same event_id
  SprayEvents: [
    'event_id',
    'date',
    'method',
    'kg_applied',
    'men',
    'drones',
    'start_time',
    'end_time',
    'hours',
    'total_ha',
    'flags',
    'status',
    'void_reason',
    'entered_by',
    'entered_at',
    'edited_by',
    'edited_at',
    'zinc_kg',
    'epsom_kg',
    'wuxal_kg',
    'fourth_name',
    'fourth_kg',
    'plan_id'
  ], // kg_applied = the sum of the four compounds; blank compounds = an older event that recorded one total kg
  // the foliar plan: what is mixed for ONE hectare, from an effective date on (Zinc, Epsom salt, Wuxal and Urea or T65)
  // the monthly drone spray schedule (the plan file): one row per planned block of one field's drone part: which weekday (1 = Monday), which week of the month (1st, 2nd ... occurrence of that weekday), morning or evening, how many ha
  FoliarTemplate: ['slot_id', 'field_id', 'part', 'weekday', 'week_no', 'half', 'ha', 'status'],
  // the schedule of ONE month: a copy of the template made when the month is first opened; a stamp is dragged to another day / half-day by people with the scheduling task
  FoliarStamps: [
    'stamp_id',
    'month',
    'field_id',
    'part',
    'ha',
    'date',
    'half',
    'orig_date',
    'orig_half',
    'template_id',
    'status',
    'moved_by',
    'moved_at',
    'created_at'
  ],
  FoliarPlans: [
    'plan_id',
    'effective_from',
    'zinc_kg_ha',
    'epsom_kg_ha',
    'wuxal_kg_ha',
    'fourth_name',
    'fourth_kg_ha',
    'note',
    'status',
    'set_by',
    'set_at',
    'void_reason',
    'voided_by',
    'voided_at'
  ],
  LeafRecords: [
    'record_id',
    'date',
    'field_id',
    'method',
    'prune_date',
    'kg',
    'area_ha',
    'flags',
    'status',
    'void_reason',
    'entered_by',
    'entered_at',
    'edited_by',
    'edited_at'
  ],
  Audit: ['ts', 'user', 'action', 'entity', 'entity_id', 'details'],
  // the change feed the phones read (see Sync.gs): one row per row written in a synced table; at = milliseconds, op = u (added or changed) | d (deleted)
  Changes: ['at', 'table', 'key', 'op'],
  // one row per phone that holds a copy of the data: only a hash of its long-lived key is kept
  Devices: ['device_id', 'key_hash', 'user_id', 'label', 'created_at', 'last_used', 'expires', 'revoked'],
  // the operations phones sent (see api_submitOp): an operation that arrives twice is done once
  OpLog: ['op_id', 'fn', 'user_id', 'result', 'at']
};

// Columns stored as numbers / booleans. Every other column is formatted as plain text in the sheet.
const NUMERIC_COLS = [
  'seq',
  'area_ha',
  'hours',
  'units',
  'kg',
  'failed_count',
  'kg_applied',
  'men',
  'drones',
  'total_ha',
  'pct_total',
  'drone_ha',
  'zinc_kg',
  'epsom_kg',
  'wuxal_kg',
  'fourth_kg',
  'zinc_kg_ha',
  'epsom_kg_ha',
  'wuxal_kg_ha',
  'fourth_kg_ha',
  'weekday',
  'week_no',
  'ha',
  'at'
];
/** Starting password for every non-Admin account (new, seeded or reset). The user must choose their own at first login. */
const DEFAULT_PASSWORD = '1111';
const BOOL_COLS = ['active', 'must_change', 'is_admin'];
const DATE_COLS = ['date', 'effective_from', 'last_prune_date', 'prune_date', 'orig_date'];
const TIME_COLS = ['start_time', 'end_time'];

const METHODS = ['Drone', 'Manual'];
/** The foliar is a mix of four compounds in water. Zinc, Epsom salt and Wuxal are always in it; the fourth is Urea or T65, depending on the month. */
const MIX_FIXED = [
  { key: 'zinc', name: 'Zinc' },
  { key: 'epsom', name: 'Epsom salt' },
  { key: 'wuxal', name: 'Wuxal' }
];
const FOURTH_COMPOUNDS = ['Urea', 'T65'];
/** The price parameter of each compound. */
const COMPOUND_PRICE_KEY = {
  Zinc: 'price_zinc_per_kg',
  'Epsom salt': 'price_epsom_per_kg',
  Wuxal: 'price_wuxal_per_kg',
  Urea: 'price_urea_per_kg',
  T65: 'price_t65_per_kg'
};
/** A field is sprayed by drone on part of it (Fields.drone_ha) and by hand on the rest. Its class: Drone (all drone), Manual (none by drone) or Mixed. */
const FIELD_METHODS = ['Drone', 'Manual', 'Mixed'];
const TEA_TYPES = ['VP', 'SD'];
/** Who a role can record for: assigned = fields assigned to that user; drone = every field with a drone part; all = every field. */
const FIELD_SCOPES = ['assigned', 'drone', 'all'];

/** Task registry. sensitive:true = only an Admin can grant it. */
const TASKS = [
  {
    id: 'spray.entry',
    group: 'Data entry',
    label: 'Foliar spray',
    desc: "Data entry: create foliar spray events (kg applied, men, drones, fields covered) for the fields in the role's field scope."
  },
  {
    id: 'foliar.plan',
    group: 'Data entry',
    label: 'Foliar planning',
    desc: 'Set what is mixed for one hectare (kg of Zinc, Epsom salt, Wuxal and Urea or T65) and from which date it applies. The spray entry screen then suggests the kg of each compound from the hectares sprayed.'
  },
  {
    id: 'foliar.calendar',
    group: 'Data entry',
    label: 'Foliar calendar',
    desc: "See the month's drone spray plan as a calendar of stamps: grey = planned, green = sprayed, red = planned day passed and not sprayed."
  },
  {
    id: 'foliar.schedule',
    group: 'Data entry',
    label: 'Foliar scheduling',
    desc: "Edit the month's drone spray plan: drag a grey or red stamp to another day (morning or evening). Includes seeing the calendar."
  },
  {
    id: 'pluck.entry',
    group: 'Data entry',
    label: 'Plucking',
    desc: "Data entry: enter green leaf weight and plucked area for the fields in the role's field scope."
  },
  {
    id: 'history.view',
    group: 'Data entry',
    label: 'Field history',
    auto: true,
    desc: "The rounds and entries of each field in the role's field scope. Given automatically to every role that can enter foliar spray or plucking."
  },
  {
    id: 'monitor.operation',
    group: 'Monitoring',
    label: 'Operation monitoring',
    desc: 'Plucking rounds and foliar rounds of every field on a timeline: how the work is going.'
  },
  {
    id: 'monitor.experiment',
    group: 'Monitoring',
    label: 'Experiment monitoring',
    wholeEstate: true,
    desc: 'Drone against manual spraying: cost, yield, money and the graphs for the chosen period. It compares the whole estate, so only a role that can see every field can have it.'
  },
  { id: 'records.manage', group: 'Administration', label: 'Records', desc: 'See every spray and plucking record; edit or void any of them.' },
  {
    id: 'estate.manage',
    group: 'Administration',
    label: 'Estate setup',
    desc: "Create the estate, divisions and fields and set each field's details."
  },
  {
    id: 'params.manage',
    group: 'Administration',
    label: 'Parameters and rates',
    desc: 'Cost rates, green leaf price, experiment windows and entry limits.'
  },
  { id: 'users.manage', group: 'Administration', label: 'Users', sensitive: true, desc: 'Create users, assign fields to AFOs, reset passwords.' },
  {
    id: 'roles.manage',
    group: 'Administration',
    label: 'Job roles',
    sensitive: true,
    desc: 'Create job roles and choose which tasks each role can use.'
  },
  { id: 'audit.view', group: 'Administration', label: 'Audit log', desc: 'View the log of who changed what and when.' }
];

/** Tasks that are about the whole estate: a job role that is limited to some fields (scope assigned or drone) cannot hold them, and an account never gets them even if an old role still lists them. */
const WHOLE_ESTATE_TASKS = ['monitor.experiment'];

/** Monitoring used to be ONE task ('dashboard.view'). A role saved with it holds both parts, so nobody loses a screen; saving the role writes the two new tasks. */
const LEGACY_TASKS = { 'dashboard.view': ['monitor.operation', 'monitor.experiment'] };
/** A list of task ids with every old id replaced by the tasks it became, no duplicates, order kept. */
function expandTasks_(list) {
  const out = [];
  (list || []).forEach(function (t) {
    (LEGACY_TASKS[t] || [t]).forEach(function (x) {
      if (out.indexOf(x) < 0) out.push(x);
    });
  });
  return out;
}

/** Parameter registry. dated:true = value has an effective-from date (rates); false = latest value wins. */
const PARAMS = [
  // --- Cost rates (dated)
  {
    key: 'drone_cost_basis',
    label: 'Drone cost basis',
    group: 'Cost rates',
    type: 'enum',
    options: ['per_hour', 'per_ha'],
    def: 'per_hour',
    dated: true,
    help: 'per_ha: the drone is paid per hectare sprayed, cost = ha sprayed x the per-ha charge. per_hour: cost = drones x hours x the hourly rate.'
  },
  {
    key: 'drone_rate_per_hour',
    label: 'Drone cost per drone-hour',
    group: 'Cost rates',
    type: 'number',
    unit: 'LKR/hr',
    def: 0,
    min: 0,
    dated: true,
    help: 'All-in rate (owned) or hire rate, including the pilot. Used when basis is per_hour.'
  },
  {
    key: 'drone_rate_per_ha',
    label: 'Drone cost per ha',
    group: 'Cost rates',
    type: 'number',
    unit: 'LKR/ha',
    def: 0,
    min: 0,
    dated: true,
    help: 'What the drone charges for one hectare sprayed (pilot and machine included). Used when the basis is per_ha.'
  },
  {
    key: 'worker_rate_per_hour',
    label: 'Manual labour cost per man-hour',
    group: 'Cost rates',
    type: 'number',
    unit: 'LKR/hr',
    def: 0,
    min: 0,
    dated: true,
    help: 'Day wage divided by paid hours per day.'
  },
  { key: 'product_name', label: 'Spray product', group: 'Cost rates', type: 'text', def: 'ZnSO4', dated: false },
  {
    key: 'price_zinc_per_kg',
    label: 'Zinc price per kg',
    group: 'Cost rates',
    type: 'number',
    unit: 'LKR/kg',
    def: 0,
    min: 0,
    dated: true,
    help: 'What one kg of Zinc costs. The foliar cost of a spray event is the kg of each of its four compounds x the price of that compound on the day of the spray.'
  },
  { key: 'price_epsom_per_kg', label: 'Epsom salt price per kg', group: 'Cost rates', type: 'number', unit: 'LKR/kg', def: 0, min: 0, dated: true },
  { key: 'price_wuxal_per_kg', label: 'Wuxal price per kg', group: 'Cost rates', type: 'number', unit: 'LKR/kg', def: 0, min: 0, dated: true },
  { key: 'price_urea_per_kg', label: 'Urea price per kg', group: 'Cost rates', type: 'number', unit: 'LKR/kg', def: 0, min: 0, dated: true },
  { key: 'price_t65_per_kg', label: 'T65 price per kg', group: 'Cost rates', type: 'number', unit: 'LKR/kg', def: 0, min: 0, dated: true },
  {
    key: 'product_unit_price',
    label: 'Foliar price per kg (older events, one total kg)',
    group: 'Cost rates',
    type: 'number',
    unit: 'LKR/kg',
    def: 0,
    min: 0,
    dated: true,
    help: 'Only for spray events recorded before the four compounds were entered separately (they have one total kg). New events are priced compound by compound with the prices above.'
  },
  {
    key: 'standard_dose_per_ha',
    label: 'Standard product dose per ha',
    group: 'Cost rates',
    type: 'number',
    unit: 'kg/ha',
    def: 0,
    min: 0,
    dated: true,
    help: 'Same for both methods. TRI ZnSO4: 1.5 or 2.75 kg/ha per application.'
  },
  {
    key: 'gl_price_per_kg',
    label: 'Green leaf price',
    group: 'Revenue',
    type: 'number',
    unit: 'LKR/kg',
    def: 0,
    min: 0,
    dated: true,
    help: 'The price of one kg of green leaf. Used for the yield revenue per ha on the Experiment monitoring screen (kg plucked on a day x the price of that day).'
  },
  {
    key: 'gl_variable_cost_per_kg',
    label: 'Variable cost per extra kg of green leaf',
    group: 'Revenue',
    type: 'number',
    unit: 'LKR/kg',
    def: 0,
    min: 0,
    dated: true,
    help: 'Plucking, transport and handling of extra leaf.'
  },
  // --- Experiment settings (latest value wins)
  { key: 'experiment_name', label: 'Experiment name', group: 'Experiment', type: 'text', def: APP_NAME, dated: false },
  { key: 'experiment_start', label: 'Experiment start date', group: 'Experiment', type: 'date', def: '', dated: false },
  { key: 'experiment_end', label: 'Experiment end date (blank = open-ended)', group: 'Experiment', type: 'date', def: '', dated: false },
  { key: 'baseline_start', label: 'Baseline window start', group: 'Experiment', type: 'date', def: '', dated: false },
  { key: 'baseline_end', label: 'Baseline window end', group: 'Experiment', type: 'date', def: '', dated: false },
  { key: 'analysis_start', label: 'Analysis window start', group: 'Experiment', type: 'date', def: '', dated: false },
  { key: 'analysis_end', label: 'Analysis window end (blank = today)', group: 'Experiment', type: 'date', def: '', dated: false },
  {
    key: 'pluck_target_gap_days',
    label: 'Plucking round target gap (days between the first days of two rounds)',
    group: 'Monitoring',
    type: 'number',
    unit: 'days',
    def: 7,
    min: 1,
    max: 60,
    dated: false,
    help: 'On the Plucking rounds screen a gap longer than this is shown red, with the days beyond it.'
  },
  {
    key: 'spray_target_gap_days',
    label: 'Foliar round target gap (days between the first days of two rounds)',
    group: 'Monitoring',
    type: 'number',
    unit: 'days',
    def: 30,
    min: 1,
    max: 120,
    dated: false,
    help: 'On the Foliar rounds screen a gap longer than this is shown red, with the days beyond it.'
  },
  {
    key: 'spray_round_reset_days',
    label: 'Foliar: days allowed between two spray events of the same round',
    group: 'Monitoring',
    type: 'number',
    unit: 'days',
    def: 28,
    min: 1,
    max: 120,
    dated: false,
    help: 'The drone plan is monthly and a part that was not finished is sprayed on a later week, so a round stays open for a month. Spray events on the same part of a field that are this many days apart or less belong to the same round. The next event more than this many days after the last one starts a NEW round, and a round not finished by then stays unfinished (the app card starts again from 0%).'
  },
  {
    key: 'pluck_round_reset_days',
    label: 'Plucking: days allowed between two pluckings of the same round',
    group: 'Monitoring',
    type: 'number',
    unit: 'days',
    def: 2,
    min: 1,
    max: 120,
    dated: false,
    help: 'Pluckings of the same part of a field that are this many days apart or less belong to the same round. The next plucking more than this many days after the last one starts a NEW round, and a round not finished by then stays unfinished (the app card starts again from 0).'
  },
  // --- Data entry controls
  {
    key: 'edit_window_days',
    label: 'Days a person may void their own entry',
    group: 'Data entry',
    type: 'number',
    unit: 'days',
    def: 3,
    min: 0,
    dated: false
  },
  {
    key: 'sync_grace_days',
    label: 'Days back an entry saved offline may still be uploaded',
    group: 'Data entry',
    type: 'number',
    unit: 'days',
    def: 3,
    min: 0,
    max: 14,
    dated: false,
    help: 'People can only enter data for today. An entry made offline today can still upload for this many days afterwards.'
  },
  {
    key: 'lock_before_date',
    label: 'Lock entries on or before this date',
    group: 'Data entry',
    type: 'date',
    def: '',
    dated: false,
    help: 'Nobody can add, edit or void records dated on or before this. Blank = no lock.'
  },
  {
    key: 'foliar_plan_tolerance_pct',
    label: 'Foliar: kg of a compound may differ from the plan by',
    group: 'Data entry',
    type: 'number',
    unit: '%',
    def: 10,
    min: 0,
    max: 100,
    dated: false,
    help: 'A spray event whose kg of a compound is further than this from the foliar plan (kg per ha x ha sprayed) is flagged for review, not blocked.'
  },
  {
    key: 'foliar_block_max_ha',
    label: 'Foliar calendar: most ha in one morning or evening block',
    group: 'Foliar schedule',
    type: 'number',
    unit: 'ha',
    def: 5,
    min: 0.5,
    max: 50,
    dated: false,
    help: 'A stamp can only be dropped into a morning or evening block if the block then holds this many ha or less.'
  },
  {
    key: 'foliar_day_max_ha',
    label: 'Foliar calendar: most ha in one day',
    group: 'Foliar schedule',
    type: 'number',
    unit: 'ha',
    def: 9,
    min: 1,
    max: 100,
    dated: false,
    help: 'The most the drones can spray in a day (morning and evening together). Moves that would go above it are refused, and events above it are flagged.'
  },
  {
    key: 'drone_ha_per_hr_max',
    label: 'Drone ha/hour, highest plausible',
    group: 'Data entry',
    type: 'number',
    unit: 'ha/hr',
    def: '',
    min: 0,
    dated: false,
    help: 'Entries above this are flagged for review, not blocked. Blank = no check.'
  },
  {
    key: 'manual_ha_per_hr_max',
    label: 'Manual ha/hour, highest plausible',
    group: 'Data entry',
    type: 'number',
    unit: 'ha/hr',
    def: '',
    min: 0,
    dated: false
  },
  {
    key: 'leaf_kg_per_ha_max',
    label: 'Green leaf kg per plucked ha, highest plausible',
    group: 'Data entry',
    type: 'number',
    unit: 'kg/ha',
    def: '',
    min: 0,
    dated: false
  }
];

/** Roles created automatically at setup (the Admin can edit them, or add more). */
const SEED_ROLES = [
  {
    name: 'Assistant Field Officer',
    description: 'Enters foliar spray and plucking for the fields assigned to them.',
    tasks: ['spray.entry', 'pluck.entry'],
    scope: 'assigned'
  },
  { name: 'Drone Pilot', description: 'Enters foliar spray details for drone fields.', tasks: ['spray.entry'], scope: 'drone' },
  { name: 'Viewer', description: 'Sees the monitoring screens only.', tasks: ['monitor.operation', 'monitor.experiment'], scope: 'all' }
];

const LOGIN_MAX_FAILS = 5;
const LOGIN_LOCK_MINUTES = 15;
const SESSION_SECONDS = 21600; // 6 hours, the CacheService maximum
const PW_ITERATIONS = 500;

// ---- DataCheck.gs
/** DATA CHECK (owner only, read-only): is what the people entered stored correctly and consistently, and can the dashboard price it?
 *
 *  Run checkData() from the Apps Script editor as the script owner. It writes NOTHING. It reads each sheet once (checkContext_) and runs the checks of DATA_CHECKS one after the other; it reports, with a few examples of each:
 *    PROBLEMS  things that make a dashboard number wrong (a spray line whose event is missing or voided, an event whose fields do not add up to its hectares / kg / hours,
 *              a round that covers more than the field part, a plucking record on a method the field does not have, two plucking records for the same part on the same day,
 *              a sheet whose header row no longer matches the app, a number stored as text ...)
 *    WARNINGS  things that need a look but do not break the arithmetic (records on a day with no cost rate or price in force, entries flagged as implausible, spray lines with no event)
 *  The same function runs in the tests (tests/data-integrity.js) on the simulated year, and on deliberately damaged data to prove each check fires. */
function checkData() {
  requireOwner_();
  const r = checkData_(),
    t = formatDataCheck_(r);
  Logger.log(t);
  console.log(t);
  return t;
}

function formatDataCheck_(r) {
  const out = [
    'DATA CHECK  ' +
      (r.problems.length ? r.problems.length + ' kind(s) of PROBLEM' : 'no problems') +
      ', ' +
      r.warnings.length +
      ' kind(s) of warning'
  ];
  out.push(
    'Checked: ' +
      Object.keys(r.counts)
        .map(function (k) {
          return r.counts[k] + ' ' + k;
        })
        .join(', ')
  );
  r.problems.forEach(function (p) {
    out.push('PROBLEM  ' + p.n + ' x ' + p.cat + '  e.g. ' + p.ex.join(' | '));
  });
  r.warnings.forEach(function (p) {
    out.push('warning  ' + p.n + ' x ' + p.cat + '  e.g. ' + p.ex.join(' | '));
  });
  return out.join('\n');
}

/** What the checks collect: PROBLEMS (a dashboard number would be wrong) and WARNINGS (need a look), each kind with a count and the first few examples. */
function makeReporter_() {
  const P = {},
    W = {};
  const adder = function (map) {
    return function (cat, msg) {
      const x = map[cat] || (map[cat] = { cat: cat, n: 0, ex: [] });
      x.n++;
      if (x.ex.length < 4) x.ex.push(msg);
    };
  };
  const list = function (o) {
    return Object.keys(o).map(function (k) {
      return o[k];
    });
  };
  return {
    bad: adder(P),
    warn: adder(W),
    problems: function () {
      return list(P);
    },
    warnings: function () {
      return list(W);
    }
  };
}

/** Every sheet the checks need, read once, with the groupings several of them use. */
function checkContext_() {
  const today = todayStr_(),
    fieldBy = fieldIndex_(),
    events = dbAll_(SHEETS.SPRAY_EVENTS),
    lines = dbAll_(SHEETS.SPRAY),
    leaf = dbAll_(SHEETS.LEAF),
    fieldRows = dbAll_(SHEETS.FIELDS);
  const activeEv = {};
  events.forEach(function (e) {
    if (e.status === 'Active') activeEv[e.event_id] = e;
  });
  return {
    today: today,
    params: getParams_(today),
    fieldBy: fieldBy,
    fieldRows: fieldRows,
    events: events,
    lines: lines,
    leaf: leaf,
    activeEv: activeEv,
    activeLines: lines.filter(function (l) {
      return l.status === 'Active' && l.event_id && activeEv[l.event_id];
    }),
    activeLeaf: leaf.filter(function (r) {
      return r.status === 'Active';
    }),
    byEv: groupBy_(
      lines.filter(function (l) {
        return l.status === 'Active' && l.event_id;
      }),
      function (l) {
        return l.event_id;
      }
    ),
    counts: { fields: fieldRows.length, 'spray events': events.length, 'spray lines': lines.length, 'plucking records': leaf.length }
  };
}

const checkNear_ = function (a, b, tol) {
  return Math.abs(a - b) <= tol;
};
const checkFirstDate_ = function (list) {
  return list
    .map(function (x) {
      return x.date;
    })
    .sort()[0];
};

// ---- 1. every sheet has the header row the app expects (a moved or renamed column would make every number wrong)
function checkHeaders_(c, rep) {
  Object.keys(SCHEMA).forEach(function (name) {
    const cols = SCHEMA[name],
      head = sheet_(name).getRange(1, 1, 1, cols.length).getValues()[0];
    cols.forEach(function (col, i) {
      if (String(head[i]) !== col)
        rep.bad('sheet header differs from the app (' + name + ')', 'column ' + (i + 1) + ' is "' + head[i] + '", the app expects "' + col + '"');
    });
  });
}

// ---- 2. the fields (the raw rows: the app clamps a too-large drone area when it reads a field, so only the sheet itself shows it)
function checkFields_(c, rep) {
  c.fieldRows.forEach(function (f) {
    if (!asBool_(f.active)) return;
    if (!(f.area_ha > 0)) rep.bad('active field without an area', f.field_id);
    else if (Number(f.drone_ha) > f.area_ha + 0.0005)
      rep.bad('field drone area above the field area', f.field_id + ' ' + f.drone_ha + ' > ' + f.area_ha);
  });
}

// ---- 3. ids are unique
function checkUniqueIds_(c, rep) {
  const dup = function (list, key, what) {
    const seen = {};
    list.forEach(function (r) {
      if (seen[r[key]]) rep.bad('duplicate ' + what, r[key]);
      seen[r[key]] = true;
    });
  };
  dup(c.events, 'event_id', 'event id');
  dup(c.lines, 'record_id', 'spray record id');
  dup(c.leaf, 'record_id', 'plucking record id');
}

// ---- 4. spray events
function checkSprayEvents_(c, rep) {
  const bad = rep.bad,
    warn = rep.warn;
  c.events.forEach(function (e) {
    const tag = 'event ' + e.event_id + ' (' + e.date + ')';
    if (e.status !== 'Active' && e.status !== 'Voided') bad('spray event with an unknown status', tag + ' "' + e.status + '"');
    if (e.status !== 'Active') return;
    if (!isDate_(e.date)) return bad('spray event with an invalid date', tag);
    if (e.date > c.today) bad('spray event dated in the future', tag);
    if (METHODS.indexOf(e.method) < 0) bad('spray event without a method', tag);
    if (!(e.kg_applied > 0)) bad('spray event without kg', tag);
    const mx = mixOf_(e);
    if (mx) {
      const compounds = round_(
        mx.reduce(function (a, m) {
          return a + m.kg;
        }, 0),
        3
      );
      if (!checkNear_(compounds, Number(e.kg_applied), 0.002))
        bad('spray event whose four compounds do not add up to its kg', tag + ' ' + compounds + ' vs ' + e.kg_applied);
      if (FOURTH_COMPOUNDS.indexOf(e.fourth_name) < 0)
        bad('spray event with an unknown fourth compound (Urea or T65)', tag + ' "' + e.fourth_name + '"');
      if (
        mx.some(function (m) {
          return m.kg < 0;
        })
      )
        bad('spray event with a negative kg of a compound', tag);
    } else if (e.epsom_kg !== '' || e.wuxal_kg !== '' || e.fourth_kg !== '' || e.fourth_name !== '')
      bad('spray event with only some of its compound columns filled', tag);
    if (!(e.total_ha > 0)) bad('spray event without hectares', tag);
    if (!(e.men >= 1)) bad('spray event without men or drones', tag);
    if (e.method === 'Drone' && Number(e.drones) !== Number(e.men))
      bad('drone event whose drones and men differ (the app stores one man per drone)', tag + ' men ' + e.men + ' drones ' + e.drones);
    if (e.method === 'Manual' && Number(e.drones) > 0) bad('manual event with drones', tag);
    if (isTime_(e.start_time) && isTime_(e.end_time)) {
      if (!checkNear_(Number(e.hours), (minutes_(e.end_time) - minutes_(e.start_time)) / 60, 0.002))
        bad('event hours differ from end - start', tag + ' ' + e.hours + ' h for ' + e.start_time + '-' + e.end_time);
    } else bad('spray event without a valid start and end time', tag);
    if (Number(e.hours) > 16) warn('spray event longer than 16 hours', tag + ' ' + e.hours + ' h');
    if (e.flags) warn('spray event flagged as implausible', tag + ' ' + e.flags);
  });
}

// ---- 5a. spray lines: each belongs to one active event and agrees with it
function checkSprayLines_(c, rep) {
  const bad = rep.bad,
    legacy = [];
  c.lines.forEach(function (l) {
    if (l.status !== 'Active') return;
    if (!l.event_id) {
      legacy.push(l);
      return;
    }
    const e = c.activeEv[l.event_id];
    if (!e) return bad('active spray line whose event is missing or voided', l.record_id + ' event ' + l.event_id);
    if (l.date !== e.date) bad('spray line on another date than its event', l.record_id);
    if (l.method !== e.method) bad('spray line of another method than its event', l.record_id);
    if (!c.fieldBy[l.field_id]) bad('spray line on an unknown field', l.record_id + ' ' + l.field_id);
    else if (partHa_(c.fieldBy[l.field_id], l.method) < MIN_LINE_HA)
      bad('spray line on a method the field does not have', l.record_id + ' ' + l.field_id + ' ' + l.method);
    if (!(l.area_ha > 0)) bad('spray line without hectares', l.record_id);
    if (Number(l.units) !== Number(e.men)) bad('spray line crew differs from its event', l.record_id + ' ' + l.units + ' vs ' + e.men);
    if (l.pct_total !== '' && (l.pct_total < 0 || l.pct_total > 100.0001))
      bad('spray line with a coverage outside 0-100%', l.record_id + ' ' + l.pct_total);
    if (!!mixOf_(l) !== !!mixOf_(e)) bad('spray line and its event differ in whether the four compounds were recorded', l.record_id);
    else if (mixOf_(e) && l.fourth_name !== e.fourth_name)
      bad('spray line with another fourth compound than its event', l.record_id + ' ' + l.fourth_name + ' vs ' + e.fourth_name);
  });
  if (legacy.length)
    rep.warn(
      'active spray lines with no event (older entries: the dashboard does not count them)',
      legacy.length + ' lines, first ' + legacy[0].record_id
    );
}

// ---- 5b. the lines of an event add up to the event
function checkEventTotals_(c, rep) {
  const bad = rep.bad;
  Object.keys(c.activeEv).forEach(function (id) {
    const e = c.activeEv[id],
      ls = c.byEv[id] || [],
      n = ls.length,
      tol = 0.0005 * n + 0.002;
    if (!n) return bad('active spray event without any field line', 'event ' + id + ' (' + e.date + ')');
    const sum = function (k) {
      return ls.reduce(function (a, l) {
        return a + (Number(l[k]) || 0);
      }, 0);
    };
    if (!checkNear_(sum('area_ha'), Number(e.total_ha), tol))
      bad('fields of an event do not add up to its hectares', id + ' lines ' + round_(sum('area_ha'), 3) + ' vs ' + e.total_ha);
    if (!checkNear_(sum('kg_applied'), Number(e.kg_applied), tol))
      bad('fields of an event do not add up to its kg', id + ' lines ' + round_(sum('kg_applied'), 3) + ' vs ' + e.kg_applied);
    if (!checkNear_(sum('hours'), Number(e.hours), tol))
      bad('fields of an event do not add up to its hours', id + ' lines ' + round_(sum('hours'), 3) + ' vs ' + e.hours);
    if (mixOf_(e))
      [
        ['zinc_kg', 'Zinc'],
        ['epsom_kg', 'Epsom salt'],
        ['wuxal_kg', 'Wuxal'],
        ['fourth_kg', 'the fourth compound']
      ].forEach(function (k) {
        if (!checkNear_(sum(k[0]), Number(e[k[0]]), tol))
          bad('fields of an event do not add up to its kg of ' + k[1], id + ' lines ' + round_(sum(k[0]), 3) + ' vs ' + e[k[0]]);
      });
  });
}

// ---- 6. rounds: no round covers more than the field part (spray and plucking)
function checkRounds_(c, rep) {
  const roundCheck = function (what, recs, resetDays) {
    const by = groupBy_(recs, function (r) {
      return r.field_id;
    });
    Object.keys(by).forEach(function (fid) {
      const f = c.fieldBy[fid];
      if (!f) return;
      partMethods_(f).forEach(function (m) {
        const ext = partExt_(f, m);
        if (!(ext > 0)) return;
        buildRounds_(partRecs_(f, m, by[fid]), ext, resetDays).forEach(function (rd) {
          if (rd.done_ha > ext + ROUND_TOL)
            rep.bad(
              what + ' round covering more than the field part',
              fieldLabel_(f) + ' ' + m + ' round from ' + rd.start + ': ' + rd.done_ha + ' ha of ' + ext
            );
        });
      });
    });
  };
  roundCheck('spray', c.activeLines, c.params.spray_round_reset_days);
  roundCheck('plucking', c.activeLeaf, c.params.pluck_round_reset_days);
}

// ---- 7. plucking records (and: no two for one part on one day)
function checkPluckingRecords_(c, rep) {
  const bad = rep.bad,
    seenLeaf = {};
  c.activeLeaf.forEach(function (r) {
    const tag = r.record_id + ' (' + r.date + ' ' + r.field_id + ')';
    if (!isDate_(r.date)) return bad('plucking record with an invalid date', tag);
    if (r.date > c.today) bad('plucking record dated in the future', tag);
    const f = c.fieldBy[r.field_id];
    if (!f) return bad('plucking record on an unknown field', tag);
    if (METHODS.indexOf(r.method) < 0) bad('plucking record without a method', tag);
    else if (partHa_(f, r.method) < MIN_LINE_HA) bad('plucking record on a method the field does not have', tag + ' ' + r.method);
    if (!(r.kg > 0)) bad('plucking record without kg', tag);
    if (!(r.area_ha > 0)) bad('plucking record without hectares', tag);
    else if (METHODS.indexOf(r.method) >= 0 && r.area_ha > partHa_(f, r.method) + ROUND_TOL)
      bad('plucking record above the area of its field part', tag + ' ' + r.area_ha + ' ha of ' + partHa_(f, r.method));
    const key = r.date + '|' + r.field_id + '|' + r.method;
    if (seenLeaf[key]) bad('two active plucking records for one field part on one day', tag);
    seenLeaf[key] = true;
    if (r.flags) rep.warn('plucking record flagged as implausible', tag + ' ' + r.flags);
  });
}

// ---- 8. can the dashboard price it? (a record dated before the first rate gets cost 0 - see checkCostRates)
function checkPricing_(c, rep) {
  const rateAt = costRates_(),
    act = Object.keys(c.activeEv).map(function (k) {
      return c.activeEv[k];
    });
  const none = function (list, f, what) {
    const n = list.filter(function (x) {
      return isDate_(x.date) && !(f(rateAt(x.date)) > 0);
    });
    if (n.length) rep.warn(what, n.length + ' of ' + list.length + ', first ' + checkFirstDate_(n));
  };
  none(
    act.filter(function (e) {
      return e.method === 'Drone';
    }),
    function (r) {
      return r.basis === 'per_ha' ? r.drone_ha : r.drone_hr;
    },
    'drone events on a day with no drone charge in force (their labour counts as 0)'
  );
  none(
    act.filter(function (e) {
      return e.method === 'Manual';
    }),
    function (r) {
      return r.worker;
    },
    'manual events on a day with no man-hour rate in force (their labour counts as 0)'
  );
  none(
    act.filter(function (e) {
      return !mixOf_(e);
    }),
    function (r) {
      return r.price;
    },
    'older spray events (one total kg) on a day with no foliar price in force (their foliar cost counts as 0)'
  );
  const noPrice = act.filter(function (e) {
    return mixOf_(e) && foliarUnpriced_(e, rateAt(e.date));
  });
  if (noPrice.length)
    rep.warn(
      'spray events with a compound that has no price in force on that day (that compound counts as 0 in the foliar cost)',
      noPrice.length + ' of ' + act.length + ', first ' + checkFirstDate_(noPrice)
    );
  const noPlan = act.filter(function (e) {
    return mixOf_(e) && !e.plan_id;
  });
  if (noPlan.length)
    rep.warn('spray events recorded with no foliar plan in force on their day', noPlan.length + ', first ' + checkFirstDate_(noPlan));
  none(
    c.activeLeaf,
    function (r) {
      return r.gl;
    },
    'plucking records on a day with no green leaf price in force (their revenue counts as 0)'
  );
}

// ---- 9. the foliar schedule (plan file template and the stamps of the months)
function checkSchedule_(c, rep) {
  const bad = rep.bad,
    warn = rep.warn,
    tplRows = templateAll_(),
    stm = dbAll_(SHEETS.STAMPS).filter(function (s) {
      return s.status === 'Active';
    });
  c.counts['plan blocks'] = tplRows.length;
  c.counts['calendar stamps'] = stm.length;
  tplRows.forEach(function (t) {
    if (!c.fieldBy[t.field_id]) bad('plan block of a field that does not exist', t.slot_id);
    else if (!(partHa_(c.fieldBy[t.field_id], 'Drone') > 0))
      warn('plan block of a field that has no drone area now', fieldLabel_(c.fieldBy[t.field_id]));
  });
  const perPart = {},
    perBlock = {},
    perDay = {};
  stm.forEach(function (s) {
    const f = c.fieldBy[s.field_id];
    if (!f) return bad('calendar stamp of a field that does not exist', s.stamp_id);
    if (!isDate_(s.date) || monthKey_(s.date) !== s.month)
      bad('calendar stamp whose date is not in its month', s.stamp_id + ' ' + s.date + ' / ' + s.month);
    if (HALVES.indexOf(s.half) < 0) bad('calendar stamp that is neither morning nor evening', s.stamp_id);
    if (!(s.ha > 0)) bad('calendar stamp without hectares', s.stamp_id);
    const k = s.month + '|' + s.field_id;
    perPart[k] = (perPart[k] || 0) + s.ha;
    perBlock[s.date + ' ' + s.half] = (perBlock[s.date + ' ' + s.half] || 0) + s.ha;
    perDay[s.date] = (perDay[s.date] || 0) + s.ha;
  });
  Object.keys(perPart).forEach(function (k) {
    const f = c.fieldBy[k.split('|')[1]];
    if (f && Math.abs(perPart[k] - partHa_(f, 'Drone')) > 0.006)
      warn(
        'the stamps of a field in a month do not add up to its drone area (the plan is out of date: load the plan file again)',
        fieldLabel_(f) + ' ' + k.split('|')[0] + ': stamps ' + round_(perPart[k], 2) + ' ha, drone area ' + partHa_(f, 'Drone') + ' ha'
      );
  });
  Object.keys(perBlock).forEach(function (k) {
    if (perBlock[k] > c.params.foliar_block_max_ha + 1e-9)
      warn('a morning or evening block holds more ha than the limit', k + ' ' + round_(perBlock[k], 2) + ' ha');
  });
  Object.keys(perDay).forEach(function (k) {
    if (perDay[k] > c.params.foliar_day_max_ha + 1e-9) warn('a day holds more ha than the limit', k + ' ' + round_(perDay[k], 2) + ' ha');
  });
}

/** The checks, in the order they are reported. Each is (context, reporter) and can be run alone on a damaged sheet. */
const DATA_CHECKS = [
  checkHeaders_,
  checkFields_,
  checkUniqueIds_,
  checkSprayEvents_,
  checkSprayLines_,
  checkEventTotals_,
  checkRounds_,
  checkPluckingRecords_,
  checkPricing_,
  checkSchedule_
];

function checkData_() {
  return withRequest_(function () {
    const rep = makeReporter_(),
      ctx = checkContext_();
    DATA_CHECKS.forEach(function (check) {
      check(ctx, rep);
    });
    return { problems: rep.problems(), warnings: rep.warnings(), counts: ctx.counts };
  });
}

// ---- Estate.gs
/** Estate structure: one estate -> divisions -> fields. The admin enters counts; fields are generated, then filled in.
 *  A field code is unique within its division (every division can have a "1VP"); the field_id is the estate-wide key. */
function fieldLabel_(f) {
  return (f.division ? f.division + ' ' : '') + f.code;
}

function fieldComplete_(f) {
  return !!(f.active && f.area_ha > 0 && f.last_prune_date && f.tea_type && f.method);
}

/** Work out a field's drone / manual split. drone_ha is the part of the field sprayed by drone, the rest is sprayed by hand.
 *  A field whose drone_ha was never entered but has the older single "method" keeps meaning: Drone = all drone, Manual = none.
 *  Sets f.drone_ha, f.manual_ha (both '' until known) and f.method = Drone / Manual / Mixed ('' until known). */
function normField_(f) {
  const area = f.area_ha > 0 ? f.area_ha : 0;
  let d = f.drone_ha;
  if (d === '' || d === undefined || d === null) d = f.method === 'Drone' ? area || '' : f.method === 'Manual' ? 0 : '';
  if (d === '' || !area) {
    f.drone_ha = d;
    f.manual_ha = '';
    if (area || (f.method !== 'Drone' && f.method !== 'Manual')) f.method = '';
    return f;
  } // no area yet: an older single method is kept as a hint
  d = Math.max(0, Math.min(d, area));
  if (d < ROUND_TOL) d = 0;
  else if (area - d < ROUND_TOL) d = area; // a sliver counts as none / all (also when a field is saved: cleanFieldRow_)
  f.drone_ha = round_(d, 3);
  f.manual_ha = round_(area - d, 3);
  f.method = d === 0 ? 'Manual' : d === area ? 'Drone' : 'Mixed';
  return f;
}
/** All fields, with the drone / manual split worked out. */
function fieldsAll_() {
  return dbAll_(SHEETS.FIELDS).map(normField_);
}
/** The part of a field (ha) that is sprayed with one method: 'Drone' or 'Manual'. */
function partHa_(f, method) {
  const v = method === 'Drone' ? f.drone_ha : f.manual_ha;
  return v > 0 ? v : 0;
}
/** The whole divisions assigned to a user account (a comma list in Users.division_ids). */
function userDivisions_(u) {
  return String((u && u.division_ids) || '')
    .split(',')
    .map(function (s) {
      return s.trim();
    })
    .filter(Boolean);
}
/** May this account record on / see this field? (Server-side rule; the app also hides other fields.) */
function inScope_(c, f) {
  if (c.isAdmin || c.role.field_scope === 'all') return true;
  if (c.role.field_scope === 'drone') return partHa_(f, 'Drone') > 0;
  return f.afo_user_id === c.user.user_id || userDivisions_(c.user).indexOf(f.division_id) >= 0; // the fields ticked for this account, or every field of a division assigned to it
}
function padNum_(n, w) {
  let s = String(n);
  while (s.length < w) s = '0' + s;
  return s;
}

/** Fields indexed by id, with division name and AFO name joined. */
function fieldIndex_() {
  const divs = {};
  dbAll_(SHEETS.DIVISIONS).forEach(function (d) {
    divs[d.division_id] = d;
  });
  const users = {};
  dbAll_(SHEETS.USERS).forEach(function (u) {
    users[u.user_id] = u;
  });
  const byId = {};
  fieldsAll_().forEach(function (f) {
    const d = divs[f.division_id] || {};
    f.division = d.name || '';
    f.division_seq = d.seq || 0;
    const u = users[f.afo_user_id];
    f.afo_name = u ? u.display_name : '';
    f.afo_active = u ? !!u.active : false;
    byId[f.field_id] = f;
  });
  return byId;
}

function api_getEstate(token) {
  return wrap_(function () {
    authz_(token, ['estate.manage', 'users.manage']);
    const estate = dbAll_(SHEETS.ESTATE)[0] || null;
    const idx = fieldIndex_();
    const fields = Object.keys(idx)
      .map(function (k) {
        return idx[k];
      })
      .sort(function (a, b) {
        return a.division_seq === b.division_seq ? a.seq - b.seq : a.division_seq - b.division_seq;
      });
    const divisions = dbAll_(SHEETS.DIVISIONS)
      .sort(function (a, b) {
        return a.seq - b.seq;
      })
      .map(function (d) {
        const mine = fields.filter(function (f) {
          return f.division_id === d.division_id;
        });
        return {
          division_id: d.division_id,
          name: d.name,
          seq: d.seq,
          active: d.active,
          field_count: mine.length,
          complete_count: mine.filter(fieldComplete_).length,
          assigned_count: mine.filter(function (f) {
            return !!f.afo_user_id;
          }).length
        };
      });
    return {
      estate: estate ? { estate_id: estate.estate_id, name: estate.name } : null,
      divisions: divisions,
      fields: fields.map(function (f) {
        return {
          field_id: f.field_id,
          division_id: f.division_id,
          division: f.division,
          seq: f.seq,
          code: f.code,
          name: f.name,
          area_ha: f.area_ha,
          last_prune_date: f.last_prune_date,
          tea_type: f.tea_type,
          method: f.method,
          drone_ha: f.drone_ha,
          manual_ha: f.manual_ha,
          afo_user_id: f.afo_user_id,
          afo_name: f.afo_name,
          afo_active: f.afo_active,
          active: f.active,
          complete: fieldComplete_(f)
        };
      })
    };
  });
}

function api_createEstate(token, p) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    p = p || {};
    const name = cleanText_(p.name, 80, 'Estate name', true);
    const n = cleanNum_(p.division_count, 'Number of divisions', { required: true, min: 1, max: 30, int: true });
    return withLock_(function () {
      if (dbAll_(SHEETS.ESTATE).length) fail_('The estate already exists.', 'VALIDATION');
      const now = nowIso_();
      dbAppend_(SHEETS.ESTATE, [{ estate_id: newId_('EST'), name: name, created_at: now }]);
      const divs = [];
      for (let i = 1; i <= n; i++) divs.push({ division_id: newId_('DIV'), name: 'Division ' + i, seq: i, active: true, created_at: now });
      dbAppend_(SHEETS.DIVISIONS, divs);
      audit_(c.user.username, 'ESTATE_CREATE', 'estate', '', { name: name, divisions: n });
      return true;
    });
  });
}

function api_renameEstate(token, name) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    const nm = cleanText_(name, 80, 'Estate name', true);
    return withLock_(function () {
      const e = dbAll_(SHEETS.ESTATE)[0];
      if (!e) fail_('Create the estate first.', 'VALIDATION');
      dbUpdate_(SHEETS.ESTATE, e._row, { name: nm });
      audit_(c.user.username, 'ESTATE_RENAME', 'estate', e.estate_id, { name: nm });
      return true;
    });
  });
}

function api_addDivisions(token, count) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    const n = cleanNum_(count, 'Number of divisions', { required: true, min: 1, max: 30, int: true });
    return withLock_(function () {
      if (!dbAll_(SHEETS.ESTATE).length) fail_('Create the estate first.', 'VALIDATION');
      const divs = dbAll_(SHEETS.DIVISIONS);
      if (divs.length + n > 60) fail_('Too many divisions.', 'VALIDATION');
      let seq = divs.reduce(function (m, d) {
        return Math.max(m, d.seq);
      }, 0);
      const add = [];
      for (let i = 0; i < n; i++) {
        seq++;
        add.push({ division_id: newId_('DIV'), name: 'Division ' + seq, seq: seq, active: true, created_at: nowIso_() });
      }
      dbAppend_(SHEETS.DIVISIONS, add);
      audit_(c.user.username, 'DIVISIONS_ADD', 'division', '', { added: n });
      return n;
    });
  });
}

function api_saveDivision(token, d) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    d = d || {};
    const name = cleanText_(d.name, 60, 'Division name', true);
    return withLock_(function () {
      const rows = dbAll_(SHEETS.DIVISIONS);
      const cur = rows.filter(function (r) {
        return r.division_id === d.division_id;
      })[0];
      if (!cur) fail_('Division not found.', 'VALIDATION');
      if (
        rows.some(function (r) {
          return r.name.toLowerCase() === name.toLowerCase() && r.division_id !== cur.division_id;
        })
      )
        fail_('Another division has this name.', 'VALIDATION');
      const active = d.active === undefined ? cur.active !== false : d.active !== false; // a request that does not say keeps the division as it is
      dbUpdate_(SHEETS.DIVISIONS, cur._row, { name: name, active: active });
      audit_(c.user.username, 'DIVISION_UPDATE', 'division', cur.division_id, { name: name, active: active });
      return true;
    });
  });
}

/** Generate N blank fields in a division. They are filled in afterwards (area, last prune date, VP/SD, drone area). */
function api_addFields(token, divisionId, count) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    const n = cleanNum_(count, 'Number of fields', { required: true, min: 1, max: 300, int: true });
    return withLock_(function () {
      const div = dbAll_(SHEETS.DIVISIONS).filter(function (d) {
        return d.division_id === divisionId;
      })[0];
      if (!div) fail_('Division not found.', 'VALIDATION');
      const all = dbAll_(SHEETS.FIELDS);
      if (all.length + n > 2000) fail_('Too many fields.', 'VALIDATION');
      const codes = {};
      all.forEach(function (f) {
        if (f.division_id === divisionId) codes[f.code.toLowerCase()] = true;
      });
      let seq = all
        .filter(function (f) {
          return f.division_id === divisionId;
        })
        .reduce(function (m, f) {
          return Math.max(m, f.seq);
        }, 0);
      const add = [],
        now = nowIso_();
      for (let i = 0; i < n; i++) {
        let code;
        do {
          seq++;
          code = 'D' + div.seq + '-' + padNum_(seq, 2);
        } while (codes[code.toLowerCase()]);
        codes[code.toLowerCase()] = true;
        add.push({
          field_id: newId_('FLD'),
          division_id: divisionId,
          seq: seq,
          code: code,
          name: '',
          area_ha: '',
          last_prune_date: '',
          tea_type: '',
          method: '',
          drone_ha: '',
          afo_user_id: '',
          active: true,
          created_at: now,
          updated_at: now
        });
      }
      dbAppend_(SHEETS.FIELDS, add);
      audit_(c.user.username, 'FIELDS_ADD', 'field', divisionId, { division: div.name, added: n });
      return n;
    });
  });
}

/** Add ONE field with its details (code, name, area, drone area; the rest of the area is sprayed by hand). A blank drone area means all by hand. */
function api_addField(token, divisionId, r) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    r = r || {};
    return withLock_(function () {
      const div = dbAll_(SHEETS.DIVISIONS).filter(function (d) {
        return d.division_id === divisionId;
      })[0];
      if (!div) fail_('Division not found.', 'VALIDATION');
      const area = cleanNum_(r.area_ha, 'Area (ha)', { required: true, gt: 0, max: 1000 });
      const cl = cleanFieldRow_(
        Object.assign({}, r, { area_ha: area, drone_ha: r.drone_ha === '' || r.drone_ha === undefined || r.drone_ha === null ? 0 : r.drone_ha })
      );
      const all = dbAll_(SHEETS.FIELDS);
      if (all.length >= 2000) fail_('Too many fields.', 'VALIDATION');
      if (
        all.some(function (f) {
          return f.division_id === divisionId && String(f.code).toLowerCase() === cl.code.toLowerCase();
        })
      )
        fail_('Code "' + cl.code + '" is already used in this division.', 'VALIDATION');
      const seq =
          all
            .filter(function (f) {
              return f.division_id === divisionId;
            })
            .reduce(function (m, f) {
              return Math.max(m, f.seq);
            }, 0) + 1,
        now = nowIso_(),
        id = newId_('FLD');
      dbAppend_(SHEETS.FIELDS, [
        {
          field_id: id,
          division_id: divisionId,
          seq: seq,
          code: cl.code,
          name: cl.name,
          area_ha: cl.area_ha,
          last_prune_date: cl.last_prune_date,
          tea_type: cl.tea_type,
          method: '',
          afo_user_id: '',
          drone_ha: cl.drone_ha,
          active: true,
          created_at: now,
          updated_at: now
        }
      ]);
      audit_(c.user.username, 'FIELD_ADD', 'field', id, { division: div.name, code: cl.code, area_ha: cl.area_ha, drone_ha: cl.drone_ha });
      return { field_id: id };
    });
  });
}

/** A field leaves the drone plan (deleted or made inactive): its template blocks and its stamps go too. Ticking Active again does not bring them back: load the plan file again. */
function retirePlanBlocks_(fieldId, status) {
  [SHEETS.TEMPLATE, SHEETS.STAMPS].forEach(function (name) {
    dbUpdateMany_(
      name,
      dbAll_(name)
        .filter(function (x) {
          return x.field_id === fieldId && x.status === 'Active';
        })
        .map(function (x) {
          return { row: x._row, patch: { status: status } };
        })
    );
  });
}

/** Remove a field. A field that never had any spray or plucking record is deleted for good (with its drone plan blocks); a field that has records is only made inactive: it disappears from the entry
 *  screens and the land, and its records stay (ticking Active again brings the field back, not its plan blocks). */
function api_removeField(token, fieldId) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    return withLock_(function () {
      const f = dbAll_(SHEETS.FIELDS).filter(function (x) {
        return x.field_id === String(fieldId || '');
      })[0];
      if (!f) fail_('Field not found.', 'VALIDATION');
      const used =
        dbAll_(SHEETS.SPRAY).some(function (l) {
          return l.field_id === f.field_id;
        }) ||
        dbAll_(SHEETS.LEAF).some(function (l) {
          return l.field_id === f.field_id;
        });
      retirePlanBlocks_(f.field_id, used ? 'Voided' : 'Removed');
      if (used) {
        dbUpdate_(SHEETS.FIELDS, f._row, { active: false, updated_at: nowIso_() });
        audit_(c.user.username, 'FIELD_REMOVE', 'field', f.field_id, { code: f.code, deleted: false });
        return { deleted: false };
      }
      dbDeleteRow_(SHEETS.FIELDS, f._row);
      audit_(c.user.username, 'FIELD_DELETE', 'field', f.field_id, { code: f.code, deleted: true });
      return { deleted: true };
    });
  });
}

function cleanFieldRow_(r) {
  const code = cleanText_(r.code, 20, 'Field code', true);
  if (!/^[A-Za-z0-9._\- ]+$/.test(code)) fail_('Field code may use letters, digits, space, dot, dash and underscore only.', 'VALIDATION');
  const prune = String(r.last_prune_date || '').trim();
  if (prune && !isDate_(prune)) fail_('Last prune date must be a valid date.', 'VALIDATION');
  if (prune && prune > todayStr_()) fail_('Last prune date cannot be in the future.', 'VALIDATION');
  const tea = String(r.tea_type || '');
  if (tea && TEA_TYPES.indexOf(tea) < 0) fail_('Tea type must be VP or SD.', 'VALIDATION');
  const legacy = String(r.method || '');
  if (legacy && METHODS.indexOf(legacy) < 0) fail_('Spray method must be Drone or Manual.', 'VALIDATION');
  const area = round_(cleanNum_(r.area_ha, 'Area (ha)', { gt: 0, max: 1000 }), 3);
  let drone = cleanNum_(r.drone_ha, 'Drone area (ha)', { min: 0, max: 1000 });
  if (drone !== '') drone = round_(drone, 3);
  if (drone === '' && legacy && area !== '') drone = legacy === 'Drone' ? area : 0; // older form: one method for the whole field
  if (drone !== '' && area === '') fail_('Enter the field area before the drone area.', 'VALIDATION');
  if (drone !== '' && drone > area + AREA_TOL)
    fail_('The drone area (' + drone + ' ha) is more than the field area (' + area + ' ha).', 'VALIDATION');
  if (drone !== '') {
    drone = Math.min(drone, area);
    if (drone < ROUND_TOL) drone = 0;
    else if (area - drone < ROUND_TOL) drone = area;
  } // a sliver counts as none / all (normField_ does the same when reading)
  const cl = {
    code: code,
    name: cleanText_(r.name, 80, 'Field name', false),
    area_ha: area,
    last_prune_date: prune,
    tea_type: tea,
    drone_ha: drone,
    active: r.active !== false
  };
  cl.method = normField_({ area_ha: area, drone_ha: drone, method: '' }).method;
  return cl;
}

/** Save all edited rows of a division in one go (all or nothing). */
function api_saveFields(token, divisionId, rows) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    rows = Array.isArray(rows) ? rows : [];
    if (!rows.length) fail_('Nothing to save.', 'VALIDATION');
    if (rows.length > 500) fail_('Too many rows.', 'VALIDATION');
    return withLock_(function () {
      const all = fieldsAll_();
      const byId = {};
      all.forEach(function (f) {
        byId[f.field_id] = f;
      });
      const clean = [],
        seen = {},
        inList = {};
      rows.forEach(function (r) {
        inList[r.field_id] = true;
      });
      rows.forEach(function (r, i) {
        const cur = byId[r.field_id];
        if (!cur || cur.division_id !== divisionId) fail_('Row ' + (i + 1) + ': field not found in this division.', 'VALIDATION');
        let cl;
        try {
          cl = cleanFieldRow_(r);
        } catch (e) {
          if (e instanceof AppError) throw new AppError('Field ' + (cur.code || i + 1) + ': ' + e.message, e.code);
          throw e;
        }
        const key = cl.code.toLowerCase();
        if (seen[key]) fail_('Code "' + cl.code + '" is used twice in this list.', 'VALIDATION');
        seen[key] = true;
        if (
          all.some(function (f) {
            return f.division_id === divisionId && f.field_id !== cur.field_id && f.code.toLowerCase() === key && !inList[f.field_id];
          })
        ) {
          fail_('Code "' + cl.code + '" is already used in this division.', 'VALIDATION');
        }
        clean.push({ cur: cur, cl: cl });
      });
      let changed = 0;
      const methodChanges = [];
      clean.forEach(function (x) {
        const a = x.cur,
          b = x.cl;
        if (
          a.code === b.code &&
          a.name === b.name &&
          a.area_ha === (b.area_ha === '' ? '' : b.area_ha) &&
          a.last_prune_date === b.last_prune_date &&
          a.tea_type === b.tea_type &&
          a.method === b.method &&
          a.drone_ha === b.drone_ha &&
          a.active === b.active
        )
          return;
        if (a.drone_ha !== b.drone_ha) methodChanges.push(b.code + ': drone ' + a.drone_ha + '->' + b.drone_ha + ' ha');
        b.updated_at = nowIso_();
        dbUpdate_(SHEETS.FIELDS, a._row, b);
        if (a.active && !b.active) retirePlanBlocks_(a.field_id, 'Voided'); // the same as removing it
        changed++;
      });
      audit_(c.user.username, 'FIELDS_SAVE', 'field', divisionId, {
        rows: rows.length,
        changed: changed,
        method_changes: methodChanges.slice(0, 20)
      });
      return { changed: changed };
    });
  });
}

// ---- FoliarPlan.gs
/** The FOLIAR PLAN: what is mixed for ONE hectare.
 *
 *  The foliar is four compounds mixed with water: Zinc, Epsom salt, Wuxal and (depending on the month) Urea or T65. Once a month it is decided how many kg of each
 *  compound go on one hectare. A plan row holds those four kg per ha and the date from which it applies; the plan in force on a day is the latest one whose date is on or
 *  before that day. The spray entry screen multiplies the plan by the hectares sprayed to suggest the kg of each compound; the person can change the kg (what was really
 *  mixed is what is stored) and an event that is far from the plan is flagged.
 *  Plans are never deleted: a plan entered for the same date replaces the earlier one, and a wrong plan can be voided (the audit trail keeps both).
 */

/** The active plans, oldest first (the later row wins when two have the same date). */
function plansAll_() {
  return dbAll_(SHEETS.PLANS)
    .filter(function (p) {
      return p.status === 'Active' && isDate_(p.effective_from);
    })
    .sort(function (a, b) {
      return a.effective_from === b.effective_from ? a._row - b._row : a.effective_from < b.effective_from ? -1 : 1;
    });
}
/** The plan in force on a date (from plansAll_()), or null. */
function planOn_(date, plans) {
  let best = null;
  (plans || plansAll_()).forEach(function (p) {
    if (p.effective_from <= date) best = p;
  });
  return best;
}
function shapePlan_(p) {
  return {
    plan_id: p.plan_id,
    effective_from: p.effective_from,
    zinc: Number(p.zinc_kg_ha) || 0,
    epsom: Number(p.epsom_kg_ha) || 0,
    wuxal: Number(p.wuxal_kg_ha) || 0,
    fourth_name: p.fourth_name,
    fourth: Number(p.fourth_kg_ha) || 0,
    note: p.note,
    set_by: p.set_by,
    set_at: p.set_at
  };
}

/** The Foliar planning screen: the plan in force today and the whole history. */
function api_listPlans(token) {
  return wrap_(function () {
    authz_(token, ['foliar.plan']);
    const today = todayStr_(),
      all = dbAll_(SHEETS.PLANS),
      act = plansAll_(),
      cur = planOn_(today, act);
    const next = act.filter(function (p) {
      return p.effective_from > today;
    })[0];
    return {
      today: today,
      fourth_options: FOURTH_COMPOUNDS.slice(),
      current: cur ? shapePlan_(cur) : null,
      next: next ? shapePlan_(next) : null,
      history: all
        .slice()
        .reverse()
        .map(function (p) {
          const o = shapePlan_(p);
          o.status = p.status;
          o.void_reason = p.void_reason;
          o.voided_by = p.voided_by;
          o.in_force = !!cur && cur.plan_id === p.plan_id;
          return o;
        })
    };
  });
}

function api_savePlan(token, p) {
  return wrap_(function () {
    const c = authz_(token, ['foliar.plan']);
    p = p || {};
    const eff = String(p.effective_from || todayStr_());
    if (!isDate_(eff)) fail_('Choose the date from which this plan applies.', 'VALIDATION');
    if (eff > addDays_(todayStr_(), 366)) fail_('The date is more than a year ahead. Check it.', 'VALIDATION');
    const vals = {};
    MIX_FIXED.forEach(function (m) {
      vals[m.key] = cleanNum_(p[m.key], m.name + ' (kg per ha)', { required: true, min: 0, max: 10000 });
    });
    const fourthName = String(p.fourth_name || '');
    if (FOURTH_COMPOUNDS.indexOf(fourthName) < 0) fail_('Choose the fourth compound: ' + FOURTH_COMPOUNDS.join(' or ') + '.', 'VALIDATION');
    const fourth = cleanNum_(p.fourth, fourthName + ' (kg per ha)', { required: true, min: 0, max: 10000 });
    if (!(vals.zinc + vals.epsom + vals.wuxal + fourth > 0)) fail_('Enter the kg per hectare of at least one compound.', 'VALIDATION');
    const note = cleanText_(p.note, 200, 'Note', false);
    return withLock_(function () {
      const now = nowIso_(),
        replaced = [];
      dbAll_(SHEETS.PLANS).forEach(function (old) {
        if (old.status === 'Active' && old.effective_from === eff) {
          // one plan per date: the new one replaces the old one
          dbUpdate_(SHEETS.PLANS, old._row, {
            status: 'Voided',
            void_reason: 'Replaced by a new plan for the same date',
            voided_by: c.user.username,
            voided_at: now
          });
          replaced.push(old.plan_id);
        }
      });
      const id = newId_('PLN');
      dbAppend_(SHEETS.PLANS, [
        {
          plan_id: id,
          effective_from: eff,
          zinc_kg_ha: vals.zinc,
          epsom_kg_ha: vals.epsom,
          wuxal_kg_ha: vals.wuxal,
          fourth_name: fourthName,
          fourth_kg_ha: fourth,
          note: note,
          status: 'Active',
          set_by: c.user.username,
          set_at: now
        }
      ]);
      audit_(c.user.username, 'FOLIAR_PLAN_SAVE', 'foliar_plan', id, {
        effective_from: eff,
        zinc: vals.zinc,
        epsom: vals.epsom,
        wuxal: vals.wuxal,
        fourth_name: fourthName,
        fourth: fourth,
        replaced: replaced
      });
      return { plan_id: id, replaced: replaced.length };
    });
  });
}

function api_voidPlan(token, planId, reason) {
  return wrap_(function () {
    const c = authz_(token, ['foliar.plan']);
    const why = cleanText_(reason, 200, 'Reason', false) || 'Voided by ' + c.user.username;
    return withLock_(function () {
      const p = dbAll_(SHEETS.PLANS).filter(function (x) {
        return x.plan_id === String(planId || '');
      })[0];
      if (!p || p.status !== 'Active') fail_('Plan not found.', 'VALIDATION');
      dbUpdate_(SHEETS.PLANS, p._row, { status: 'Voided', void_reason: why, voided_by: c.user.username, voided_at: nowIso_() });
      audit_(c.user.username, 'FOLIAR_PLAN_VOID', 'foliar_plan', p.plan_id, { effective_from: p.effective_from, reason: why });
      return true;
    });
  });
}

/** The four compounds of a spray event or line as a list [{name, kg}], or null for an older event that recorded one total kg. */
function mixOf_(r) {
  if (r.zinc_kg === '' || r.zinc_kg === undefined || r.zinc_kg === null) return null;
  return [
    { name: 'Zinc', kg: Number(r.zinc_kg) || 0 },
    { name: 'Epsom salt', kg: Number(r.epsom_kg) || 0 },
    { name: 'Wuxal', kg: Number(r.wuxal_kg) || 0 },
    { name: r.fourth_name === 'T65' ? 'T65' : 'Urea', kg: Number(r.fourth_kg) || 0 }
  ];
}

/** What to check an event's compounds against: flags (text) for every compound whose kg is further than the tolerance from plan kg/ha x ha. */
function planFlags_(plan, mix, totalHa, tolPct) {
  if (!plan) return ['NO FOLIAR PLAN FOR THIS DATE'];
  const out = [],
    tol = (Number(tolPct) || 0) / 100;
  const check = function (name, actual, perHa) {
    const want = perHa * totalHa;
    if (!(want > 0)) {
      if (actual > 0) out.push(name.toUpperCase() + ' NOT IN PLAN');
      return;
    }
    const dev = (actual - want) / want;
    if (Math.abs(dev) > tol + 1e-9) out.push(name.toUpperCase() + ' ' + (dev > 0 ? '+' : '') + Math.round(dev * 100) + '% VS PLAN');
  };
  check('Zinc', mix.zinc, Number(plan.zinc_kg_ha) || 0);
  check('Epsom salt', mix.epsom, Number(plan.epsom_kg_ha) || 0);
  check('Wuxal', mix.wuxal, Number(plan.wuxal_kg_ha) || 0);
  if (mix.fourth_name !== plan.fourth_name) {
    if (mix.fourth > 0 || Number(plan.fourth_kg_ha) > 0)
      out.push(mix.fourth_name.toUpperCase() + ' INSTEAD OF ' + String(plan.fourth_name).toUpperCase());
  } // (only the name in the list was changed, no kg either side: nothing to flag)
  else check(mix.fourth_name, mix.fourth, Number(plan.fourth_kg_ha) || 0);
  return out;
}

// ---- FoliarSchedule.gs
/** THE MONTHLY DRONE SPRAY SCHEDULE (the plan file) and its calendar of stamps.
 *
 *  The plan file (FoliarPlanData.gs, from "Foliar Plan.xlsx") says, for every field that is sprayed by drone, on which weekday, in which week of the month and in which half of the day
 *  (morning AM / evening PM) how many ha of its drone part are sprayed. It is the same every month. It is kept in the sheet FoliarTemplate.
 *  When a month is opened for the first time it is copied into the sheet FoliarStamps: one STAMP per planned block of one field's drone part. A stamp has a date and a half-day.
 *  "Week n" means the n-th occurrence of that weekday in the month (the 1st Monday, the 2nd Monday ...), so week 1 is days 1-7 of the month, week 2 is days 8-14 and so on.
 *
 *  The colour of a stamp is not stored: it is worked out from the spraying that was really recorded. The drone ha sprayed on a field part in the month are shared over that part's stamps in
 *  the order of their dates: a stamp is DONE (green) when its ha were sprayed, OVERDUE (red) when its day has passed and it is not done, otherwise PLANNED (grey; today's are DUE).
 *  People with the scheduling task drag a grey or red stamp to another day of the month (today or later) and half-day; a green stamp cannot be moved. A morning or evening block holds
 *  at most `foliar_block_max_ha` (5 ha) and a day at most `foliar_day_max_ha` (9 ha).
 *  Only drone parts are scheduled for now (no plan for the manual parts yet).
 */
const STAMP_TOL = 0.05;
const HALVES = ['AM', 'PM'];

function weekdayOf_(date) {
  const d = new Date(date + 'T00:00:00Z').getUTCDay();
  return d === 0 ? 7 : d;
} // 1 = Monday ... 7 = Sunday
function monthKey_(date) {
  return String(date).slice(0, 7);
}
function validMonth_(m) {
  return /^\d{4}-(0[1-9]|1[0-2])$/.test(String(m));
}
function monthDays_(month) {
  return new Date(Date.UTC(+month.slice(0, 4), +month.slice(5, 7), 0)).getUTCDate();
}
function addMonths_(month, k) {
  const y = +month.slice(0, 4),
    m = +month.slice(5, 7) - 1 + k;
  return new Date(Date.UTC(y, m, 1)).toISOString().slice(0, 7);
}
/** The date of the n-th occurrence (1 = first) of a weekday in a month, or '' when the month does not have one. */
function nthWeekday_(month, weekday, n) {
  const off = (weekday - weekdayOf_(month + '-01') + 7) % 7,
    day = 1 + off + 7 * (n - 1);
  return day <= monthDays_(month) ? month + '-' + padNum_(day, 2) : '';
}
/** Short name of a division for a stamp: New Peacock -> NP. */
function divAbbr_(name) {
  const w = String(name || '')
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  return w.length > 1
    ? w
        .map(function (x) {
          return x.charAt(0);
        })
        .join('')
        .toUpperCase()
        .slice(0, 3)
    : String(name || '')
        .slice(0, 2)
        .toUpperCase();
}

function templateAll_() {
  return dbAll_(SHEETS.TEMPLATE).filter(function (r) {
    return r.status === 'Active';
  });
}
function stampsOfMonth_(month, rows) {
  return (rows || dbAll_(SHEETS.STAMPS)).filter(function (s) {
    return s.month === month && s.status === 'Active';
  });
}

/** Make the stamps of a month from the template: every active block of the template that has no stamp yet in this month gets one (so a month is made the first time it is opened, and made
 *  again after the plan file was loaded again, which voids the old stamps and writes a new template). Only this month and the next two can be made; a month already past stays as it is.
 *  Returns true when it made some. */
function ensureMonth_(month, today, stampRows) {
  const cur = monthKey_(today);
  if (month < cur || month > addMonths_(cur, 2)) return false;
  const missingOf = function (tpl, rows) {
    const have = {};
    (rows || dbAll_(SHEETS.STAMPS)).forEach(function (s) {
      if (s.month === month) have[s.template_id] = true;
    });
    return tpl.filter(function (t) {
      return !have[t.slot_id] && nthWeekday_(month, t.weekday, t.week_no);
    });
  };
  const first = templateAll_();
  if (!first.length) return false;
  if (!missingOf(first, stampRows).length) return false; // (stampRows: the rows the caller has already read, to save a read of the sheet)
  return withLock_(function () {
    const todo = missingOf(templateAll_());
    if (!todo.length) return false; // read again under the lock
    const now = nowIso_(),
      rows = [];
    todo.forEach(function (t) {
      const date = nthWeekday_(month, t.weekday, t.week_no);
      rows.push({
        stamp_id: newId_('STM'),
        month: month,
        field_id: t.field_id,
        part: t.part || 'Drone',
        ha: t.ha,
        date: date,
        half: t.half,
        orig_date: date,
        orig_half: t.half,
        template_id: t.slot_id,
        status: 'Active',
        moved_by: '',
        moved_at: '',
        created_at: now
      });
    });
    dbAppend_(SHEETS.STAMPS, rows);
    return rows.length > 0;
  });
}

/** The stamps of a month with their state worked out from the recorded spraying: done_ha, done_on, status ('done' | 'overdue' | 'due' | 'planned'). */
function stampStates_(month, today, stampRows, sprayRows) {
  const stamps = stampsOfMonth_(month, stampRows).map(function (s) {
    return Object.assign({}, s);
  });
  const lines = {};
  (sprayRows || dbAll_(SHEETS.SPRAY)).forEach(function (l) {
    if (l.status !== 'Active' || l.method !== 'Drone' || monthKey_(l.date) !== month) return;
    (lines[l.field_id] = lines[l.field_id] || []).push({ date: l.date, ha: Number(l.area_ha) || 0 });
  });
  const byField = groupBy_(stamps, function (s) {
    return s.field_id;
  });
  Object.keys(byField).forEach(function (fid) {
    const q = (lines[fid] || []).sort(function (a, b) {
      return a.date < b.date ? -1 : a.date > b.date ? 1 : 0;
    });
    byField[fid]
      .sort(function (a, b) {
        return a.orig_date === b.orig_date
          ? a.orig_half < b.orig_half
            ? -1
            : a.orig_half > b.orig_half
              ? 1
              : 0
          : a.orig_date < b.orig_date
            ? -1
            : 1;
      })
      .forEach(function (s) {
        let done = 0,
          last = '';
        while (s.ha - done > 1e-9 && q.length) {
          const take = Math.min(q[0].ha, s.ha - done);
          done += take;
          q[0].ha -= take;
          last = q[0].date;
          if (q[0].ha < 1e-9) q.shift();
        }
        s.done_ha = round_(done, 3);
        const complete = s.ha - done <= STAMP_TOL;
        s.done_on = complete ? last : '';
        s.state = complete ? 'done' : s.date < today ? 'overdue' : s.date === today ? 'due' : 'planned';
      });
  });
  return stamps;
}

function shapeStamp_(s, idx) {
  const f = idx[s.field_id];
  return {
    stamp_id: s.stamp_id,
    field_id: s.field_id,
    code: f ? f.code : '?',
    label: f ? fieldLabel_(f) : '?',
    division: f ? f.division : '',
    div_abbr: f ? divAbbr_(f.division) : '',
    field_ha: f ? f.area_ha : 0,
    drone_ha: f ? f.drone_ha : 0,
    part: s.part,
    ha: s.ha,
    date: s.date,
    half: s.half,
    orig_date: s.orig_date,
    orig_half: s.orig_half,
    moved: s.date !== s.orig_date || s.half !== s.orig_half,
    done_ha: s.done_ha,
    done_on: s.done_on,
    state: s.state
  };
}

/** The foliar calendar of a month: the weekday x week grid (dates), every stamp with its colour, and what the person may do. */
function api_foliarCalendar(token, month) {
  return wrap_(function () {
    const c = authz_(token, ['foliar.calendar', 'foliar.schedule']);
    const today = todayStr_(),
      params = getParams_(today);
    month = validMonth_(month) ? String(month) : monthKey_(today);
    const made = ensureMonth_(month, today),
      idx = fieldIndex_();
    const every = stampStates_(month, today),
      loadBlock = {},
      loadDay = {}; // what is already planned in each block and day, whoever's fields they are (the limits are for the drones)
    every.forEach(function (s) {
      loadBlock[s.date + ' ' + s.half] = round_((loadBlock[s.date + ' ' + s.half] || 0) + s.ha, 3);
      loadDay[s.date] = round_((loadDay[s.date] || 0) + s.ha, 3);
    });
    const stamps = every
      .filter(function (s) {
        return idx[s.field_id] && inScope_(c, idx[s.field_id]);
      })
      .map(function (s) {
        return shapeStamp_(s, idx);
      }); // a person sees only the stamps of their own fields
    const days = [];
    for (let n = 1; n <= 5; n++)
      for (let w = 1; w <= 7; w++) {
        const d = nthWeekday_(month, w, n);
        if (d) days.push({ date: d, weekday: w, week_no: n });
      }
    const sum = function (k) {
      return round_(
        stamps.reduce(function (t, s) {
          return t + (k === 'done' ? Math.min(s.done_ha, s.ha) : s.ha);
        }, 0),
        2
      );
    };
    return {
      month: month,
      today: today,
      can_edit: hasTask_(c, 'foliar.schedule'),
      block_max: params.foliar_block_max_ha,
      day_max: params.foliar_day_max_ha,
      days: days,
      stamps: stamps,
      load_block: loadBlock,
      load_day: loadDay,
      prev: addMonths_(month, -1),
      next: addMonths_(month, 1),
      made: made,
      has_template: templateAll_().length > 0,
      planned_ha: sum('plan'),
      done_ha: sum('done'),
      divisions: dbAll_(SHEETS.DIVISIONS)
        .filter(function (d) {
          return d.active;
        })
        .sort(function (a, b) {
          return a.seq - b.seq;
        })
        .map(function (d) {
          return { name: d.name, abbr: divAbbr_(d.name) };
        })
    };
  });
}

/** Move a stamp to another day (today or later, in the same month) and half-day. Not a done stamp; not into a block or a day that would hold too much. */
function api_moveFoliarStamp(token, p) {
  return wrap_(function () {
    const c = authz_(token, ['foliar.schedule']);
    p = p || {};
    const today = todayStr_(),
      params = getParams_(today);
    return withLock_(function () {
      const rows = dbAll_(SHEETS.STAMPS),
        cur = rows.filter(function (s) {
          return s.stamp_id === String(p.stamp_id || '') && s.status === 'Active';
        })[0];
      if (!cur) fail_('Stamp not found.', 'VALIDATION');
      const fld = fieldIndex_()[cur.field_id];
      if (!fld || !inScope_(c, fld)) fail_('This field is not one of your fields.', 'FORBIDDEN');
      const date = String(p.date || ''),
        half = String(p.half || '');
      if (!isDate_(date)) fail_('Choose a day.', 'VALIDATION');
      if (HALVES.indexOf(half) < 0) fail_('Choose morning or evening.', 'VALIDATION');
      if (monthKey_(date) !== cur.month) fail_('A stamp can only be moved inside the month it belongs to.', 'VALIDATION');
      if (date < today) fail_('A stamp can only be moved to today or a later day.', 'VALIDATION');
      const st = stampStates_(cur.month, today, rows).filter(function (s) {
        return s.stamp_id === cur.stamp_id;
      })[0];
      if (st.state === 'done') fail_('This stamp is green: its spraying is done and it cannot be moved.', 'VALIDATION');
      if (cur.date === date && cur.half === half) return { moved: false };
      const others = stampsOfMonth_(cur.month, rows).filter(function (s) {
        return s.stamp_id !== cur.stamp_id && s.date === date;
      });
      const block = others
          .filter(function (s) {
            return s.half === half;
          })
          .reduce(function (t, s) {
            return t + s.ha;
          }, 0),
        day = others.reduce(function (t, s) {
          return t + s.ha;
        }, 0);
      if (block + cur.ha > params.foliar_block_max_ha + 1e-9)
        fail_(
          'The ' +
            (half === 'AM' ? 'morning' : 'evening') +
            ' of ' +
            date +
            ' would hold ' +
            round_(block + cur.ha, 2) +
            ' ha; the most is ' +
            params.foliar_block_max_ha +
            ' ha.',
          'VALIDATION'
        );
      if (day + cur.ha > params.foliar_day_max_ha + 1e-9)
        fail_(date + ' would hold ' + round_(day + cur.ha, 2) + ' ha; the most in a day is ' + params.foliar_day_max_ha + ' ha.', 'VALIDATION');
      const now = nowIso_();
      dbUpdate_(SHEETS.STAMPS, cur._row, { date: date, half: half, moved_by: c.user.username, moved_at: now });
      audit_(c.user.username, 'FOLIAR_STAMP_MOVE', 'foliar_stamp', cur.stamp_id, {
        field: fieldLabel_(fld),
        ha: cur.ha,
        from: cur.date + ' ' + cur.half,
        to: date + ' ' + half
      });
      return { moved: true };
    });
  });
}

/** What a person on the Foliar spray screen sees for a day: the stamps planned for that day (for today also every overdue stamp of this and last month), only on fields they may record. */
function dueFor_(c, date, today, idx, sprayRows) {
  const out = [],
    months = [monthKey_(date)];
  if (date === today) months.push(addMonths_(monthKey_(today), -1));
  let rows = dbAll_(SHEETS.STAMPS); // read once for both months (and the spray lines come from the caller)
  if (months[0] === monthKey_(today)) {
    try {
      if (ensureMonth_(months[0], today, rows)) rows = dbAll_(SHEETS.STAMPS);
    } catch (e) {
      /* the screen still works without the plan */
    }
  }
  months.forEach(function (m) {
    stampStates_(m, today, rows, sprayRows).forEach(function (s) {
      const f = idx[s.field_id];
      if (!f || !inScope_(c, f)) return;
      if (s.date === date || (date === today && s.state === 'overdue')) out.push(shapeStamp_(s, idx));
    });
  });
  return out.sort(function (a, b) {
    return a.date === b.date ? (a.half < b.half ? -1 : a.half > b.half ? 1 : a.label < b.label ? -1 : 1) : a.date < b.date ? -1 : 1;
  });
}

// ------------------------------------------------------------------ the plan file: field drone ha + the template

/** Compare the plan file with the fields on the sheet. Nothing is written. */
function planFileCompare_() {
  const P = FOLIAR_PLAN_FILE,
    divs = {};
  dbAll_(SHEETS.DIVISIONS).forEach(function (d) {
    divs[String(d.name).toLowerCase()] = d;
  });
  const fidx = fieldIndex_(),
    fields = Object.keys(fidx)
      .map(function (k) {
        return fidx[k];
      })
      .filter(function (f) {
        return f.active;
      }),
    byKey = {}; // with the division name joined
  fields.forEach(function (f) {
    byKey[f.division_id + '|' + String(f.code).toUpperCase().replace(/\s+/g, '')] = f;
  });
  const find = function (abbr, code) {
    const d = divs[String(P.divisions[abbr] || '').toLowerCase()];
    return d ? byKey[d.division_id + '|' + code] : null;
  };
  const missing = [],
    problems = [],
    blocking = [],
    change = [],
    target = {};
  Object.keys(P.drone).forEach(function (k) {
    const f = find(k.slice(0, 2), k.slice(3));
    if (!f) missing.push(k);
    else target[f.field_id] = P.drone[k];
  });
  const idx = {};
  fields.forEach(function (f) {
    idx[f.field_id] = f;
  });
  fields.forEach(function (f) {
    const to = target[f.field_id] === undefined ? 0 : target[f.field_id],
      from = f.drone_ha === '' ? 0 : f.drone_ha;
    const inPlan = target[f.field_id] !== undefined;
    if (!(f.area_ha > 0)) {
      problems.push(fieldLabel_(f) + ': no area yet');
      if (inPlan) blocking.push(fieldLabel_(f) + ': no area yet');
      return;
    }
    if (to > f.area_ha + AREA_TOL) {
      const msg = fieldLabel_(f) + ': the plan has ' + to + ' ha by drone but the field area is ' + f.area_ha + ' ha';
      problems.push(msg);
      blocking.push(msg);
      return;
    }
    if (Math.abs(to - from) > 0.0005) change.push({ field_id: f.field_id, label: fieldLabel_(f), area_ha: f.area_ha, from: from, to: to });
  });
  const slots = [];
  P.slots.forEach(function (s) {
    const f = find(s.div, s.f);
    if (!f) {
      if (missing.indexOf(s.div + '/' + s.f) < 0) missing.push(s.div + '/' + s.f);
      return;
    }
    slots.push({ field_id: f.field_id, label: fieldLabel_(f), weekday: s.weekday, week_no: s.week, half: s.half, ha: s.ha });
  });
  return {
    missing: missing,
    problems: problems,
    blocking: blocking,
    change: change,
    slots: slots,
    slot_ha: round_(
      slots.reduce(function (t, s) {
        return t + s.ha;
      }, 0),
      2
    ),
    drone_fields: Object.keys(target).length,
    drone_ha_now: round_(
      fields.reduce(function (t, f) {
        return t + (f.drone_ha > 0 ? f.drone_ha : 0);
      }, 0),
      2
    ),
    drone_ha_plan: round_(
      Object.keys(target).reduce(function (t, id) {
        return t + target[id];
      }, 0),
      2
    ),
    target: target
  };
}
function planFileSummary_(r) {
  return {
    missing: r.missing,
    problems: r.problems,
    changes: r.change.map(function (x) {
      return { label: x.label, area_ha: x.area_ha, from: x.from, to: x.to };
    }),
    slot_count: r.slots.length,
    slot_ha: r.slot_ha,
    drone_fields: r.drone_fields,
    drone_ha_now: r.drone_ha_now,
    drone_ha_plan: r.drone_ha_plan
  };
}
/** Write it: the drone ha of every field (fields not in the plan: 0, so sprayed by hand), the template, and drop the stamps of this and later months so they are made again from the new template. Field areas are not touched. */
function planFileApply_(who) {
  return withLock_(function () {
    const r = planFileCompare_();
    if (r.missing.length)
      fail_('These fields of the plan file are not on the sheet (create them first): ' + r.missing.join(', ') + '.', 'VALIDATION');
    if (r.blocking.length) fail_('The plan file cannot be loaded yet: ' + r.blocking.join('; ') + '.', 'VALIDATION');
    const now = nowIso_(),
      today = todayStr_();
    const byId = {};
    dbAll_(SHEETS.FIELDS).forEach(function (f) {
      byId[f.field_id] = f;
    });
    dbUpdateMany_(
      SHEETS.FIELDS,
      r.change.map(function (x) {
        return { row: byId[x.field_id]._row, patch: { drone_ha: x.to, updated_at: now } };
      })
    );
    dbUpdateMany_(
      SHEETS.TEMPLATE,
      dbAll_(SHEETS.TEMPLATE)
        .filter(function (t) {
          return t.status === 'Active';
        })
        .map(function (t) {
          return { row: t._row, patch: { status: 'Replaced' } };
        })
    );
    dbAppend_(
      SHEETS.TEMPLATE,
      r.slots.map(function (s) {
        return {
          slot_id: newId_('TPL'),
          field_id: s.field_id,
          part: 'Drone',
          weekday: s.weekday,
          week_no: s.week_no,
          half: s.half,
          ha: s.ha,
          status: 'Active'
        };
      })
    );
    const cur = monthKey_(today),
      stale = dbAll_(SHEETS.STAMPS).filter(function (s) {
        return s.status === 'Active' && s.month >= cur;
      }),
      dropped = stale.length;
    dbUpdateMany_(
      SHEETS.STAMPS,
      stale.map(function (s) {
        return { row: s._row, patch: { status: 'Voided' } };
      })
    );
    audit_(who, 'FOLIAR_PLAN_FILE', 'foliar_plan_file', '', {
      fields_changed: r.change.length,
      drone_ha: r.drone_ha_plan,
      slots: r.slots.length,
      stamps_dropped: dropped
    });
    return { fields_changed: r.change.length, slots: r.slots.length, slot_ha: r.slot_ha, stamps_dropped: dropped };
  });
}
function api_foliarPlanFilePreview(token) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    if (!c.isAdmin) fail_('Only the Admin can load the plan file.', 'FORBIDDEN');
    return planFileSummary_(planFileCompare_());
  });
}
function api_foliarPlanFileApply(token) {
  return wrap_(function () {
    const c = authz_(token, ['estate.manage']);
    if (!c.isAdmin) fail_('Only the Admin can load the plan file.', 'FORBIDDEN');
    return planFileApply_(c.user.username);
  });
}
/** Owner, editor: what loading the plan file would change (writes nothing). */
function previewFoliarPlanFile() {
  requireOwner_();
  const s = planFileSummary_(planFileCompare_());
  Logger.log(JSON.stringify(s, null, 1));
  return s;
}
/** Owner, editor: load the plan file (same as the Admin's button on Foliar planning). */
function loadFoliarPlanFile() {
  requireOwner_();
  const r = planFileApply_('owner');
  Logger.log(JSON.stringify(r));
  return r;
}

// ---- History.gs
/** Field history for the people who enter data: AFOs see the rounds and entries of THEIR fields, drone pilots those of the fields with a drone part (foliar, drone
 *  part only). Same rounds as everywhere else (Records.gs buildRounds_ with the windows in Parameters); a field sprayed partly by drone and partly by hand shows its
 *  drone part and its manual part separately, each with its own rounds. Read only. Anyone with the task history.view (given automatically to every role with
 *  foliar spray or plucking entry) may call it, for the fields in their scope. */

/** Which kinds of history this person may see. */
function historyKinds_(c) {
  const k = [];
  const watches = hasTask_(c, 'monitor.operation') || hasTask_(c, 'monitor.experiment');
  if (c.isAdmin || hasTask_(c, 'spray.entry') || watches || hasTask_(c, 'records.manage')) k.push('spray');
  if (c.isAdmin || hasTask_(c, 'pluck.entry') || watches || hasTask_(c, 'records.manage')) k.push('pluck');
  return k;
}
/** The parts of a field this person sees in a kind of history: a pilot sees only the drone part of the foliar spray. */
function historyParts_(c, f, kind) {
  const parts = partMethods_(f);
  return kind === 'spray'
    ? parts.filter(function (m) {
        return allowedMethods_(c).indexOf(m) >= 0;
      })
    : parts;
}

/** The fields of this person with a one-line summary each: what the History screen lists. */
function api_historyFields(token) {
  return wrap_(function () {
    const c = authz_(token, ['history.view']),
      kinds = historyKinds_(c),
      today = todayStr_(),
      P = getParams_(today),
      idx = fieldIndex_();
    const recs = {};
    kinds.forEach(function (k) {
      const rows =
          k === 'spray'
            ? sprayLinesWithOrder_(dbAll_(SHEETS.SPRAY_EVENTS))
            : dbAll_(SHEETS.LEAF).filter(function (r) {
                return r.status === 'Active';
              }),
        by = {};
      rows.forEach(function (r) {
        (by[r.field_id] = by[r.field_id] || []).push(r);
      });
      recs[k] = by;
    });
    const fields = Object.keys(idx)
      .map(function (id) {
        return idx[id];
      })
      .filter(function (f) {
        return usableField_(c, f);
      })
      .sort(function (a, b) {
        return a.division_seq - b.division_seq || a.seq - b.seq;
      })
      .map(function (f) {
        const o = {
          field_id: f.field_id,
          code: f.code,
          name: f.name,
          division: f.division,
          area_ha: f.area_ha,
          method: f.method,
          drone_ha: f.drone_ha,
          manual_ha: f.manual_ha
        };
        kinds.forEach(function (k) {
          const mine = (recs[k][f.field_id] || []).filter(function (r) {
            return k !== 'spray' || historyParts_(c, f, k).indexOf(r.method) >= 0;
          });
          let last = '';
          mine.forEach(function (r) {
            if (r.date > last) last = r.date;
          });
          let n = 0,
            done = 0;
          historyParts_(c, f, k).forEach(function (m) {
            buildRounds_(partRecs_(f, m, mine), partExt_(f, m), resetDays_(k, P)).forEach(function (rd) {
              n++;
              if (rd.complete) done++;
            });
          });
          o[k] = { last: last, rounds: n, complete: done };
        });
        return o;
      });
    return { today: today, kinds: kinds, fields: fields };
  });
}

/** One field: for each kind of history (foliar spray, plucking) and each part of the field, its rounds, newest first, with every entry. */
function api_fieldHistory(token, fieldId) {
  return wrap_(function () {
    const c = authz_(token, ['history.view']),
      idx = fieldIndex_(),
      f = idx[String(fieldId || '')];
    if (!f || !f.active || !usableField_(c, f)) fail_('Field not found.', 'FORBIDDEN');
    const today = todayStr_(),
      P = getParams_(today),
      out = {
        today: today,
        field: {
          field_id: f.field_id,
          code: f.code,
          name: f.name,
          division: f.division,
          area_ha: f.area_ha,
          method: f.method,
          drone_ha: f.drone_ha,
          manual_ha: f.manual_ha
        },
        kinds: []
      };
    historyKinds_(c).forEach(function (kind) {
      const rows =
        kind === 'spray'
          ? sprayLinesWithOrder_(dbAll_(SHEETS.SPRAY_EVENTS)).filter(function (r) {
              return r.field_id === f.field_id;
            })
          : dbAll_(SHEETS.LEAF).filter(function (r) {
              return r.status === 'Active' && r.field_id === f.field_id;
            });
      const target = kind === 'spray' ? P.spray_target_gap_days : P.pluck_target_gap_days,
        reset = resetDays_(kind, P);
      const parts = historyParts_(c, f, kind).map(function (m) {
        const sh = shapeRounds_(kind, f, m, rows, reset, target, today, true),
          list = sh.rounds.slice().reverse();
        return {
          method: m,
          ha: round_(sh.ext, 3),
          n_rounds: list.length,
          n_complete: list.filter(function (r) {
            return r.complete;
          }).length,
          rounds: list
        };
      });
      out.kinds.push({ kind: kind, target_gap: target, reset_days: reset, parts: parts });
    });
    return out;
  });
}

// ---- Home.gs
/** Where an entry stands in its round, the way the entry windows show it: the ha of the part, the running % of the round after this entry and before it (the slider),
 *  the round number and day, and whether the round was finished by this entry. */
function homeEntryRound_(partRecords, ext, resetDays, rec) {
  let out = { ext: ext, pct: 0, before: 0, round_no: 0, round_day: 1, complete: false };
  buildRounds_(partRecords, ext, resetDays).forEach(function (rd) {
    let cum = 0;
    rd.records.some(function (x) {
      if (x.record_id !== rec.record_id) {
        cum += Number(x.area_ha) || 0;
        return false;
      }
      const after = cum + (Number(x.area_ha) || 0);
      out = {
        ext: ext,
        pct: ext > 0 ? Math.min(100, round_((after / ext) * 100, 1)) : 0,
        before: ext > 0 ? Math.min(100, round_((cum / ext) * 100, 1)) : 0,
        round_no: rd.no,
        round_day: daysBetween_(rd.start, rec.date) + 1,
        complete: rd.complete && rd.end === rec.date
      };
      return true;
    });
  });
  return out;
}

/** Home: the estate at a glance, for EVERYONE who is signed in (fields and their spray method, today's plucking and spray entries and who made them).
 *  The setup checklist stays Admin-only (api_setupStatus). Read-only. */
function api_homeSummary(token) {
  return wrap_(function () {
    const c = authz_(token, null);
    const today = todayStr_(),
      names = userNames_(),
      who = function (n) {
        return names[n] || n || '';
      };
    const estate = dbAll_(SHEETS.ESTATE)[0];
    const divs = dbAll_(SHEETS.DIVISIONS)
        .filter(function (d) {
          return d.active !== false;
        })
        .sort(function (a, b) {
          return (a.seq || 0) - (b.seq || 0);
        }),
      shown = {};
    divs.forEach(function (d) {
      shown[d.division_id] = true;
    });
    const fields = fieldsAll_().filter(function (f) {
      return f.active && shown[f.division_id] && inScope_(c, f);
    }); // only the fields this person is responsible for, in the divisions that are listed (so the totals match the list)
    const divisions = divs.map(function (d) {
      return {
        division_id: d.division_id,
        name: d.name,
        fields: fields
          .filter(function (f) {
            return f.division_id === d.division_id;
          })
          .sort(function (a, b) {
            return (a.seq || 0) - (b.seq || 0);
          })
          .map(function (f) {
            return {
              field_id: f.field_id,
              code: f.code,
              name: f.name || '',
              area_ha: f.area_ha > 0 ? f.area_ha : 0,
              drone_ha: partHa_(f, 'Drone'),
              manual_ha: partHa_(f, 'Manual'),
              method: f.method || ''
            };
          })
      };
    });
    const withDrone = fields.filter(function (f) {
        return partHa_(f, 'Drone') > 0;
      }),
      withManual = fields.filter(function (f) {
        return partHa_(f, 'Manual') > 0;
      });
    const sumHa = function (l, m) {
      return round_(
        l.reduce(function (s, f) {
          return s + partHa_(f, m);
        }, 0),
        1
      );
    };
    const idx = fieldIndex_(),
      params = getParams_(today),
      hhmm = function (t) {
        return String(t || '').slice(11, 16);
      };
    const leafAll = dbAll_(SHEETS.LEAF).filter(function (r) {
        return r.status === 'Active';
      }),
      leafBy = groupBy_(leafAll, function (r) {
        return r.field_id;
      });
    const leaf = leafAll
      .filter(function (r) {
        return r.date === today && idx[r.field_id] && inScope_(c, idx[r.field_id]);
      })
      .map(function (r) {
        const f = idx[r.field_id],
          ext = partExt_(f, f.method === 'Mixed' ? r.method : f.method);
        return Object.assign(
          {
            field_id: r.field_id,
            method: r.method,
            kg: round_(Number(r.kg) || 0, 1),
            ha: round_(Number(r.area_ha) || 0, 2),
            by: who(r.entered_by),
            at: hhmm(r.entered_at)
          },
          homeEntryRound_(partRecs_(f, r.method, leafBy[r.field_id] || []), ext, resetDays_('pluck', params), r)
        );
      });
    const sprayAll = sprayLinesWithOrder_(dbAll_(SHEETS.SPRAY_EVENTS)),
      sprayBy = groupBy_(sprayAll, function (r) {
        return r.field_id;
      });
    const spray = sprayAll
      .filter(function (r) {
        return r.date === today && idx[r.field_id] && inScope_(c, idx[r.field_id]);
      })
      .map(function (r) {
        const f = idx[r.field_id],
          ext = partExt_(f, f.method === 'Mixed' ? r.method : f.method);
        return Object.assign(
          {
            field_id: r.field_id,
            method: r.method,
            ha: round_(Number(r.area_ha) || 0, 2),
            kg: round_(Number(r.kg_applied) || 0, 2),
            start: r.start_time,
            end: r.end_time,
            by: who(r.entered_by),
            at: hhmm(r.entered_at),
            event_id: r.event_id
          },
          homeEntryRound_(partRecs_(f, r.method, sprayBy[r.field_id] || []), ext, resetDays_('spray', params), r)
        );
      });
    const uniq = function (l) {
      const o = {};
      l.forEach(function (r) {
        o[r.field_id] = true;
      });
      return Object.keys(o).length;
    };
    return {
      today: today,
      estate: estate ? estate.name : '',
      stats: {
        fields: fields.length,
        ha: round_(
          fields.reduce(function (s, f) {
            return s + (f.area_ha > 0 ? f.area_ha : 0);
          }, 0),
          1
        ),
        drone: withDrone.length,
        drone_ha: sumHa(withDrone, 'Drone'),
        manual: withManual.length,
        manual_ha: sumHa(withManual, 'Manual'),
        mixed: fields.filter(function (f) {
          return f.method === 'Mixed';
        }).length,
        plucked_fields: uniq(leaf),
        kg_today: round_(
          leaf.reduce(function (s, r) {
            return s + r.kg;
          }, 0),
          0
        ),
        spray_events: Object.keys(
          spray.reduce(function (o, r) {
            o[r.event_id || r.field_id] = 1;
            return o;
          }, {})
        ).length,
        spray_fields: uniq(spray),
        spray_ha: round_(
          spray.reduce(function (s, r) {
            return s + r.ha;
          }, 0),
          1
        )
      },
      divisions: divisions,
      pluck_today: leaf,
      spray_today: spray
    };
  });
}

// ---- Monitoring.gs
/** Operation monitoring: the ROUNDS of every field on a timeline (plucking and foliar).
 *
 *  A round is worked out from the records exactly as everywhere else (Records.gs buildRounds_): a field's records, in date order, add up to the
 *  field's area; the round is complete when the whole field has been covered. For each round this gives the days it took, the kg, the ha done,
 *  and the GAP = days from the first day of the previous round to the first day of this one. A gap longer than the target (Parameters:
 *  pluck_target_gap_days, 7 by default) is LATE by the days beyond it. Per month: how many rounds STARTED in it. The average per month uses
 *  only full calendar months (after the month of the field's first round and before the current month).
 *
 *  A round left unfinished for more than the window (spray_round_reset_days / pluck_round_reset_days, the days between two entries) is closed as it is and flagged `expired`: the next
 *  record on that part starts a new round (Records.gs buildRounds_). Foliar rounds use the same timeline: the target gap is spray_target_gap_days
 *  (a setting); the kg of a round is the foliar applied. A field sprayed partly by drone and partly by hand has TWO ROWS (drone, manual), each with its
 *  own rounds, gaps and counts: the drone part can have had two rounds while the manual part had one. */

function monthList_(fromMonth, toMonth) {
  const out = [];
  let y = Number(fromMonth.slice(0, 4)),
    m = Number(fromMonth.slice(5, 7));
  const ey = Number(toMonth.slice(0, 4)),
    em = Number(toMonth.slice(5, 7));
  while (y < ey || (y === ey && m <= em)) {
    out.push(y + '-' + (m < 10 ? '0' : '') + m);
    m++;
    if (m > 12) {
      m = 1;
      y++;
    }
  }
  return out;
}

function api_pluckRounds(token, q) {
  return wrap_(function () {
    const c = authz_(token, ['monitor.operation']);
    return roundsTimeline_('pluck', q || {}, c);
  });
}

function api_sprayRounds(token, q) {
  return wrap_(function () {
    const c = authz_(token, ['monitor.operation']);
    return roundsTimeline_('spray', q || {}, c);
  });
}

/** The rounds of ONE PART of a field (the whole field, or its drone or manual part), shaped for the screens: dates, days, % covered, kg, gap to the previous
 *  round and whether it was left unfinished. withEntries adds every entry of the round (for the field history): date, ha, kg, who and, for foliar, the event. */
function shapeRounds_(type, f, m, recs, reset, target, today, withEntries) {
  const ext = partExt_(f, m),
    rounds = buildRounds_(partRecs_(f, m, recs), ext, reset);
  const names = withEntries ? userNames_() : null;
  const shaped = rounds.map(function (rd, i) {
    let kg = 0,
      events = {};
    rd.records.forEach(function (r) {
      kg += (type === 'spray' ? r.kg_applied : r.kg) || 0;
      if (r.event_id) events[r.event_id] = true;
    });
    const expired = rd.expired || (!rd.complete && daysBetween_(rd.last, today) > reset);
    const prev = rounds[i - 1],
      gap = prev ? daysBetween_(prev.start, rd.start) : '';
    const o = {
      no: rd.no,
      start: rd.start,
      end: rd.complete ? rd.end : rd.last,
      days: rd.days,
      complete: rd.complete,
      expired: expired,
      events: Object.keys(events).length,
      kg: round_(kg, 1),
      kg_per_ha: rd.done_ha > 0 ? round_(kg / rd.done_ha, 1) : '',
      done_ha: round_(rd.done_ha, 2),
      ha: round_(ext, 3),
      percent: rd.complete ? 100 : ext > 0 ? Math.min(99, Math.round((rd.done_ha / ext) * 100)) : 0,
      gap: gap,
      late: gap !== '' && gap > target ? gap - target : 0
    };
    if (withEntries) {
      let cum = 0;
      o.entries = rd.records.map(function (r) {
        cum = round_(cum + r.area_ha, 3);
        const e = {
          date: r.date,
          ha: round_(r.area_ha, 2),
          kg: round_(type === 'spray' ? r.kg_applied : r.kg, 2),
          reached: ext > 0 ? (cum >= ext - ROUND_TOL ? 100 : Math.min(99, Math.round((cum / ext) * 100))) : 0,
          by: names[r.entered_by] || r.entered_by
        };
        if (type === 'spray') {
          e.hours = r.hours;
          e.units = r.units;
          e.start_time = r.start_time;
          e.end_time = r.end_time;
        }
        return e;
      });
    }
    return o;
  });
  return { ext: ext, rounds: shaped };
}
/** username -> the person's name, for the current username and every earlier one (entries keep the username of the day they were made). */
function userNames_() {
  const m = {},
    users = dbAll_(SHEETS.USERS);
  users.forEach(function (u) {
    userPrevNames_(u).forEach(function (n) {
      m[n] = u.display_name || u.username;
    });
  });
  users.forEach(function (u) {
    m[u.username] = u.display_name || u.username;
  });
  return m;
}

function roundsTimeline_(type, q, c) {
  const today = todayStr_(),
    P = getParams_(today),
    target = type === 'spray' ? P.spray_target_gap_days : P.pluck_target_gap_days,
    reset = resetDays_(type, P);
  const idx = fieldIndex_();
  const recs =
    type === 'spray'
      ? sprayLinesWithOrder_(dbAll_(SHEETS.SPRAY_EVENTS))
      : dbAll_(recSheet_(type)).filter(function (r) {
          return r.status === 'Active';
        });
  const by = {};
  let minDate = '';
  recs.forEach(function (r) {
    (by[r.field_id] = by[r.field_id] || []).push(r);
    if (!minDate || r.date < minDate) minDate = r.date;
  });
  const to = q.to && isDate_(q.to) ? q.to : today;
  let from =
    q.from && isDate_(q.from) ? q.from : minDate ? (minDate > addDays_(today, -180) ? minDate : addDays_(today, -180)) : addDays_(today, -60);
  if (from > to) from = addDays_(to, -60);
  if (daysBetween_(from, to) > 800) from = addDays_(to, -800);
  const fields = Object.keys(idx)
    .map(function (k) {
      return idx[k];
    })
    .filter(function (f) {
      return f.active && fieldComplete_(f) && (!c || inScope_(c, f));
    }) // a person sees the rounds of their own fields only
    .sort(function (a, b) {
      return a.division_seq - b.division_seq || a.seq - b.seq;
    });
  const thisMonth = monthKey_(today),
    months = monthList_(monthKey_(from), monthKey_(to));
  const out = fields.map(function (f) {
    // one ROW per part of the field: a field with one method has one row; a Mixed field has a drone row and a manual row, each with its OWN rounds, gaps and counts
    const rows = partMethods_(f).map(function (m) {
      const sh = shapeRounds_(type, f, m, by[f.field_id] || [], reset, target, today, false),
        ext = sh.ext,
        shaped = sh.rounds;
      const started = {};
      shaped.forEach(function (r) {
        const k = monthKey_(r.start);
        started[k] = (started[k] || 0) + 1;
      });
      let avg = '';
      if (shaped.length) {
        const full = monthList_(monthKey_(shaped[0].start), thisMonth).slice(1, -1); // whole months after the first round's month and before this month
        if (full.length) {
          let n = 0;
          full.forEach(function (mm) {
            n += started[mm] || 0;
          });
          avg = round_(n / full.length, 1);
        }
      }
      return {
        method: m,
        ha: round_(ext, 3),
        n_rounds: shaped.length,
        n_complete: shaped.filter(function (r) {
          return r.complete;
        }).length,
        avg_per_month: avg,
        rounds: shaped.filter(function (r) {
          return r.end >= from && r.start <= to;
        }),
        months: months.map(function (mm) {
          return { month: mm, n: started[mm] || 0 };
        })
      };
    });
    const one = rows.length === 1 ? rows[0] : null;
    return {
      field_id: f.field_id,
      code: f.code,
      label: fieldLabel_(f),
      division: f.division,
      method: f.method,
      drone_ha: f.drone_ha,
      drone_pct: f.area_ha ? Math.round((f.drone_ha / f.area_ha) * 100) : 0,
      area_ha: f.area_ha,
      rows: rows,
      n_rounds: rows.reduce(function (n, r) {
        return n + r.n_rounds;
      }, 0),
      rounds: one ? one.rounds : [],
      months: one ? one.months : [],
      avg_per_month: one ? one.avg_per_month : ''
    }; // (the last three: a field with one method only)
  });
  return { today: today, from: from, to: to, target_gap: target, reset_days: reset, fields: out };
}

/** Experiment monitoring: what the charts need, for ALL time. One row per day that has spray events or plucking, with each method's totals of that day:
 *  kg of foliar applied, ha covered, hours worked (and unit-hours = hours x men or drones), the number of events, and the green leaf plucked (leaf_kg from
 *  leaf_ha, by the method of the part of the field it came from). The charts add these up per day, month, year or all time and divide: foliar kg per ha (Foliar
 *  saving), ha per hour (Work speed), foliar material cost per ha (Material cost saving: kg applied x the dated price per kg), spraying labour cost per ha (Labour cost saving: hours x men x the man-hour rate, or hours x drones x the drone-hour rate, from the Admin's
 *  dated cost parameters) green leaf kg per ha (Leaf yield) and the revenue from the green leaf per ha (Yield revenue: kg plucked x the dated green leaf price per kg). Net margin is worked out on the screen from the SUMS of a period, never per day: (revenue - labour and foliar cost) / the land of the method, so the response also carries each method's land (ha) and the first day both kinds of data exist. A spray event is always one method, so a field sprayed partly by drone and
 *  partly by hand simply counts on both sides, each with its own events. */
function expCell_() {
  return { kg: 0, ha: 0, hours: 0, unit_hours: 0, events: 0, cost: 0, mat_cost: 0, leaf_kg: 0, leaf_ha: 0, leaf_n: 0, leaf_rev: 0 };
}
/** The cost rates the Admin entered (Parameters, group Cost rates), as a function of the date: each rate is the one in force on that day (dated parameters). */
function costRates_() {
  const rows = dbAll_(SHEETS.PARAMS);
  const pick = function (key, date) {
    let best = null;
    rows.forEach(function (r) {
      if (r.key === key && r.effective_from <= date && (!best || r.effective_from >= best.effective_from)) best = r;
    });
    const def = paramDef_(key);
    return best ? parseParam_(def, best.value) : def.def;
  };
  const memo = {}; // a day's rates are looked up once
  return function (date) {
    return (
      memo[date] ||
      (memo[date] = {
        basis: pick('drone_cost_basis', date),
        worker: Number(pick('worker_rate_per_hour', date)) || 0,
        drone_hr: Number(pick('drone_rate_per_hour', date)) || 0,
        drone_ha: Number(pick('drone_rate_per_ha', date)) || 0,
        price: Number(pick('product_unit_price', date)) || 0,
        gl: Number(pick('gl_price_per_kg', date)) || 0,
        compound: Object.keys(COMPOUND_PRICE_KEY).reduce(function (o, name) {
          o[name] = Number(pick(COMPOUND_PRICE_KEY[name], date)) || 0;
          return o;
        }, {})
      })
    );
  };
}
/** The foliar (material) cost of a spray event or of one field line: the kg of each of its four compounds x the price of that compound on the day (r = the rates of that day). An older
 *  event with one total kg (no compounds) is priced kg x the single foliar price. */
function foliarCost_(rec, r) {
  const mix = mixOf_(rec);
  if (!mix) return (Number(rec.kg_applied) || 0) * r.price;
  let t = 0;
  mix.forEach(function (m) {
    t += m.kg * (r.compound[m.name] || 0);
  });
  return t;
}
/** The LABOUR cost of one spray event: manual = hours x men x the cost of a man-hour; drone = hours x drones x the cost of a drone-hour (pilot included), or ha x the cost per ha when the Admin
 *  chose that basis. The foliar product itself is not included: that is the MATERIAL cost, kg applied x the price per kg of that day (see api_expCharts). */
function eventCost_(e, r) {
  const hrs = Number(e.hours) || 0;
  if (e.method === 'Drone') return r.basis === 'per_ha' ? (Number(e.total_ha) || 0) * r.drone_ha : hrs * (Number(e.drones) || 0) * r.drone_hr;
  return hrs * (Number(e.men) || 0) * r.worker;
}
/** True when something that was used in this event or line has no price in force on its day (a compound with kg but no price; an older event with one total kg and no single price): its foliar cost would be short or 0. */
function foliarUnpriced_(rec, r) {
  const mix = mixOf_(rec);
  if (!mix) return (Number(rec.kg_applied) || 0) > 0 && !(r.price > 0);
  return mix.some(function (m) {
    return m.kg > 0 && !(r.compound[m.name] > 0);
  });
}
function api_expCharts(token, q) {
  return wrap_(function () {
    authz_(token, ['monitor.experiment']);
    const byDay = {},
      rateAt = costRates_(),
      activeEv = {};
    let spraysSeen = 0,
      unpricedSeen = 0;
    dbAll_(SHEETS.SPRAY_EVENTS).forEach(function (e) {
      if (e.status !== 'Active' || METHODS.indexOf(e.method) < 0 || !isDate_(e.date)) return;
      activeEv[e.event_id] = e.method;
      spraysSeen++;
      if (foliarUnpriced_(e, rateAt(e.date))) unpricedSeen++;
      const d = byDay[e.date] || (byDay[e.date] = { date: e.date, Drone: expCell_(), Manual: expCell_() }),
        c = d[e.method],
        hrs = Number(e.hours) || 0;
      c.kg += Number(e.kg_applied) || 0;
      c.ha += Number(e.total_ha) || 0;
      c.hours += hrs;
      c.unit_hours += hrs * (Number(e.men) || 0);
      c.events += 1;
      c.cost += eventCost_(e, rateAt(e.date));
      c.mat_cost += foliarCost_(e, rateAt(e.date));
    });
    // green leaf plucked, by the method of the part of the field it came from (a field with one method: that method; a Mixed field: its drone part and its manual part)
    dbAll_(SHEETS.LEAF).forEach(function (r) {
      if (r.status !== 'Active' || METHODS.indexOf(r.method) < 0 || !isDate_(r.date)) return;
      const d = byDay[r.date] || (byDay[r.date] = { date: r.date, Drone: expCell_(), Manual: expCell_() }),
        c = d[r.method];
      c.leaf_kg += Number(r.kg) || 0;
      c.leaf_ha += Number(r.area_ha) || 0;
      c.leaf_n += 1;
      c.leaf_rev += (Number(r.kg) || 0) * rateAt(r.date).gl;
    });
    let firstSpray = '',
      firstLeaf = '';
    const days = Object.keys(byDay)
      .sort()
      .map(function (k) {
        const d = byDay[k];
        METHODS.forEach(function (m) {
          const c = d[m];
          c.kg = round_(c.kg, 3);
          c.ha = round_(c.ha, 3);
          c.hours = round_(c.hours, 3);
          c.unit_hours = round_(c.unit_hours, 3);
          c.cost = round_(c.cost, 2);
          c.mat_cost = round_(c.mat_cost, 2);
          c.leaf_kg = round_(c.leaf_kg, 3);
          c.leaf_ha = round_(c.leaf_ha, 3);
          c.leaf_rev = round_(c.leaf_rev, 2);
          if (c.events && !firstSpray) firstSpray = k;
          if (c.leaf_n && !firstLeaf) firstLeaf = k;
        });
        return d;
      });
    const now = rateAt(todayStr_()); // are the rates entered? (the amounts themselves are not sent)
    // Net margin: the land (ha) each method covers now (a mixed field counts its two parts), and the first day from which spraying AND plucking are both on record
    const land = { Drone: 0, Manual: 0 };
    fieldsAll_().forEach(function (f) {
      if (!asBool_(f.active)) return;
      METHODS.forEach(function (m) {
        land[m] += partHa_(f, m);
      });
    });
    METHODS.forEach(function (m) {
      land[m] = round_(land[m], 3);
    });
    // Foliar saving works per PLOT (the drone part or the manual part of one field): the lines of the active events (date, field, method, that field's kg, that field's ha, that field's foliar cost) and the area of each plot
    const spray = [],
      plots = {};
    dbAll_(SHEETS.SPRAY).forEach(function (l) {
      const m = activeEv[l.event_id];
      if (l.status !== 'Active' || !m || !isDate_(l.date)) return;
      const r = rateAt(l.date),
        kg = Number(l.kg_applied) || 0,
        lab = eventCost_(
          { method: m, hours: Number(l.hours) || 0, men: Number(l.units) || 0, drones: Number(l.units) || 0, total_ha: Number(l.area_ha) || 0 },
          r
        ); // a line is its field's share of the event: its hours, its crew (men / drones) and its ha
      spray.push([l.date, l.field_id, m, round_(kg, 3), round_(Number(l.area_ha) || 0, 3), round_(foliarCost_(l, r), 4), round_(lab, 4)]); // the 6th: the foliar cost of the line = the kg of each of its four compounds x the price of that compound on that day (an older line: its kg x the single price); the 7th: its labour cost (drone: ha x the drone charge; manual: hours x men x the man-hour rate)
    });
    // the same for the green leaf: one row per active plucking record (date, field, method of the part plucked, kg, ha plucked, revenue = kg x the green leaf price per kg of that day)
    const leaf = [];
    dbAll_(SHEETS.LEAF).forEach(function (r) {
      if (r.status !== 'Active' || METHODS.indexOf(r.method) < 0 || !isDate_(r.date)) return;
      const kg = Number(r.kg) || 0;
      leaf.push([r.date, r.field_id, r.method, round_(kg, 3), round_(Number(r.area_ha) || 0, 3), round_(kg * rateAt(r.date).gl, 4)]);
    });
    fieldsAll_().forEach(function (f) {
      METHODS.forEach(function (m) {
        const ha = partHa_(f, m);
        if (ha > 0) plots[f.field_id + '|' + m] = { l: fieldLabel_(f), ha: round_(ha, 3) };
      });
    });
    const firstNet = firstSpray && firstLeaf ? (firstSpray > firstLeaf ? firstSpray : firstLeaf) : '';
    return {
      today: todayStr_(),
      first: days.length ? days[0].date : '',
      first_spray: firstSpray,
      first_leaf: firstLeaf,
      first_net: firstNet,
      land: land,
      spray: spray,
      leaf: leaf,
      plots: plots,
      rates_set: {
        worker: now.worker > 0,
        drone: now.basis === 'per_ha' ? now.drone_ha > 0 : now.drone_hr > 0,
        material: spraysSeen
          ? unpricedSeen === 0
          : now.price > 0 ||
            (now.compound.Zinc > 0 && now.compound['Epsom salt'] > 0 && now.compound.Wuxal > 0 && (now.compound.Urea > 0 || now.compound.T65 > 0)),
        leaf: now.gl > 0
      },
      days: days
    };
  });
}

// ---- Params.gs
/** Admin parameters (cost rates, windows, validation limits). Rates carry an effective-from date. */

function paramDef_(key) {
  return PARAMS.filter(function (p) {
    return p.key === key;
  })[0];
}
function parseParam_(def, raw) {
  if (raw === '' || raw === null || raw === undefined) return def.def;
  if (def.type === 'number') {
    const n = Number(raw);
    return isFinite(n) ? n : def.def;
  }
  if (def.type === 'bool') return asBool_(raw);
  return String(raw);
}

/** All parameter values as of a date (defaults if never set). */
function getParams_(onDate) {
  onDate = onDate || todayStr_();
  const rows = dbAll_(SHEETS.PARAMS);
  const out = {};
  PARAMS.forEach(function (def) {
    const mine = rows.filter(function (r) {
      return r.key === def.key;
    });
    let pick = null;
    if (def.dated) {
      mine.forEach(function (r) {
        if (r.effective_from <= onDate && (!pick || r.effective_from >= pick.effective_from)) pick = r;
      });
    } else if (mine.length) {
      pick = mine[mine.length - 1];
    }
    out[def.key] = pick ? parseParam_(def, pick.value) : def.def;
  });
  return out;
}

function validateParamValue_(def, value) {
  if (def.type === 'number') {
    if (value === '' || value === null || value === undefined) {
      if (def.def === '') return '';
      fail_(def.label + ' is required.', 'VALIDATION');
    }
    return String(cleanNum_(value, def.label, { min: def.min, max: def.max }));
  }
  if (def.type === 'bool') return String(value === true || String(value).toLowerCase() === 'true');
  if (def.type === 'enum') {
    if (def.options.indexOf(String(value)) < 0) fail_(def.label + ' must be one of: ' + def.options.join(', ') + '.', 'VALIDATION');
    return String(value);
  }
  if (def.type === 'date') {
    if (value === '' || value === null || value === undefined) return '';
    if (!isDate_(String(value))) fail_(def.label + ' must be a date (YYYY-MM-DD).', 'VALIDATION');
    return String(value);
  }
  return cleanText_(value, 100, def.label, false);
}

function api_listParams(token) {
  return wrap_(function () {
    authz_(token, ['params.manage']);
    const rows = dbAll_(SHEETS.PARAMS);
    const today = todayStr_();
    const cur = getParams_(today);
    return PARAMS.map(function (def) {
      const mine = rows.filter(function (r) {
        return r.key === def.key;
      });
      const last = mine.length ? mine[mine.length - 1] : null;
      return {
        key: def.key,
        label: def.label,
        group: def.group,
        type: def.type,
        unit: def.unit || '',
        options: def.options || null,
        help: def.help || '',
        dated: !!def.dated,
        min: def.min,
        value: cur[def.key],
        is_default: !mine.length,
        effective_from: last ? last.effective_from : '',
        history_count: mine.length
      };
    });
  });
}

function api_setParam(token, p) {
  return wrap_(function () {
    const c = authz_(token, ['params.manage']);
    p = p || {};
    const def = paramDef_(p.key);
    if (!def) fail_('Unknown parameter.', 'VALIDATION');
    const value = validateParamValue_(def, p.value);
    let eff = '1900-01-01';
    if (def.dated) {
      eff = String(p.effective_from || todayStr_());
      if (!isDate_(eff)) fail_('Effective-from must be a date (YYYY-MM-DD).', 'VALIDATION');
    }
    const note = cleanText_(p.note, 200, 'Note', false);
    return withLock_(function () {
      const rows = dbAll_(SHEETS.PARAMS).filter(function (r) {
        return r.key === def.key;
      });
      const before = rows.length ? rows[rows.length - 1].value : '(default)';
      dbAppend_(SHEETS.PARAMS, [
        {
          param_id: newId_('PRM'),
          key: def.key,
          value: value,
          effective_from: eff,
          note: note,
          set_by: c.user.username,
          set_at: nowIso_()
        }
      ]);
      audit_(c.user.username, 'PARAM_SET', 'parameter', def.key, { from: before, to: value, effective_from: eff });
      return true;
    });
  });
}

function api_paramHistory(token, key) {
  return wrap_(function () {
    authz_(token, ['params.manage']);
    if (!paramDef_(key)) fail_('Unknown parameter.', 'VALIDATION');
    return dbAll_(SHEETS.PARAMS)
      .filter(function (r) {
        return r.key === key;
      })
      .map(function (r) {
        return { value: r.value, effective_from: r.effective_from, note: r.note, set_by: r.set_by, set_at: r.set_at };
      })
      .reverse();
  });
}

// ---- Records.gs
/** Foliar spray records and plucking records: entry (with rounds), replace on resubmit, day view, void, admin list and edit.
 *
 *  ROUNDS: for each field the records are walked in date order. The area of each record adds up; when the running total reaches the
 *  field area (within ROUND_TOL ha) the round is COMPLETE and the next record starts a new round. This gives, per field:
 *    days to complete a round = (completion date - first day of the round) + 1
 *    round gap                = first day of this round - first day of the previous round
 *  Each entry stores only that day's area; the slider in the app shows the cumulative position (earlier days frozen).
 *  A round left UNFINISHED for more than the window (spray_round_reset_days / pluck_round_reset_days: the days between two entries; Parameters) is closed as it is (expired) and the
 *  next record starts a NEW round, so the entries of one round can always be told apart. See resetDays_.
 *
 *  A field sprayed partly by drone and partly by hand (class Mixed) is plucked in TWO PARTS, entered side by side: the part sprayed by drone and
 *  the part sprayed by hand. Each part has its own record (the record's method is Drone or Manual), its own kg and its own area, and cannot
 *  exceed its part of the field. Each part has its OWN ROUNDS (the drone part may be on round 3 while the manual part is on round 2). A field with
 *  one method has one record per day and one set of rounds, as before.
 */

const AREA_TOL = 0.006; // rounding slack when comparing areas (ha)
const ROUND_TOL = 0.05; // a round counts as complete when this close to the whole field (ha)

function lockCheck_(date, params) {
  const lk = params.lock_before_date;
  if (lk && date <= lk) fail_('Records dated on or before ' + lk + ' are locked.', 'LOCKED');
}
/** People enter data for TODAY only. Entries made offline may still upload for sync_grace_days afterwards. */
function entryDateCheck_(date, params) {
  if (!isDate_(date)) fail_('Choose a valid date.', 'VALIDATION');
  const today = todayStr_();
  if (date > today) fail_('The date cannot be in the future.', 'VALIDATION');
  lockCheck_(date, params);
  if (date < addDays_(today, -params.sync_grace_days)) fail_('Past days are view only. Data can only be entered or changed for today.', 'VALIDATION');
}
function recSheet_(type) {
  return type === 'spray' ? SHEETS.SPRAY : type === 'pluck' ? SHEETS.LEAF : '';
}
function entryTask_(type) {
  return type === 'spray' ? 'spray.entry' : 'pluck.entry';
}
function typeOrFail_(type) {
  if (!recSheet_(type)) fail_('Invalid record type.', 'VALIDATION');
  return type;
}

/** A field this person may record on: in scope, active and fully set up. */
function usableField_(c, f) {
  return !!f && fieldComplete_(f) && inScope_(c, f);
}
function fieldOrFail_(c, idx, id) {
  const f = idx[id];
  if (!f || !f.active) fail_('A selected field was not found.', 'VALIDATION');
  if (!inScope_(c, f)) fail_('Field ' + fieldLabel_(f) + ' is not one of your fields.', 'FORBIDDEN');
  if (!fieldComplete_(f))
    fail_('Field ' + fieldLabel_(f) + ' is not fully set up yet (area, last prune date, tea type and drone area). Ask the Admin.', 'VALIDATION');
  return f;
}

// ---------------------------------------------------------------- rounds
/** How many days a round may wait unfinished before the next record starts a new round (the settings spray_round_reset_days and pluck_round_reset_days). */
function resetDays_(type, params) {
  params = params || getParams_(todayStr_());
  return type === 'spray' ? params.spray_round_reset_days : params.pluck_round_reset_days;
}
/** resetDays: a record more than that many days after the last record of an unfinished round starts a new round; the unfinished one stays as it was (expired). */
function buildRounds_(recs, fieldHa, resetDays) {
  const list = recs.slice().sort(function (a, b) {
    return a.date < b.date
      ? -1
      : a.date > b.date
        ? 1
        : a.entered_at < b.entered_at
          ? -1
          : a.entered_at > b.entered_at
            ? 1
            : (a._ord || 0) - (b._ord || 0);
  });
  const rounds = [];
  let cur = null;
  list.forEach(function (r) {
    if (cur && resetDays > 0 && daysBetween_(cur.last, r.date) > resetDays) {
      cur.expired = true;
      cur.days = daysBetween_(cur.start, cur.last) + 1;
      rounds.push(cur);
      cur = null;
    }
    if (!cur)
      cur = {
        no: rounds.length + 1,
        start: r.date,
        end: '',
        done_ha: 0,
        by: { Drone: 0, Manual: 0 },
        days: 1,
        complete: false,
        expired: false,
        records: [],
        last: r.date
      };
    cur.done_ha = round_(cur.done_ha + r.area_ha, 3);
    cur.records.push(r);
    cur.last = r.date;
    if (cur.by[r.method] !== undefined) cur.by[r.method] = round_(cur.by[r.method] + r.area_ha, 3); // foliar: how much of the round each method has covered
    if (cur.done_ha >= fieldHa - ROUND_TOL) {
      cur.complete = true;
      cur.end = r.date;
      cur.days = daysBetween_(cur.start, r.date) + 1;
      rounds.push(cur);
      cur = null;
    }
  });
  if (cur) {
    cur.days = daysBetween_(cur.start, cur.last) + 1;
    rounds.push(cur);
  }
  return rounds;
}
/** The records of one PART of a field. A field sprayed or plucked partly by drone and partly by hand (class Mixed) has two parts, each with its OWN rounds:
 *  the drone part can be on its second round while the manual part is on its first. A field with one method has one part (all its records). */
function partRecs_(f, method, recs) {
  return f.method === 'Mixed'
    ? recs.filter(function (r) {
        return r.method === method;
      })
    : recs;
}
/** The methods that have a part on a field: [Drone], [Manual] or both. */
function partMethods_(f) {
  return f.method === 'Mixed' ? METHODS.slice() : METHODS.indexOf(f.method) >= 0 ? [f.method] : [];
}
/** The ha of a part of a field (a field with one method: the whole field). */
function partExt_(f, method) {
  return f.method === 'Mixed' ? partHa_(f, method) : f.method === method ? f.area_ha : 0;
}
/** One part's open (unfinished) round as of a date, from records dated BEFORE that date (foliar: all lines given). Its rounds are its own: a part's round
 *  is complete when its own ha are covered; after that, or after the waiting limit, the next record starts a new round. */
function partProgressAt_(recs, f, method, date, resetDays) {
  const ext = partExt_(f, method),
    rounds = buildRounds_(partRecs_(f, method, recs), ext, resetDays),
    last = rounds[rounds.length - 1];
  const stale = !!last && !last.complete && resetDays > 0 && daysBetween_(last.last, date) > resetDays; // left unfinished too long: the card starts again
  const open = last && !last.complete && !last.expired && !stale ? last : null;
  const prev = open ? rounds[rounds.length - 2] : last,
    done = open ? Math.min(open.done_ha, ext) : 0;
  return {
    ext: ext,
    done: done,
    remaining: Math.max(ext - done, 0),
    done_pct: ext > 0 ? (done / ext) * 100 : 0,
    round: open ? open.no : rounds.length + 1,
    started: open ? open.start : '',
    day_no: open ? daysBetween_(open.start, date) + 1 : 1,
    prev_start: prev ? prev.start : '',
    prev_end: prev && prev.complete ? prev.end : '',
    done_today: !!last && last.complete && last.end === date
  };
}
/** The open round(s) of a field as of a date: `parts` has one entry per method that has a part; the rest is a summary of the whole field (for a field
 *  with one method it is that part). */
function progressAt_(recs, f, date, resetDays) {
  const parts = {};
  let done = 0,
    round = 1,
    started = '',
    prevStart = '',
    prevEnd = '';
  partMethods_(f).forEach(function (m, i) {
    const pp = partProgressAt_(recs, f, m, date, resetDays);
    parts[m] = pp;
    done += pp.done;
    if (pp.round > round) round = pp.round;
    if (pp.started && (!started || pp.started < started)) started = pp.started;
    if (i === 0) {
      prevStart = pp.prev_start;
      prevEnd = pp.prev_end;
    }
  });
  return {
    round: round,
    done_ha: round_(done, 2),
    done_exact: done,
    parts: parts,
    remaining_ha: round_(Math.max(f.area_ha - done, 0), 2),
    done_pct: f.area_ha ? round_((done / f.area_ha) * 100, 1) : 0,
    started: started,
    day_no: started ? daysBetween_(started, date) + 1 : 1,
    prev_start: prevStart,
    prev_end: prevEnd
  };
}
/** Area a line adds today. "complete" (the 100% button) means: everything left in the round. */
function areaFor_(line, f, doneBefore, label) {
  const remaining = round_(Math.max(f.area_ha - doneBefore, 0), 3);
  if (remaining <= 0) fail_('This round is already complete.', 'VALIDATION');
  if (line.complete === true || line.complete === 'true') return remaining;
  const area = cleanNum_(line.area_ha, label, { required: true, gt: 0 });
  if (area > remaining + AREA_TOL)
    fail_(label + ' (' + area + ' ha) is more than the area left in this round (' + round_(remaining, 2) + ' ha).', 'VALIDATION');
  return Math.min(area, remaining);
}
function recordsBefore_(type, date) {
  const m = {};
  dbAll_(recSheet_(type)).forEach(function (r) {
    if (r.status === 'Active' && r.date < date) (m[r.field_id] = m[r.field_id] || []).push(r);
  });
  return m;
}

/** Problems found when a field's records (spray lines or plucking records) are replayed in order. Returns {key: message}; a problem that already existed before a change is not new. */
function replayProblems_(type, recs, idx, fieldIds) {
  const out = {},
    by = groupBy_(recs, function (r) {
      return r.field_id;
    }),
    reset = resetDays_(type),
    spray = type === 'spray';
  fieldIds.forEach(function (fid) {
    const f = idx[fid];
    if (!f) return;
    const label = fieldLabel_(f);
    partMethods_(f).forEach(function (m) {
      const ext = partExt_(f, m),
        rounds = buildRounds_(partRecs_(f, m, by[fid] || []), ext, reset);
      rounds.forEach(function (rd, i) {
        const over = rd.done_ha - ext;
        if (over > ROUND_TOL)
          out[fid + ':cap:' + m + ':' + rd.start] =
            label +
            ': with this change the ' +
            (spray ? 'spray events' : 'plucking records') +
            ' after it would cover ' +
            round_(over, 2) +
            ' ha more than the ' +
            m.toLowerCase() +
            ' part of the field. Lower the ' +
            (spray ? 'coverage' : 'area') +
            ' here, or change the later ' +
            (spray ? 'event' : 'record') +
            ' first.';
        if (i > 0 && rounds[i - 1].complete && rd.start <= rounds[i - 1].end)
          out[fid + ':twice:' + m + ':' + rd.start] =
            label +
            ': with this change the ' +
            m.toLowerCase() +
            ' part would be finished and then ' +
            (spray ? 'sprayed' : 'plucked') +
            ' again on the same day. Change the later ' +
            (spray ? 'event' : 'record') +
            ' as well.';
      });
    });
  });
  return out;
}
/** Refuse a change (records before -> records after) that makes a field's records impossible; problems that were already there are not blamed on it. */
function checkReplay_(type, before, after, idx, fieldIds) {
  const had = replayProblems_(type, before, idx, fieldIds),
    now = replayProblems_(type, after, idx, fieldIds);
  const fresh = Object.keys(now).filter(function (k) {
    return !had[k];
  });
  if (fresh.length) fail_(now[fresh[0]], 'VALIDATION');
}

// ---------------------------------------------------------------- field pickers (what the person sees)
function api_entryFields(token, type, date) {
  return wrap_(function () {
    typeOrFail_(type);
    const c = authz_(token, [entryTask_(type)]);
    if (!isDate_(date)) fail_('Choose a valid date.', 'VALIDATION');
    const idx = fieldIndex_(),
      reset = resetDays_(type);
    const rows = dbAll_(recSheet_(type)).filter(function (r) {
      return r.status === 'Active';
    });
    const done = {},
      last = {},
      before = {};
    rows.forEach(function (r) {
      if (r.date === date) done[r.field_id] = true;
      if (r.date <= date && (!last[r.field_id] || r.date > last[r.field_id])) last[r.field_id] = r.date;
      if (r.date < date) (before[r.field_id] = before[r.field_id] || []).push(r);
    });
    const out = Object.keys(idx)
      .map(function (k) {
        return idx[k];
      })
      .filter(function (f) {
        return usableField_(c, f);
      })
      .sort(function (a, b) {
        return a.division_seq === b.division_seq ? a.seq - b.seq : a.division_seq - b.division_seq;
      })
      .map(function (f) {
        const p = progressAt_(before[f.field_id] || [], f, date, reset),
          parts = type === 'pluck' ? partsOf_(f, p) : undefined;
        delete p.done_exact;
        delete p.parts;
        return {
          field_id: f.field_id,
          code: f.code,
          name: f.name,
          division: f.division,
          area_ha: f.area_ha,
          method: f.method,
          drone_ha: f.drone_ha,
          manual_ha: f.manual_ha,
          parts: parts,
          tea_type: f.tea_type,
          recorded: !!done[f.field_id],
          last_date: last[f.field_id] || '',
          progress: p
        };
      });
    return { fields: out, today: todayStr_(), editable_from: addDays_(todayStr_(), -getParams_(todayStr_()).sync_grace_days) };
  });
}

// ---------------------------------------------------------------- saving (new entry, or REPLACE the record for that field and day)
function pluckLine_(line, f, params, doneBefore) {
  const kg = cleanNum_(line.kg, 'Green leaf weight', { required: true, gt: 0, max: 1000000 });
  const area = areaFor_(line, f, doneBefore, 'Plucked area');
  const flags = [];
  const mx = params.leaf_kg_per_ha_max;
  if (mx !== '' && mx !== undefined && kg / area > mx) flags.push('KG/HA HIGH (' + round_(kg / area, 1) + ')');
  return { kg: round_(kg, 3), area_ha: round_(area, 3), flags: flags.join('; ') };
}

/** The lines of one field in an entry: one for a field with one method; for a Mixed field a `parts` list, one item per part (Drone / Manual). */
function plotItems_(line, f) {
  if (f.method !== 'Mixed') return [{ part: f.method, line: line, key: f.field_id }];
  const parts = Array.isArray(line.parts) ? line.parts : [];
  if (!parts.length) fail_('This field is sprayed partly by drone and partly by hand: enter its drone part and its manual part.', 'VALIDATION');
  return parts.map(function (pt) {
    if (!pt || METHODS.indexOf(pt.part) < 0 || partHa_(f, pt.part) <= 0)
      fail_('This field has no ' + String(pt && pt.part).toLowerCase() + ' part.', 'VALIDATION');
    return { part: pt.part, line: pt, key: f.field_id + '|' + pt.part };
  });
}
/** The key a stored record has within its field: a Mixed field has one record per part. */
function recKey_(f, r) {
  return f && f.method === 'Mixed' ? r.field_id + '|' + r.method : r.field_id;
}

function savePluckLines_(c, p) {
  p = p || {};
  const lines = Array.isArray(p.lines) ? p.lines : [];
  if (!lines.length) fail_('Select at least one field.', 'VALIDATION');
  if (lines.length > 100) fail_('Too many fields in one entry.', 'VALIDATION');
  const sheet = recSheet_('pluck');
  return withLock_(function () {
    const params = getParams_(todayStr_());
    entryDateCheck_(p.date, params);
    const idx = fieldIndex_();
    const existing = {},
      before = {},
      active = dbAll_(sheet).filter(function (r) {
        return r.status === 'Active';
      });
    active.forEach(function (r) {
      if (r.date === p.date) existing[recKey_(idx[r.field_id], r)] = r;
      else if (r.date < p.date) (before[r.field_id] = before[r.field_id] || []).push(r);
    });
    const seen = {},
      touched = {},
      inserts = [],
      updates = [],
      warnings = [],
      now = nowIso_();
    lines.forEach(function (line) {
      const f = fieldOrFail_(c, idx, line.field_id);
      let part = '';
      touched[f.field_id] = true;
      try {
        const prog = progressAt_(before[f.field_id] || [], f, p.date, resetDays_('pluck', params));
        plotItems_(line, f).forEach(function (it) {
          part = f.method === 'Mixed' ? ' (' + it.part.toLowerCase() + ' part)' : '';
          if (seen[it.key]) fail_('appears twice in this entry.', 'VALIDATION');
          seen[it.key] = true;
          const pp = partProgress_(f, it.part, prog),
            ext = pp.ext,
            doneBefore = pp.done;
          const vals = pluckLine_(it.line, { area_ha: ext }, params, doneBefore);
          const prev = existing[it.key];
          if (prev) updates.push({ prev: prev, vals: vals, f: f });
          else
            inserts.push(
              Object.assign(
                {
                  record_id: newId_('PLK'),
                  date: p.date,
                  field_id: f.field_id,
                  method: it.part,
                  prune_date: f.last_prune_date,
                  status: 'Active',
                  entered_by: c.user.username,
                  entered_at: now
                },
                vals
              )
            );
          if (vals.flags) warnings.push(fieldLabel_(f) + part + ': ' + vals.flags);
        });
      } catch (e) {
        if (e instanceof AppError) throw new AppError('Field ' + fieldLabel_(f) + part + ': ' + e.message, e.code);
        throw e;
      }
    });
    // a day saved late (or saved again) must not make the days AFTER it impossible
    const newVals = {};
    updates.forEach(function (u) {
      newVals[u.prev.record_id] = u.vals;
    });
    checkReplay_(
      'pluck',
      active,
      active
        .map(function (r) {
          return newVals[r.record_id] ? Object.assign({}, r, newVals[r.record_id]) : r;
        })
        .concat(inserts),
      idx,
      Object.keys(touched)
    );
    updates.forEach(function (u) {
      const was = { kg: u.prev.kg, area: u.prev.area_ha };
      dbUpdate_(sheet, u.prev._row, Object.assign({}, u.vals, { edited_by: c.user.username, edited_at: now }));
      audit_(c.user.username, 'RECORD_REPLACE', 'pluck', u.prev.record_id, {
        date: p.date,
        field: u.f.code,
        part: u.prev.method,
        before: was,
        after: u.vals
      });
    });
    dbAppend_(sheet, inserts);
    audit_(c.user.username, 'PLUCK_SAVE', 'pluck', '', { date: p.date, new: inserts.length, replaced: updates.length });
    return { saved: inserts.length + updates.length, replaced: updates.length, warnings: warnings };
  });
}

function api_savePluck(token, p) {
  return wrap_(function () {
    return savePluckLines_(authz_(token, ['pluck.entry']), p);
  });
}

// ---------------------------------------------------------------- day view, void, admin list/edit
function canVoid_(c, r, f, params, today) {
  if (r.status !== 'Active') return false;
  if (params.lock_before_date && r.date <= params.lock_before_date) return false;
  if (hasTask_(c, 'records.manage')) return f ? inScope_(c, f) : c.isAdmin; // (the Records power only reaches the fields this account may see)
  if (r.start_time !== undefined) return false; // spray lines belong to an event: delete the event from the Foliar spray tab
  if (!hasTask_(c, entryTask_('pluck')) || !f || !inScope_(c, f) || !ownsName_(c.user, r.entered_by)) return false;
  return r.date >= addDays_(today, -params.sync_grace_days) && daysBetween_(String(r.entered_at).slice(0, 10), today) <= params.edit_window_days;
}

function shapeRecord_(type, r, idx, c, params, today) {
  const f = idx[r.field_id];
  const o = {
    record_id: r.record_id,
    date: r.date,
    field_id: r.field_id,
    field_code: f ? f.code : '?',
    field_name: f ? f.name : '',
    division: f ? f.division : '',
    method: r.method,
    status: r.status,
    void_reason: r.void_reason,
    flags: r.flags,
    entered_by: r.entered_by,
    entered_at: r.entered_at,
    edited_by: r.edited_by,
    edited_at: r.edited_at,
    area_ha: r.area_ha
  };
  if (type === 'spray') {
    o.start_time = r.start_time;
    o.end_time = r.end_time;
    o.hours = r.hours;
    o.units = r.units;
    o.event_id = r.event_id;
    o.kg_applied = r.kg_applied;
    o.pct_total = r.pct_total;
  } else {
    o.kg = r.kg;
    o.kg_per_ha = r.area_ha ? round_(r.kg / r.area_ha, 1) : '';
  }
  o.field_area_ha = f ? f.area_ha : '';
  o.can_void = canVoid_(c, r, f, params, today);
  o.can_edit =
    type === 'pluck' && r.status === 'Active' && hasTask_(c, 'records.manage') && !(params.lock_before_date && r.date <= params.lock_before_date);
  return o;
}

/** Records for one date on the fields this person may see, with the round each one belongs to. */
function api_dayRecords(token, type, date) {
  return wrap_(function () {
    typeOrFail_(type);
    const c = authz_(token, [entryTask_(type)]);
    if (!isDate_(date)) fail_('Choose a valid date.', 'VALIDATION');
    const idx = fieldIndex_(),
      params = getParams_(todayStr_()),
      today = todayStr_();
    const all = dbAll_(recSheet_(type)).filter(function (r) {
      return r.status === 'Active';
    });
    const day = all.filter(function (r) {
      return r.date === date && idx[r.field_id] && inScope_(c, idx[r.field_id]);
    });
    const byField = {};
    all.forEach(function (r) {
      (byField[r.field_id] = byField[r.field_id] || []).push(r);
    });
    const roundsOf = {};
    return day.map(function (r) {
      const o = shapeRecord_(type, r, idx, c, params, today);
      const f = idx[r.field_id],
        rk = recKey_(f, r);
      const rounds =
        roundsOf[rk] ||
        (roundsOf[rk] = buildRounds_(
          partRecs_(f, r.method, byField[r.field_id]),
          partExt_(f, f.method === 'Mixed' ? r.method : f.method),
          resetDays_(type, params)
        ));
      rounds.forEach(function (rd) {
        if (
          rd.records.some(function (x) {
            return x.record_id === r.record_id;
          })
        )
          o.round = { no: rd.no, day: daysBetween_(rd.start, r.date) + 1, complete: rd.complete && rd.end === r.date, start: rd.start };
      });
      return o;
    });
  });
}

function api_listRecords(token, q) {
  return wrap_(function () {
    const c = authz_(token, ['records.manage']);
    q = q || {};
    const type = q.type === 'pluck' ? 'pluck' : 'spray';
    const params = getParams_(todayStr_()),
      today = todayStr_(),
      idx = fieldIndex_();
    let rows = dbAll_(recSheet_(type)).filter(function (r) {
      const f = idx[r.field_id];
      return f ? inScope_(c, f) : c.isAdmin;
    });
    if (q.from && isDate_(q.from))
      rows = rows.filter(function (r) {
        return r.date >= q.from;
      });
    if (q.to && isDate_(q.to))
      rows = rows.filter(function (r) {
        return r.date <= q.to;
      });
    if (q.field_id)
      rows = rows.filter(function (r) {
        return r.field_id === q.field_id;
      });
    if (FIELD_METHODS.indexOf(q.method) >= 0)
      rows = rows.filter(function (r) {
        return r.method === q.method;
      });
    if (q.status === 'Active' || q.status === 'Voided')
      rows = rows.filter(function (r) {
        return r.status === q.status;
      });
    rows.sort(function (a, b) {
      return a.date === b.date ? (a.entered_at < b.entered_at ? 1 : -1) : a.date < b.date ? 1 : -1;
    });
    return {
      total: rows.length,
      rows: rows.slice(0, 500).map(function (r) {
        return shapeRecord_(type, r, idx, c, params, today);
      })
    };
  });
}

function findRecord_(type, id) {
  const rec = dbAll_(recSheet_(typeOrFail_(type))).filter(function (r) {
    return r.record_id === id;
  })[0];
  if (!rec) fail_('Record not found.', 'VALIDATION');
  return rec;
}

function api_voidRecord(token, type, id, reason) {
  return wrap_(function () {
    const c = authz_(token, null);
    typeOrFail_(type);
    const why = cleanText_(reason, 200, 'Reason', true);
    if (why.length < 3) fail_('Please give a reason.', 'VALIDATION');
    return withLock_(function () {
      const params = getParams_(todayStr_());
      const rec = findRecord_(type, id);
      if (type === 'spray' && rec.event_id) {
        // a spray record is one field of an event: the whole event goes
        if (!hasTask_(c, 'records.manage')) fail_('Delete spray events from the Foliar spray tab.', 'FORBIDDEN');
        return deleteEvent_(c, rec.event_id, why, true);
      }
      const idx = fieldIndex_();
      if (!canVoid_(c, rec, idx[rec.field_id], params, todayStr_()))
        fail_('You cannot void this record (not yours, past days are view only, locked, or already voided).', 'FORBIDDEN');
      if (type === 'pluck') {
        const active = dbAll_(recSheet_('pluck')).filter(function (r) {
          return r.status === 'Active';
        });
        checkReplay_(
          'pluck',
          active,
          active.filter(function (r) {
            return r.record_id !== id;
          }),
          idx,
          [rec.field_id]
        );
      }
      dbUpdate_(recSheet_(type), rec._row, { status: 'Voided', void_reason: why, edited_by: c.user.username, edited_at: nowIso_() });
      audit_(c.user.username, 'RECORD_VOID', type, id, { reason: why, date: rec.date });
      return true;
    });
  });
}

/** Admin edit of the measurements (not date or field: void and re-enter for those). */
function api_editRecord(token, type, id, patch) {
  return wrap_(function () {
    const c = authz_(token, ['records.manage']);
    typeOrFail_(type);
    patch = patch || {};
    if (type === 'spray')
      fail_(
        'A spray record is one field of a spray event. Open the event on the Foliar spray tab to change it, or void the event here and enter it again.',
        'VALIDATION'
      );
    return withLock_(function () {
      const params = getParams_(todayStr_());
      const rec = findRecord_(type, id);
      if (rec.status !== 'Active') fail_('Voided records cannot be edited.', 'VALIDATION');
      lockCheck_(rec.date, params);
      const f = fieldIndex_()[rec.field_id];
      if (!f) fail_('Field not found.', 'VALIDATION');
      if (!inScope_(c, f)) fail_('This record is on a field that is not yours.', 'FORBIDDEN');
      const prog = progressAt_(recordsBefore_(type, rec.date)[rec.field_id] || [], f, rec.date, resetDays_(type, params)),
        part = f.method === 'Mixed' && METHODS.indexOf(rec.method) >= 0 ? rec.method : f.method,
        pp = partProgress_(f, part, prog);
      const fieldLike = { method: rec.method, area_ha: pp.ext };
      const doneBefore = pp.done;
      const before = { kg: rec.kg, area: rec.area_ha };
      const upd = pluckLine_({ kg: patch.kg, area_ha: patch.area_ha }, fieldLike, params, doneBefore);
      upd.edited_by = c.user.username;
      upd.edited_at = nowIso_();
      const active = dbAll_(recSheet_(type)).filter(function (r) {
        return r.status === 'Active';
      });
      checkReplay_(
        type,
        active,
        active.map(function (r) {
          return r.record_id === id ? Object.assign({}, r, upd) : r;
        }),
        fieldIndex_(),
        [rec.field_id]
      );
      dbUpdate_(recSheet_(type), rec._row, upd);
      audit_(c.user.username, 'RECORD_EDIT', type, id, { before: before, after: upd });
      return true;
    });
  });
}

// ---- SprayEvents.gs
/** Foliar spray EVENTS.
 *
 *  Workers know how much foliar they MIXED and which fields they sprayed, not the kg per field. So foliar is recorded per event:
 *    SprayEvents   one row per event: kg applied, men, drones (Drone only), start and end time, total ha covered.
 *    SprayRecords  one row per field covered by the event (same event_id): the ha sprayed by THIS event, the cumulative % of the field
 *                  reached (pct_total), and that field's share of the event's kg and hours (shared out by hectares).
 *  A field's coverage adds up over events (and days) within a round, exactly like plucking: a round is complete when the whole field (for a
 *  Mixed field: the method's part) is covered (see Records.gs). The slider in the app shows the cumulative %, and its minimum is what earlier events
 *  already covered.
 *  An event is one method only (all Drone or all Manual).
 *  The foliar is a mix of four compounds (Zinc, Epsom salt, Wuxal and Urea or T65, see FoliarPlan.gs): an event stores the kg of each (kg_applied is their sum), and each field line its share
 *  of each, by hectares. An older event (before this) has one total kg and no compounds; it is still read and priced with the single foliar price.
 *  A round left unfinished for more than the window (Parameters: spray_round_reset_days) is closed as it is and the next
 *  spray event on that part starts a NEW round, so the card in the app starts again from 0% and the events of one round can always be told apart.
 *
 *  A field can be sprayed partly by drone and partly by hand (Fields.drone_ha; the rest is manual). Each method covers only ITS part of the
 *  field: in a Drone event the coverage % is a % of the field's drone part, in a Manual event a % of its manual part. The two parts have SEPARATE
 *  rounds: the drone part can be on its second round while the manual part is on its first. A part that was finished today cannot be sprayed again
 *  today; on a later day it starts a new round.
 *
 *  ANY event of a day can be changed or deleted, not only the last one. Each event stores the hectares it covered; a field's running
 *  % is worked out by replaying its events in order. So after a change the field's events are replayed and the change is refused only if
 *  that would leave the field in an impossible state (a later event then covers more than a method's part, or the field is finished and
 *  sprayed again on the same day). The stored running % of every later event is then recalculated. Deleting marks the event and its
 *  lines Voided (audit kept).
 */

const MIN_LINE_HA = 0.005;
function ordKey_(date, row) {
  return date + '#' + ('0000000' + row).slice(-7);
}
function evKey_(e) {
  return ordKey_(e.date, e._row);
}

/** Active spray lines, each tagged with the creation order of its event. Lines from before events existed come first on their day. */
function sprayLinesWithOrder_(events) {
  const rowOf = {};
  events.forEach(function (e) {
    rowOf[e.event_id] = e._row;
  });
  return dbAll_(SHEETS.SPRAY)
    .filter(function (r) {
      return r.status === 'Active';
    })
    .map(function (r) {
      r._ord = r.event_id ? rowOf[r.event_id] || 0 : 0;
      r._key = ordKey_(r.date, r._ord);
      return r;
    });
}
function groupBy_(list, keyFn) {
  const m = {};
  list.forEach(function (x) {
    const k = keyFn(x);
    (m[k] = m[k] || []).push(x);
  });
  return m;
}

/** After a change: the stored running % (pct_total) of every active line of these fields, worked out again by replaying them. pct_total is the % of the
 *  METHOD'S PART of the field covered so far in that part's round (for a field with one method: of the whole field). */
function recomputePct_(fieldIds, idx) {
  const events = dbAll_(SHEETS.SPRAY_EVENTS),
    by = groupBy_(sprayLinesWithOrder_(events), function (l) {
      return l.field_id;
    }),
    reset = resetDays_('spray');
  fieldIds.forEach(function (fid) {
    const f = idx[fid];
    if (!f || !(f.area_ha > 0)) return;
    partMethods_(f).forEach(function (m) {
      const ext = partExt_(f, m);
      if (!(ext > 0)) return;
      buildRounds_(partRecs_(f, m, by[fid] || []), ext, reset).forEach(function (rd) {
        let cum = 0;
        rd.records.forEach(function (l) {
          cum = round_(cum + l.area_ha, 3);
          const pct = rd.complete && ext - cum < ROUND_TOL ? 100 : round_((cum / ext) * 100, 1);
          if (l.pct_total !== pct) dbUpdate_(SHEETS.SPRAY, l._row, { pct_total: pct });
        });
      });
    });
  });
}
/** One method's part of a field in its open round: ha of the part, ha done, ha left, % done, its round number. prog comes from progressAt_. */
function partProgress_(f, method, prog) {
  const pp = prog && prog.parts && prog.parts[method];
  if (pp) return pp;
  return {
    ext: partExt_(f, method),
    done: 0,
    remaining: partExt_(f, method),
    done_pct: 0,
    round: 1,
    started: '',
    day_no: 1,
    prev_start: '',
    prev_end: '',
    done_today: false
  };
}
/** For the app: each method's part of a field and how much of it is done in ITS open round (prog from progressAt_). `done` = the part was finished today. */
function partsOf_(f, prog) {
  const parts = {};
  METHODS.forEach(function (m) {
    const pp = partProgress_(f, m, prog);
    parts[m] = {
      ha: round_(pp.ext, 3),
      done_ha: round_(pp.done, 3),
      remaining_ha: round_(pp.remaining, 3),
      done_pct: round_(pp.done_pct, 1),
      done: pp.ext > 0 && !!pp.done_today,
      round: pp.round,
      started: pp.started,
      day_no: pp.day_no
    };
  });
  return parts;
}
/** The methods this person may record: a pilot role records Drone events only. */
function allowedMethods_(c) {
  return !c.isAdmin && c.role.field_scope === 'drone' ? ['Drone'] : METHODS.slice();
}

function shapeEvent_(e, lines, no, c, idx, byField, params, today) {
  const inScopeAll =
    lines.length > 0 &&
    lines.every(function (l) {
      return idx[l.field_id] && inScope_(c, idx[l.field_id]);
    });
  const visible = lines.some(function (l) {
    return idx[l.field_id] && inScope_(c, idx[l.field_id]);
  });
  if (!visible) return null;
  const key = evKey_(e);
  const outLines = lines.map(function (l) {
    const f = idx[l.field_id];
    if (!f || !inScope_(c, f))
      return {
        field_id: '',
        code: '',
        label: 'Another field',
        division: '',
        field_ha: 0,
        part_ha: 0,
        mixed: false,
        area_ha: l.area_ha,
        pct_total: l.pct_total,
        pct_method: 0,
        before_pct: 0,
        max_pct: 0,
        kg: l.kg_applied,
        mix: mixOf_(l),
        hidden: true
      }; // not one of this person's fields (or a field this phone does not hold): no name, no code
    const others = (byField[l.field_id] || []).filter(function (x) {
      return x.event_id !== e.event_id;
    });
    const pp = f
      ? partProgress_(
          f,
          e.method,
          progressAt_(
            others.filter(function (x) {
              return x._key < key;
            }),
            f,
            e.date,
            params.spray_round_reset_days
          )
        )
      : { ext: 0, done: 0 };
    const after = pp.ext > 0 ? (pp.ext - pp.done - l.area_ha < ROUND_TOL ? 100 : round_(((pp.done + l.area_ha) / pp.ext) * 100, 1)) : 0;
    // later events of the same method in the same round already use part of the method's part: this line cannot be raised past what is left
    let laterHa = 0;
    if (f)
      buildRounds_(partRecs_(f, e.method, byField[l.field_id] || []), pp.ext, params.spray_round_reset_days).forEach(function (rd) {
        if (
          rd.records.some(function (x) {
            return x.event_id === e.event_id;
          })
        )
          rd.records.forEach(function (x) {
            if (x._key > key && x.method === e.method) laterHa += x.area_ha;
          });
      });
    const maxPct = pp.ext > 0 ? Math.max(0, Math.min(100, Math.floor(((pp.ext - laterHa) / pp.ext) * 100 + 1e-9))) : 0;
    return {
      field_id: l.field_id,
      code: f ? f.code : '?',
      label: f ? fieldLabel_(f) : '?',
      division: f ? f.division : '',
      field_ha: f ? f.area_ha : 0,
      part_ha: pp.ext,
      mixed: !!f && f.method === 'Mixed',
      area_ha: l.area_ha,
      pct_total: l.pct_total,
      pct_method: after,
      before_pct: pp.ext > 0 ? round_((pp.done / pp.ext) * 100, 1) : 0,
      max_pct: maxPct,
      kg: l.kg_applied,
      mix: mixOf_(l)
    };
  });
  const dateOk = !(params.lock_before_date && e.date <= params.lock_before_date) && e.date >= addDays_(today, -params.sync_grace_days);
  let reason = '';
  if (!dateOk) reason = 'Past days are view only.';
  else if (!inScopeAll) reason = 'It includes fields that are not yours.';
  else if (allowedMethods_(c).indexOf(e.method) < 0) reason = 'Your job role records Drone spray events only.';
  return {
    event_id: e.event_id,
    date: e.date,
    no: no,
    method: e.method,
    kg: e.kg_applied,
    men: e.men,
    drones: e.drones,
    start_time: e.start_time,
    end_time: e.end_time,
    hours: e.hours,
    total_ha: e.total_ha,
    kg_per_ha: e.total_ha > 0 ? round_(e.kg_applied / e.total_ha, 2) : '',
    mix: mixOf_(e),
    plan_id: e.plan_id,
    flags: e.flags,
    entered_by: e.entered_by,
    entered_at: e.entered_at,
    edited_by: e.edited_by,
    lines: outLines,
    can_edit: !reason,
    can_delete: !reason,
    lock_reason: reason
  };
}

/** What the Foliar spray screen needs for one day: the person's fields (with how much of the current round is already sprayed) and the day's events. */
function api_sprayDay(token, date) {
  return wrap_(function () {
    const c = authz_(token, ['spray.entry']);
    if (!isDate_(date)) fail_('Choose a valid date.', 'VALIDATION');
    const today = todayStr_(),
      params = getParams_(today),
      idx = fieldIndex_();
    const events = dbAll_(SHEETS.SPRAY_EVENTS).filter(function (e) {
      return e.status === 'Active';
    });
    const lines = sprayLinesWithOrder_(events),
      byField = groupBy_(lines, function (l) {
        return l.field_id;
      }),
      byEvent = groupBy_(
        lines.filter(function (l) {
          return l.event_id;
        }),
        function (l) {
          return l.event_id;
        }
      );
    const fields = Object.keys(idx)
      .map(function (k) {
        return idx[k];
      })
      .filter(function (f) {
        return usableField_(c, f);
      })
      .sort(function (a, b) {
        return a.division_seq === b.division_seq ? a.seq - b.seq : a.division_seq - b.division_seq;
      })
      .map(function (f) {
        const mine = (byField[f.field_id] || []).filter(function (x) {
            return x.date <= date;
          }),
          p = progressAt_(mine, f, date, params.spray_round_reset_days); // events after the chosen day do not count
        let lastDate = '';
        mine.forEach(function (x) {
          if (x.date > lastDate) lastDate = x.date;
        });
        const parts = partsOf_(f, p);
        return {
          field_id: f.field_id,
          code: f.code,
          name: f.name,
          division: f.division,
          area_ha: f.area_ha,
          method: f.method,
          drone_ha: f.drone_ha,
          manual_ha: f.manual_ha,
          parts: parts,
          tea_type: f.tea_type,
          last_date: lastDate,
          done_today: f.method !== 'Mixed' && !!parts[f.method] && parts[f.method].done,
          progress: { round: p.round, done_ha: p.done_ha, done_pct: p.done_pct, remaining_ha: p.remaining_ha, started: p.started, day_no: p.day_no }
        };
      });
    const day = events
      .filter(function (e) {
        return e.date === date;
      })
      .sort(function (a, b) {
        return a._row - b._row;
      });
    const out = [];
    day.forEach(function (e) {
      const o = shapeEvent_(e, byEvent[e.event_id] || [], out.length + 1, c, idx, byField, params, today); // numbered among the events this person sees (a phone holds no other)
      if (o) out.push(o);
    });
    return {
      fields: fields,
      events: out,
      methods: allowedMethods_(c),
      today: today,
      editable_from: addDays_(today, -params.sync_grace_days),
      plans: plansAll_().map(shapePlan_),
      fourth_options: FOURTH_COMPOUNDS.slice(),
      tolerance_pct: params.foliar_plan_tolerance_pct,
      due: dueFor_(c, date, today, idx, lines),
      day_max_ha: params.foliar_day_max_ha
    };
  });
}

function api_saveSprayEvent(token, ev) {
  return wrap_(function () {
    const c = authz_(token, ['spray.entry']);
    return withLock_(function () {
      return saveEvent_(c, ev || {});
    });
  });
}

/** The kg of the four compounds of an event: Zinc, Epsom salt, Wuxal and the fourth (Urea or T65). All four must be given (0 when a compound was not used) and the total must be above 0.
 *  A phone from before the compounds existed sends one total kg: that is accepted as it is (legacy: true) and the event then has no compounds. */
function parseMix_(ev) {
  const keys = ['zinc_kg', 'epsom_kg', 'wuxal_kg', 'fourth_kg', 'fourth_name'],
    any = keys.some(function (k) {
      return ev[k] !== undefined && ev[k] !== null && ev[k] !== '';
    });
  if (!any) {
    const kg = cleanNum_(ev.kg, 'Foliar applied (kg)', { required: true, gt: 0, max: 100000 });
    return { legacy: true, total: kg };
  }
  const mix = { legacy: false };
  MIX_FIXED.forEach(function (m) {
    mix[m.key] = cleanNum_(ev[m.key + '_kg'], m.name + ' (kg)', { required: true, min: 0, max: 100000 });
  });
  mix.fourth_name = String(ev.fourth_name || '');
  if (FOURTH_COMPOUNDS.indexOf(mix.fourth_name) < 0) fail_('Choose the fourth compound: ' + FOURTH_COMPOUNDS.join(' or ') + '.', 'VALIDATION');
  mix.fourth = cleanNum_(ev.fourth_kg, mix.fourth_name + ' (kg)', { required: true, min: 0, max: 100000 });
  mix.total = round_(mix.zinc + mix.epsom + mix.wuxal + mix.fourth, 3);
  if (!(mix.total > 0)) fail_('Enter the kg of the compounds that were mixed: the total must be more than 0.', 'VALIDATION');
  if (mix.total > 100000) fail_('The total kg is too high.', 'VALIDATION');
  return mix;
}

const ALL_PCT = 99.95; // a coverage of this % or more means the whole part
const NOTHING_LEFT_HA = 0.0005; // less than this is left of a part: there is nothing to spray
const PER_HR_MAX_PARAM = { Drone: 'drone_ha_per_hr_max', Manual: 'manual_ha_per_hr_max' };

/** The fields of an event: each once, each one this person may record on, with the coverage asked for (the % of the part reached by this event). */
function eventFields_(c, ev, idx) {
  const raw = Array.isArray(ev.lines) ? ev.lines : [];
  if (!raw.length) fail_('Select at least one field.', 'VALIDATION');
  if (raw.length > 100) fail_('Too many fields in one event.', 'VALIDATION');
  const seen = {};
  return raw.map(function (l) {
    l = l || {};
    const f = fieldOrFail_(c, idx, l.field_id);
    if (seen[f.field_id]) fail_('Field ' + fieldLabel_(f) + ' appears twice.', 'VALIDATION');
    seen[f.field_id] = true;
    return { f: f, pct: l.pct };
  });
}

/** How it was sprayed: the phone sends it; an older phone did not, and then only a field with one method tells. */
function eventMethod_(c, ev, fl) {
  let method = String(ev.method || '');
  if (!method) {
    const kinds = {};
    fl.forEach(function (x) {
      kinds[x.f.method] = true;
    });
    method = kinds.Drone && !kinds.Manual && !kinds.Mixed ? 'Drone' : kinds.Manual && !kinds.Drone && !kinds.Mixed ? 'Manual' : '';
    if (kinds.Drone && kinds.Manual) fail_('One spray event can only include fields of one method (all Drone or all Manual).', 'VALIDATION');
  }
  if (METHODS.indexOf(method) < 0) fail_('Choose how the foliar was sprayed: Drone or Manual.', 'VALIDATION');
  if (allowedMethods_(c).indexOf(method) < 0) fail_('Your job role records Drone spray events only.', 'FORBIDDEN');
  return method;
}

/** The crew and the time of an event. A Drone event has one man per drone. */
function eventCrew_(ev, method) {
  const drones = method === 'Drone' ? cleanNum_(ev.drones, 'Number of drones', { required: true, min: 1, max: 50, int: true }) : 0;
  const men = method === 'Drone' ? drones : cleanNum_(ev.men, 'Number of men', { required: true, min: 1, max: 500, int: true });
  if (!isTime_(ev.start_time) || !isTime_(ev.end_time)) fail_('Enter the start and end time.', 'VALIDATION');
  const minutes = minutes_(ev.end_time) - minutes_(ev.start_time);
  if (minutes <= 0) fail_('End time must be after start time.', 'VALIDATION');
  return { drones: drones, men: men, hours: round_(minutes / 60, 3) };
}

/** The event being changed (null for a new one), when this person may change it. */
function existingEvent_(c, ev, id, events) {
  const existing = events.filter(function (e) {
    return e.event_id === id;
  })[0];
  if (!existing) return null;
  if (existing.status !== 'Active') fail_('This spray event was deleted.', 'VALIDATION');
  if (existing.date !== ev.date) fail_('The date of a spray event cannot be changed.', 'VALIDATION');
  if (allowedMethods_(c).indexOf(existing.method) < 0) fail_('Your job role records Drone spray events only.', 'FORBIDDEN');
  return existing;
}

/** The ha one field adds in an event, from the coverage asked for. pp = the progress of that part of the field before the event (partProgress_). */
function lineArea_(label, method, mixed, ext, pp, pctAsked) {
  const part = mixed ? 'the ' + method.toLowerCase() + ' part of this field' : 'this field';
  if (pp.done_today) fail_(label + ': ' + part + ' was already fully sprayed earlier today.', 'VALIDATION');
  if (pp.remaining <= NOTHING_LEFT_HA) fail_(label + ': there is nothing left to spray in the ' + method.toLowerCase() + ' part.', 'VALIDATION');
  const pct = cleanNum_(pctAsked, label + ' coverage', { required: true, gt: 0, max: 100 });
  const target = pct >= ALL_PCT ? ext : (ext * pct) / 100;
  let area = target - pp.done;
  if (area < MIN_LINE_HA)
    fail_(
      label +
        ': set the coverage above ' +
        Math.round(pp.done_pct) +
        '% (that much of the ' +
        method.toLowerCase() +
        ' part was already sprayed earlier in this round).',
      'VALIDATION'
    );
  if (pp.remaining - area < ROUND_TOL) area = pp.remaining; // within a few square metres of the whole part: it is complete
  area = round_(Math.min(area, pp.remaining), 3);
  return { area: area, pct_total: round_(((pp.done + area) / pp.ext) * 100, 1) };
}

/** What is unusual about an event: flags are shown, never a reason to refuse. Returns {flags, plan} (plan = the foliar plan in force that day, none for an older one-total event). */
function eventFlags_(params, method, total, hours, events, id, mix, date) {
  const flags = [],
    hi = params[PER_HR_MAX_PARAM[method]];
  if (hi !== '' && hi !== undefined && total / hours > hi) flags.push('HA/HR HIGH (' + round_(total / hours, 2) + ')');
  if (method === 'Drone') {
    // the drones cannot spray more than a day's capacity
    const dayHa =
      events
        .filter(function (e) {
          return e.status === 'Active' && e.method === 'Drone' && e.date === date && e.event_id !== id;
        })
        .reduce(function (t, e) {
          return t + (Number(e.total_ha) || 0);
        }, 0) + total;
    if (dayHa > params.foliar_day_max_ha + 1e-9) flags.push('DRONE DAY ' + round_(dayHa, 2) + ' HA (MORE THAN ' + params.foliar_day_max_ha + ')');
  }
  const plan = mix.legacy ? null : planOn_(date, plansAll_());
  if (!mix.legacy)
    planFlags_(plan, mix, total, params.foliar_plan_tolerance_pct).forEach(function (fl) {
      flags.push(fl);
    });
  return { flags: flags, plan: plan };
}

/** Share a total (kg, hours) over the lines by their ha: each line is rounded to 3 places and the last one takes what is left, so the lines always add up to the total. */
function shareOut_(total, areas) {
  const sum = areas.reduce(function (t, a) {
    return t + a;
  }, 0);
  let left = round_(total, 3);
  return areas.map(function (a, i) {
    const v = i === areas.length - 1 ? left : round_(total * (sum > 0 ? a / sum : 0), 3);
    left = round_(left - v, 3);
    return v;
  });
}

/** Write an event: its row, its field lines (new, changed, removed), the running % of the fields it touches, and the audit entry. */
function persistEvent_(c, s) {
  const ev = s.ev,
    mix = s.mix,
    by = c.user.username,
    now = nowIso_(),
    existing = s.existing,
    flagText = s.flags.join('; ');
  const none = function (v) {
    return mix.legacy ? '' : v;
  }; // an older one-total event has no compounds
  const evVals = {
    method: s.method,
    kg_applied: round_(mix.total, 3),
    men: s.crew.men,
    drones: s.crew.drones,
    start_time: ev.start_time,
    end_time: ev.end_time,
    hours: s.crew.hours,
    total_ha: s.total,
    flags: flagText,
    zinc_kg: none(round_(mix.zinc, 3)),
    epsom_kg: none(round_(mix.epsom, 3)),
    wuxal_kg: none(round_(mix.wuxal, 3)),
    fourth_name: none(mix.fourth_name),
    fourth_kg: none(round_(mix.fourth, 3)),
    plan_id: s.plan ? s.plan.plan_id : ''
  };
  if (existing) dbUpdate_(SHEETS.SPRAY_EVENTS, existing._row, Object.assign({}, evVals, { edited_by: by, edited_at: now }));
  else dbAppend_(SHEETS.SPRAY_EVENTS, [Object.assign({ event_id: s.id, date: ev.date, status: 'Active', entered_by: by, entered_at: now }, evVals)]);
  const areas = s.lines.map(function (l) {
    return l.area;
  });
  const kg = shareOut_(mix.total, areas),
    hours = shareOut_(s.crew.hours, areas),
    zinc = shareOut_(mix.zinc, areas),
    epsom = shareOut_(mix.epsom, areas),
    wuxal = shareOut_(mix.wuxal, areas),
    fourth = shareOut_(mix.fourth, areas);
  const inserts = [],
    kept = {};
  s.lines.forEach(function (l, i) {
    kept[l.f.field_id] = true;
    const vals = {
      method: s.method,
      start_time: ev.start_time,
      end_time: ev.end_time,
      hours: hours[i],
      units: s.method === 'Drone' ? s.crew.drones : s.crew.men,
      area_ha: l.area,
      kg_applied: kg[i],
      pct_total: l.pct_total,
      flags: flagText,
      zinc_kg: none(zinc[i]),
      epsom_kg: none(epsom[i]),
      wuxal_kg: none(wuxal[i]),
      fourth_name: none(mix.fourth_name),
      fourth_kg: none(fourth[i])
    };
    if (l.old) dbUpdate_(SHEETS.SPRAY, l.old._row, Object.assign(vals, { edited_by: by, edited_at: now }));
    else
      inserts.push(
        Object.assign(
          {
            record_id: newId_('SPR'),
            date: ev.date,
            field_id: l.f.field_id,
            method: s.method,
            prune_date: l.f.last_prune_date,
            status: 'Active',
            entered_by: existing ? existing.entered_by : by,
            entered_at: existing ? existing.entered_at : now,
            event_id: s.id
          },
          vals
        )
      );
  });
  s.oldLines.forEach(function (l) {
    if (!kept[l.field_id])
      dbUpdate_(SHEETS.SPRAY, l._row, { status: 'Voided', void_reason: 'Removed from the event', edited_by: by, edited_at: now });
  });
  dbAppend_(SHEETS.SPRAY, inserts);
  recomputePct_(Object.keys(s.touched), s.idx);
  audit_(by, existing ? 'SPRAY_EVENT_EDIT' : 'SPRAY_EVENT_SAVE', 'spray_event', s.id, {
    date: ev.date,
    method: s.method,
    kg: mix.total,
    mix: mix.legacy ? undefined : [mix.zinc, mix.epsom, mix.wuxal, mix.fourth_name + ' ' + mix.fourth],
    men: s.crew.men,
    drones: s.crew.drones,
    ha: s.total,
    fields: s.lines.length,
    before: existing
      ? { kg: existing.kg_applied, men: existing.men, drones: existing.drones, ha: existing.total_ha, fields: s.oldLines.length }
      : undefined
  });
}

/** Save a spray event (new, or the same event again edited). Everything is checked before anything is written. */
function saveEvent_(c, ev) {
  const params = getParams_(todayStr_());
  entryDateCheck_(ev.date, params);
  const id = String(ev.event_id || '');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) fail_('The spray event has no valid id.', 'VALIDATION');
  const idx = fieldIndex_(),
    fl = eventFields_(c, ev, idx),
    method = eventMethod_(c, ev, fl);
  const mix = parseMix_(ev),
    crew = eventCrew_(ev, method); // the four compounds (or, from an older phone, one total kg), the crew and the time

  const events = dbAll_(SHEETS.SPRAY_EVENTS),
    existing = existingEvent_(c, ev, id, events);
  const all = sprayLinesWithOrder_(events),
    oldLines = all.filter(function (l) {
      return l.event_id === id;
    }),
    byField = groupBy_(all, function (l) {
      return l.field_id;
    });
  if (
    existing &&
    !oldLines.every(function (l) {
      return idx[l.field_id] && inScope_(c, idx[l.field_id]);
    })
  )
    fail_('This event includes fields that are not yours.', 'FORBIDDEN');
  const myRow = existing ? existing._row : sheet_(SHEETS.SPRAY_EVENTS).getLastRow() + 1,
    myKey = ordKey_(ev.date, myRow);

  const lines = fl.map(function (x) {
    const f = x.f,
      label = fieldLabel_(f),
      ext = partHa_(f, method);
    if (ext < MIN_LINE_HA) fail_(label + ': no part of this field is sprayed by ' + method.toLowerCase() + '.', 'VALIDATION');
    const before = (byField[f.field_id] || []).filter(function (l) {
      return l.event_id !== id && l._key < myKey;
    }); // what the field had before this event, on the day and in the order of the sheet
    const pp = partProgress_(f, method, progressAt_(before, f, ev.date, params.spray_round_reset_days));
    const a = lineArea_(label, method, f.method === 'Mixed', ext, pp, x.pct);
    return {
      f: f,
      area: a.area,
      pct_total: a.pct_total,
      old: oldLines.filter(function (l) {
        return l.field_id === f.field_id;
      })[0]
    };
  });
  // replay: the fields of this event (before and after the change) must still make sense with every other event in place
  const touched = {};
  lines.forEach(function (l) {
    touched[l.f.field_id] = true;
  });
  oldLines.forEach(function (l) {
    touched[l.field_id] = true;
  });
  const enteredAt = existing ? existing.entered_at : nowIso_();
  const mine = lines.map(function (l) {
    return { field_id: l.f.field_id, date: ev.date, method: method, area_ha: l.area, entered_at: enteredAt, _ord: myRow, event_id: id };
  });
  checkReplay_(
    'spray',
    all,
    all
      .filter(function (l) {
        return l.event_id !== id;
      })
      .concat(mine),
    idx,
    Object.keys(touched)
  );

  const total = round_(
    lines.reduce(function (t, l) {
      return t + l.area;
    }, 0),
    3
  );
  const unusual = eventFlags_(params, method, total, crew.hours, events, id, mix, ev.date);
  persistEvent_(c, {
    ev: ev,
    id: id,
    method: method,
    mix: mix,
    crew: crew,
    existing: existing,
    oldLines: oldLines,
    lines: lines,
    total: total,
    flags: unusual.flags,
    plan: unusual.plan,
    idx: idx,
    touched: touched
  });
  return {
    saved: 1,
    replaced: existing ? 1 : 0,
    event_id: id,
    fields: lines.length,
    total_ha: total,
    warnings: unusual.flags.length ? [unusual.flags.join('; ')] : []
  };
}

function api_deleteSprayEvent(token, eventId) {
  return wrap_(function () {
    const c = authz_(token, ['spray.entry']);
    return withLock_(function () {
      return deleteEvent_(c, String(eventId || ''), 'Deleted by ' + c.user.username, false);
    });
  });
}

/** Void an event and all its field lines. adminMode (from the Records screen) skips the "today only" rule but not the lock or the "latest event" rule. */
function deleteEvent_(c, id, reason, adminMode) {
  const params = getParams_(todayStr_()),
    today = todayStr_();
  const events = dbAll_(SHEETS.SPRAY_EVENTS),
    e = events.filter(function (x) {
      return x.event_id === id;
    })[0];
  if (!e || e.status !== 'Active') fail_('Spray event not found.', 'VALIDATION');
  lockCheck_(e.date, params);
  const idx = fieldIndex_(),
    all = sprayLinesWithOrder_(events),
    mine = all.filter(function (l) {
      return l.event_id === id;
    });
  if (!adminMode && e.date < addDays_(today, -params.sync_grace_days))
    fail_('Past days are view only. Data can only be changed for today.', 'VALIDATION');
  if (
    !mine.every(function (l) {
      return idx[l.field_id] && inScope_(c, idx[l.field_id]);
    })
  )
    fail_('This event includes fields that are not yours.', 'FORBIDDEN');
  if (!adminMode && allowedMethods_(c).indexOf(e.method) < 0) fail_('Your job role records Drone spray events only.', 'FORBIDDEN');
  const touched = {};
  mine.forEach(function (l) {
    touched[l.field_id] = true;
  });
  checkReplay_(
    'spray',
    all,
    all.filter(function (l) {
      return l.event_id !== id;
    }),
    idx,
    Object.keys(touched)
  );
  const now = nowIso_();
  dbUpdate_(SHEETS.SPRAY_EVENTS, e._row, { status: 'Voided', void_reason: reason, edited_by: c.user.username, edited_at: now });
  mine.forEach(function (l) {
    dbUpdate_(SHEETS.SPRAY, l._row, { status: 'Voided', void_reason: reason, edited_by: c.user.username, edited_at: now });
  });
  recomputePct_(Object.keys(touched), idx);
  audit_(c.user.username, 'SPRAY_EVENT_DELETE', 'spray_event', id, {
    date: e.date,
    kg: e.kg_applied,
    ha: e.total_ha,
    fields: mine.length,
    reason: reason
  });
  return true;
}

// ---- Sync.gs
/**
 * The change feed and the doors a phone uses to keep its own copy of the data (see OFFLINE_REDESIGN_PLAN.md).
 *
 * Every phone holds a copy of the tables its account may see and runs the app's own code on that copy. This file is the server's half:
 *   api_syncPlan      what to copy (the tables, the layout, the epoch, the scope of this account)
 *   api_syncSnapshot  a whole table, a page at a time (the first copy, or after something that makes every copy untrustworthy)
 *   api_syncPull      only the rows written since the phone last asked, from the Changes sheet (written by the central writers in Util.gs)
 *   api_submitOp      the one door for an operation made on the phone (it may arrive more than once: it is done once)
 *   api_deviceKey / api_revokeDevice   a long-lived key per phone, valid for the sync doors only (a session lasts 6 hours at most)
 * What a phone may hold is decided here, per account: only the fields it may see and what belongs to them, no password data, no prices unless the account sees the whole estate.
 */

/** The tables a phone holds a copy of (the Audit, Changes, Devices and OpLog sheets never leave the server). */
const SYNC_TABLES = [
  SHEETS.ESTATE,
  SHEETS.DIVISIONS,
  SHEETS.FIELDS,
  SHEETS.ROLES,
  SHEETS.USERS,
  SHEETS.PARAMS,
  SHEETS.SPRAY_EVENTS,
  SHEETS.SPRAY,
  SHEETS.LEAF,
  SHEETS.PLANS,
  SHEETS.TEMPLATE,
  SHEETS.STAMPS
];
const SYNC_FORMAT = 1; // changes when the way a copy is made changes: every phone then takes everything again
const SYNC_OVERLAP_MS = 60000; // a pull also reads the last minute again (a slow request may write a little after a faster one): rows are replaced by key, so this is harmless
const SYNC_PAGE_ROWS = 1500; // rows in one snapshot answer
const SYNC_MAX_CHANGED = 3000; // more changed rows than this in one pull: the phone takes a snapshot instead
const SYNC_KEEP_CHANGES = 20000; // the Changes sheet keeps about this many rows
const DEVICE_DAYS = 60; // a phone's key lasts this long after its last use
const DEVICE_TOUCH_MS = 12 * 3600000; // the last-use time is written at most this often
const OPS_QUEUEABLE = ['api_saveSprayEvent', 'api_deleteSprayEvent', 'api_savePluck', 'api_voidRecord', 'api_editRecord', 'api_moveFoliarStamp'];
const OPS_FINAL_CODES = ['VALIDATION', 'FORBIDDEN', 'ERROR']; // an operation the server REFUSED for good is remembered as done; a busy or unexpected answer is not (the phone tries again)
let DEVICE_AUTH_ = false; // true only while a sync door runs: then a phone's key is accepted in place of a session token

// ---------- epoch, floor
function epoch_() {
  return PropertiesService.getScriptProperties().getProperty('SYNC_EPOCH') || '0';
}
/** Something happened that the change feed cannot describe (setup, an import, a reset, a failed feed write): every phone takes everything again. */
function bumpEpoch_() {
  PropertiesService.getScriptProperties().setProperty('SYNC_EPOCH', String(Date.now()));
}
function feedFloor_() {
  return Number(PropertiesService.getScriptProperties().getProperty('SYNC_FLOOR') || 0);
}
/** Keep the Changes sheet short: the oldest rows go; a phone that has not asked since before the oldest row left takes a snapshot. */
function pruneChanges_(sh) {
  const n = sh.getLastRow();
  if (n < SYNC_KEEP_CHANGES * 1.5) return;
  const drop = n - 1 - SYNC_KEEP_CHANGES;
  const nextAt = sh.getRange(2 + drop, 1, 1, 1).getValues()[0][0];
  PropertiesService.getScriptProperties().setProperty('SYNC_FLOOR', String(nextAt));
  sh.deleteRows(2, drop);
}

// ---------- what this account may hold
/** The fields the account may see, and whether it sees everything. sig changes when that set changes: the phone then takes a snapshot, because rows that were never its own before are not in the feed. */
function syncScope_(c) {
  const all = !!(c.isAdmin || c.role.field_scope === 'all');
  const fields = {};
  fieldsAll_().forEach(function (f) {
    if (all || inScope_(c, f)) fields[f.field_id] = true;
  });
  const estate = !!(c.isAdmin || hasTask_(c, 'monitor.experiment')); // prices, wages: only for an account that sees the whole estate
  const ids = Object.keys(fields).sort();
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify([SYNC_FORMAT, all, estate, c.user.user_id, ids]))
  ).slice(0, 16);
  return { all: all, estate: estate, fields: fields, sig: sig };
}
/** The spray events with at least one line on a field the account may see (all lines of such an event come along). */
function syncEventsOf_(scope) {
  const out = {};
  dbAll_(SHEETS.SPRAY).forEach(function (l) {
    if (l.event_id && scope.fields[l.field_id]) out[l.event_id] = true;
  });
  return out;
}
const SYNC_PRIVATE_GROUPS = ['Cost rates', 'Revenue']; // parameter groups that hold money
function syncPrivateKeys_() {
  const out = {};
  PARAMS.forEach(function (d) {
    if (SYNC_PRIVATE_GROUPS.indexOf(d.group) >= 0) out[d.key] = true;
  });
  return out;
}
/** What the phone receives for one row: the row as a list in the sheet's column order, or null when the account may not hold it. */
function syncRow_(c, scope, look, table, o) {
  switch (table) {
    case SHEETS.FIELDS:
      if (!scope.all && !scope.fields[o.field_id]) return null;
      break;
    case SHEETS.LEAF:
      if (!scope.all && !scope.fields[o.field_id]) return null;
      break;
    case SHEETS.SPRAY_EVENTS:
      if (!scope.all && !look.events[o.event_id]) return null;
      break;
    case SHEETS.SPRAY:
      if (!scope.all && !scope.fields[o.field_id] && !(o.event_id && look.events[o.event_id])) return null;
      break;
    case SHEETS.USERS:
      o = Object.assign({}, o, { pw_hash: '', failed_count: '', locked_until: '', last_login: '' }); // the staff list (names) is shown on screens; nothing to log in with
      if (o.user_id !== c.user.user_id) o.must_change = false;
      break;
    case SHEETS.PARAMS:
      if (!scope.estate && look.privateKeys[o.key]) return null;
      break;
  }
  return toRow_(table, o);
}
function syncLook_(scope) {
  return { events: scope.all ? {} : syncEventsOf_(scope), privateKeys: syncPrivateKeys_() };
}
/** A short signature of the layout of the synced tables. Plain JavaScript on purpose: the phone computes the same one from its own copy of the code and refuses to work on a copy
 *  whose layout the server has since changed (it waits for the app to update). */
function syncSchemaSig_() {
  const text = JSON.stringify(
    SYNC_TABLES.map(function (t) {
      return [t, SCHEMA[t]];
    })
  );
  let h = 2166136261;
  for (let i = 0; i < text.length; i++) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 16777619) >>> 0;
  }
  return SYNC_FORMAT + '-' + h.toString(36);
}
/** Run a sync door: a phone's key is accepted for the length of it. */
function inSyncDoor_(fn) {
  DEVICE_AUTH_ = true;
  try {
    return fn();
  } finally {
    DEVICE_AUTH_ = false;
  }
}

// ---------- the doors
function api_syncPlan(auth) {
  return wrap_(function () {
    return inSyncDoor_(function () {
      const c = authz_(auth);
      const scope = syncScope_(c);
      return {
        fmt: SYNC_FORMAT,
        schema: syncSchemaSig_(),
        tz: tz_(),
        epoch: epoch_(),
        floor: feedFloor_(),
        now: Date.now(),
        scope: scope.sig,
        tables: SYNC_TABLES.slice(),
        me: meOf_(c)
      };
    });
  });
}

/** A whole table as this account may hold it, a page at a time. `now` (taken before reading) is what the phone asks from next time. */
function api_syncSnapshot(auth, table, from) {
  return wrap_(function () {
    return inSyncDoor_(function () {
      const c = authz_(auth);
      table = String(table || '');
      if (SYNC_TABLES.indexOf(table) < 0) fail_('Unknown table.', 'VALIDATION');
      const now = Date.now(),
        scope = syncScope_(c),
        look = syncLook_(scope);
      const rows = [];
      dbAll_(table).forEach(function (o) {
        const r = syncRow_(c, scope, look, table, o);
        if (r) rows.push(r);
      });
      const start = Math.max(0, Number(from) || 0),
        page = rows.slice(start, start + SYNC_PAGE_ROWS);
      return {
        table: table,
        rows: page,
        more: start + page.length < rows.length,
        next: start + page.length,
        total: rows.length,
        now: now,
        epoch: epoch_(),
        scope: scope.sig
      };
    });
  });
}

/** The rows written since `since` (a server time from an earlier answer): changed rows as lists, deleted (or no longer visible) keys. snapshot:true = take every table again. */
function api_syncPull(auth, since, epochSeen, scopeSeen) {
  return wrap_(function () {
    return inSyncDoor_(function () {
      const c = authz_(auth);
      const now = Date.now(),
        scope = syncScope_(c),
        epoch = epoch_();
      const out = { now: now, epoch: epoch, scope: scope.sig, snapshot: false, rows: {}, deleted: {} };
      since = Number(since) || 0;
      if (!since || String(epochSeen) !== epoch || String(scopeSeen) !== scope.sig || since - SYNC_OVERLAP_MS < feedFloor_()) {
        out.snapshot = true;
        return out;
      }
      const from = since - SYNC_OVERLAP_MS;
      const sh = sheet_(SHEETS.CHANGES),
        n = sh.getLastRow();
      const keys = {};
      let count = 0;
      if (n >= 2) {
        const all = sh.getRange(2, 1, n - 1, 4).getValues();
        for (let i = all.length - 1; i >= 0; i--) {
          if (Number(all[i][0]) < from) continue; // (rows are written in time order, but a slower request may log a moment late: read them all)
          const t = String(all[i][1]);
          if (SYNC_TABLES.indexOf(t) < 0) continue;
          const set = keys[t] || (keys[t] = {});
          if (!set[all[i][2]]) {
            set[all[i][2]] = true;
            count++;
          }
        }
      }
      if (count > SYNC_MAX_CHANGED) {
        out.snapshot = true;
        return out;
      }
      const look = syncLook_(scope);
      Object.keys(keys).forEach(function (t) {
        const byKey = {};
        dbAll_(t).forEach(function (o) {
          byKey[String(o[SCHEMA[t][0]])] = o;
        });
        Object.keys(keys[t]).forEach(function (k) {
          const o = byKey[k],
            r = o ? syncRow_(c, scope, look, t, o) : null;
          if (r) (out.rows[t] = out.rows[t] || []).push(r);
          else (out.deleted[t] = out.deleted[t] || []).push(k);
        });
      });
      return out;
    });
  });
}

/** The current version of named rows (the phone asks for the rows an operation of its own touched, once the server has answered that operation): keys = { table: [key, ...] }. */
function api_syncRows(auth, keys) {
  return wrap_(function () {
    return inSyncDoor_(function () {
      const c = authz_(auth);
      const scope = syncScope_(c),
        look = syncLook_(scope),
        out = { rows: {}, missing: {} };
      keys = keys && typeof keys === 'object' ? keys : {};
      Object.keys(keys).forEach(function (t) {
        if (SYNC_TABLES.indexOf(t) < 0 || !Array.isArray(keys[t])) return;
        const byKey = {};
        dbAll_(t).forEach(function (o) {
          byKey[String(o[SCHEMA[t][0]])] = o;
        });
        keys[t].slice(0, SYNC_MAX_CHANGED).forEach(function (k) {
          const o = byKey[String(k)],
            r = o ? syncRow_(c, scope, look, t, o) : null;
          if (r) (out.rows[t] = out.rows[t] || []).push(r);
          else (out.missing[t] = out.missing[t] || []).push(String(k));
        });
      });
      return out;
    });
  });
}

// ---------- a phone's key
function isDeviceKey_(t) {
  return typeof t === 'string' && t.length === 64 && t.indexOf('dk_') === 0;
}
function deviceKeyHash_(key) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, key));
}
function isoIn_(ms) {
  return Utilities.formatDate(new Date(Date.now() + ms), tz_(), "yyyy-MM-dd'T'HH:mm:ss");
}
/** The account a phone's key belongs to ('' when the key is unknown, revoked or too old). */
function deviceUserId_(key) {
  const hash = deviceKeyHash_(key),
    cache = CacheService.getScriptCache(),
    ck = 'DV_' + hash;
  const hit = cache.get(ck);
  if (hit) return hit;
  const row = dbAll_(SHEETS.DEVICES).filter(function (r) {
    return r.key_hash === hash && !r.revoked;
  })[0];
  if (!row || row.expires < nowIso_()) return '';
  cache.put(ck, row.user_id, 300);
  if (!row.last_used || row.last_used < isoIn_(-DEVICE_TOUCH_MS))
    dbUpdate_(SHEETS.DEVICES, row._row, { last_used: nowIso_(), expires: isoIn_(DEVICE_DAYS * 86400000) });
  return row.user_id;
}
/** A new key for this phone. Needs a normal login (a session), so it is made right after the person has typed the password. */
function api_deviceKey(token, label) {
  return wrap_(function () {
    const c = authz_(token);
    const key = 'dk_' + (Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '')).slice(0, 61);
    const id = newId_('DEV');
    withLock_(function () {
      dbAppend_(SHEETS.DEVICES, [
        {
          device_id: id,
          key_hash: deviceKeyHash_(key),
          user_id: c.user.user_id,
          label: cleanText_(label, 60, 'Label'),
          created_at: nowIso_(),
          last_used: nowIso_(),
          expires: isoIn_(DEVICE_DAYS * 86400000),
          revoked: ''
        }
      ]);
      audit_(c.user.username, 'DEVICE_ADD', 'device', id, '');
    });
    return { device_id: id, key: key, days: DEVICE_DAYS };
  });
}
/** Sign this phone out of the sync (its key stops working). */
function api_revokeDevice(auth) {
  return wrap_(function () {
    return inSyncDoor_(function () {
      const c = authz_(auth, null, true);
      const hash = isDeviceKey_(auth) ? deviceKeyHash_(auth) : '';
      return withLock_(function () {
        const rows = dbAll_(SHEETS.DEVICES).filter(function (r) {
          return !r.revoked && (hash ? r.key_hash === hash : false);
        });
        rows.forEach(function (r) {
          dbUpdate_(SHEETS.DEVICES, r._row, { revoked: nowIso_() });
          CacheService.getScriptCache().remove('DV_' + r.key_hash);
        });
        audit_(c.user.username, 'DEVICE_REVOKE', 'device', rows.length ? rows[0].device_id : '', '');
        return { revoked: rows.length };
      });
    });
  });
}
/** Every phone of an account stops syncing (the account was disabled, or its password was reset). Call inside the lock. */
function revokeDevices_(userId) {
  dbAll_(SHEETS.DEVICES)
    .filter(function (r) {
      return r.user_id === userId && !r.revoked;
    })
    .forEach(function (r) {
      dbUpdate_(SHEETS.DEVICES, r._row, { revoked: nowIso_() });
      CacheService.getScriptCache().remove('DV_' + r.key_hash);
    });
}

// ---------- an operation made on a phone
/** Do an operation of a phone once, however often it arrives. fn is one of OPS_QUEUEABLE and is run for the account of the key exactly as the screen would. Answers {ok, data | error, code} of that function. */
function api_submitOp(auth, opId, fn, args) {
  return wrap_(function () {
    return inSyncDoor_(function () {
      const c = authz_(auth);
      const id = String(opId || ''),
        name = String(fn || '');
      if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) fail_('Unknown action.', 'VALIDATION');
      if (OPS_QUEUEABLE.indexOf(name) < 0 || typeof globalThis[name] !== 'function') fail_('Unknown action.', 'VALIDATION');
      const done = function () {
        return dbAll_(SHEETS.OPLOG).filter(function (r) {
          return r.op_id === id;
        })[0];
      };
      const before = done();
      if (before) {
        if (before.user_id !== c.user.user_id) fail_('Unknown action.', 'VALIDATION'); // an id that somebody else used is not an answer for this person
        return JSON.parse(before.result);
      }
      const env = globalThis[name].apply(null, [auth].concat(Array.isArray(args) ? args : []));
      if (env && (env.ok || OPS_FINAL_CODES.indexOf(env.code) >= 0))
        withLock_(function () {
          if (done()) return;
          dbAppend_(SHEETS.OPLOG, [{ op_id: id, fn: name, user_id: c.user.user_id, result: JSON.stringify(env), at: nowIso_() }]);
          const sh = sheet_(SHEETS.OPLOG),
            n = sh.getLastRow();
          if (n > 6000) sh.deleteRows(2, n - 4001);
        });
      return env;
    });
  });
}

// ---- Users.gs
/** Users (with AFO field assignment), job roles (tasks + field scope), audit log and the Home checklist. */

function activeAdminCount_(users, roles, excludeUserId) {
  return users.filter(function (u) {
    if (!u.active || u.user_id === excludeUserId) return false;
    const r = roles.filter(function (x) {
      return x.role_id === u.role_id;
    })[0];
    return r && r.active && r.is_admin;
  }).length;
}
/** Non-Admin accounts start with the standard password (DEFAULT_PASSWORD) and must change it at first login. Admin accounts get a random one: an Admin can do everything, so a guessable starting password is not acceptable. */
/** Usernames. A person may change their own; entries keep the username of the day they were made, so every earlier username stays on the account (prev_usernames, comma
 *  separated) and nobody else may take it. default_username is the first username the account ever had (empty = never changed): the Admin can put it back. */
const USERNAME_RE = /^[a-z0-9._-]{3,30}$/;
function userPrevNames_(u) {
  return String(u.prev_usernames || '')
    .split(',')
    .map(function (s) {
      return s.trim();
    })
    .filter(function (s) {
      return s;
    });
}
function userDefaultName_(u) {
  return u.default_username || u.username;
}
/** Is this name (a username an entry was made under) one this person has had? */
function ownsName_(user, name) {
  return name === user.username || userPrevNames_(user).indexOf(name) >= 0;
}
function usernameTaken_(users, name, exceptUserId) {
  return users.some(function (u) {
    return u.user_id !== exceptUserId && (u.username === name || userPrevNames_(u).indexOf(name) >= 0);
  });
}
/** Move an account to another username, remembering the one it leaves. */
function renameUser_(user, to) {
  const prev = userPrevNames_(user).filter(function (n) {
    return n !== to;
  });
  if (prev.indexOf(user.username) < 0) prev.push(user.username);
  dbUpdate_(SHEETS.USERS, user._row, {
    username: to,
    default_username: userDefaultName_(user),
    prev_usernames: prev.join(','),
    updated_at: nowIso_()
  });
}

function tempPassword_(isAdminRole) {
  return isAdminRole ? randomString_(12) : DEFAULT_PASSWORD;
}

// ---------------------------------------------------------------- tasks & roles
function api_listTasks(token) {
  return wrap_(function () {
    authz_(token, ['roles.manage']);
    return TASKS.filter(function (t) {
      return !t.auto;
    }).map(function (t) {
      return {
        id: t.id,
        group: t.group,
        label: t.label,
        desc: t.desc,
        sensitive: !!t.sensitive,
        whole_estate: WHOLE_ESTATE_TASKS.indexOf(t.id) >= 0
      };
    });
  });
}

function api_listRoles(token) {
  return wrap_(function () {
    authz_(token, ['roles.manage', 'users.manage']);
    const users = dbAll_(SHEETS.USERS);
    return dbAll_(SHEETS.ROLES).map(function (r) {
      return {
        role_id: r.role_id,
        role_name: r.role_name,
        description: r.description,
        is_admin: r.is_admin,
        active: r.active,
        field_scope: r.is_admin ? 'all' : r.field_scope || 'assigned',
        tasks: r.is_admin
          ? TASKS.map(function (t) {
              return t.id;
            })
          : expandTasks_(
              String(r.tasks || '')
                .split(',')
                .filter(Boolean)
            ).filter(function (t) {
              return (r.field_scope || 'assigned') === 'all' || WHOLE_ESTATE_TASKS.indexOf(t) < 0;
            }),
        user_count: users.filter(function (u) {
          return u.role_id === r.role_id && u.active;
        }).length
      };
    });
  });
}

function api_saveRole(token, r) {
  return wrap_(function () {
    const c = authz_(token, ['roles.manage']);
    r = r || {};
    const name = cleanText_(r.role_name, 50, 'Role name', true);
    const desc = cleanText_(r.description, 200, 'Description', false);
    const scope = String(r.field_scope || 'assigned');
    if (FIELD_SCOPES.indexOf(scope) < 0) fail_('Choose which fields this role can record on.', 'VALIDATION');
    const valid = {};
    TASKS.forEach(function (t) {
      valid[t.id] = t;
    });
    const tasks = expandTasks_(Array.isArray(r.tasks) ? r.tasks : []); // an old phone app may still send 'dashboard.view': it means both parts of Monitoring
    tasks.forEach(function (t) {
      if (!valid[t]) fail_('Unknown task: ' + String(t).slice(0, 40), 'VALIDATION');
      if (scope !== 'all' && WHOLE_ESTATE_TASKS.indexOf(t) >= 0)
        fail_(
          valid[t].label +
            ' compares the whole estate, so it needs a job role that can see every field. Choose "Every field" for this role or untick it.',
          'VALIDATION'
        );
      if (valid[t].sensitive && !c.isAdmin) fail_('Only an Admin can grant "' + valid[t].label + '".', 'FORBIDDEN');
    });
    return withLock_(function () {
      const roles = dbAll_(SHEETS.ROLES);
      if (
        roles.some(function (x) {
          return x.role_name.toLowerCase() === name.toLowerCase() && x.role_id !== r.role_id;
        })
      )
        fail_('A role with this name already exists.', 'VALIDATION');
      if (r.role_id) {
        const cur = roles.filter(function (x) {
          return x.role_id === r.role_id;
        })[0];
        if (!cur) fail_('Role not found.', 'VALIDATION');
        if (cur.is_admin) fail_('The Admin role cannot be changed.', 'FORBIDDEN');
        if (
          !c.isAdmin &&
          String(cur.tasks || '')
            .split(',')
            .some(function (t) {
              return valid[t] && valid[t].sensitive;
            })
        ) {
          fail_('Only an Admin can change a role that holds sensitive tasks.', 'FORBIDDEN');
        }
        const active = r.active !== false;
        if (
          !active &&
          dbAll_(SHEETS.USERS).some(function (u) {
            return u.role_id === cur.role_id && u.active;
          })
        )
          fail_('This role still has active users. Move or deactivate them first.', 'VALIDATION');
        dbUpdate_(SHEETS.ROLES, cur._row, {
          role_name: name,
          description: desc,
          tasks: tasks.join(','),
          field_scope: scope,
          active: active,
          updated_at: nowIso_()
        });
        audit_(c.user.username, 'ROLE_UPDATE', 'role', cur.role_id, { name: name, tasks: tasks, scope: scope, active: active });
        return cur.role_id;
      }
      const id = newId_('ROLE');
      dbAppend_(SHEETS.ROLES, [
        {
          role_id: id,
          role_name: name,
          description: desc,
          tasks: tasks.join(','),
          field_scope: scope,
          is_admin: false,
          active: true,
          created_at: nowIso_(),
          updated_at: nowIso_()
        }
      ]);
      audit_(c.user.username, 'ROLE_CREATE', 'role', id, { name: name, tasks: tasks, scope: scope });
      return id;
    });
  });
}

// ---------------------------------------------------------------- users
function api_listUsers(token) {
  return wrap_(function () {
    authz_(token, ['users.manage']);
    const roles = {};
    dbAll_(SHEETS.ROLES).forEach(function (r) {
      roles[r.role_id] = r;
    });
    const fields = dbAll_(SHEETS.FIELDS);
    const now = nowIso_();
    return dbAll_(SHEETS.USERS).map(function (u) {
      const r = roles[u.role_id] || {};
      const divs = userDivisions_(u),
        direct = fields.filter(function (f) {
          return f.afo_user_id === u.user_id && f.active;
        });
      const mine = fields.filter(function (f) {
        return f.active && (f.afo_user_id === u.user_id || divs.indexOf(f.division_id) >= 0);
      });
      return {
        user_id: u.user_id,
        username: u.username,
        display_name: u.display_name,
        role_id: u.role_id,
        role_name: r.role_name || '?',
        is_admin: !!r.is_admin,
        field_scope: r.is_admin ? 'all' : r.field_scope || 'assigned',
        active: u.active,
        must_change: u.must_change,
        default_username: userDefaultName_(u),
        renamed: userDefaultName_(u) !== u.username,
        locked: !!(u.locked_until && u.locked_until > now),
        last_login: u.last_login,
        field_count: mine.length,
        field_ha: round_(
          mine.reduce(function (s, f) {
            return s + (f.area_ha || 0);
          }, 0),
          1
        ),
        field_ids: direct.map(function (f) {
          return f.field_id;
        }),
        division_ids: divs
      };
    });
  });
}

/** The division ids of a list, checked (they must exist and be active), as the comma list that is stored. */
function cleanDivisionIds_(list) {
  const divs = {};
  dbAll_(SHEETS.DIVISIONS).forEach(function (d) {
    divs[d.division_id] = d;
  });
  const out = [];
  (Array.isArray(list) ? list : []).forEach(function (id) {
    if (!divs[id] || divs[id].active === false) fail_('A selected division was not found.', 'VALIDATION');
    if (out.indexOf(id) < 0) out.push(id);
  });
  return out.join(',');
}

/** Set exactly which fields belong to this user (used for roles with scope "assigned"). */
function assignFields_(userId, wanted) {
  const fields = dbAll_(SHEETS.FIELDS);
  const want = {};
  (wanted || []).forEach(function (id) {
    want[id] = true;
  });
  const known = {};
  fields.forEach(function (f) {
    known[f.field_id] = f;
  });
  Object.keys(want).forEach(function (id) {
    if (!known[id] || !known[id].active) fail_('A selected field was not found.', 'VALIDATION');
  });
  const moves = [],
    now = nowIso_();
  fields.forEach(function (f) {
    if (want[f.field_id] && f.afo_user_id !== userId) moves.push({ row: f._row, patch: { afo_user_id: userId, updated_at: now } });
    else if (!want[f.field_id] && f.afo_user_id === userId) moves.push({ row: f._row, patch: { afo_user_id: '', updated_at: now } });
  });
  dbUpdateMany_(SHEETS.FIELDS, moves);
  return moves.length;
}

function api_saveUser(token, u) {
  return wrap_(function () {
    const c = authz_(token, ['users.manage']);
    u = u || {};
    const display = cleanText_(u.display_name, 60, 'Name', true);
    return withLock_(function () {
      const users = dbAll_(SHEETS.USERS),
        roles = dbAll_(SHEETS.ROLES);
      const role = roles.filter(function (r) {
        return r.role_id === u.role_id;
      })[0];
      if (!role || !role.active) fail_('Choose an active job role.', 'VALIDATION');
      if (role.is_admin && !c.isAdmin) fail_('Only an Admin can assign the Admin role.', 'FORBIDDEN');
      const assigned = !role.is_admin && (role.field_scope || 'assigned') === 'assigned';
      if (u.user_id) {
        const cur = users.filter(function (x) {
          return x.user_id === u.user_id;
        })[0];
        if (!cur) fail_('User not found.', 'VALIDATION');
        const curRole =
          roles.filter(function (r) {
            return r.role_id === cur.role_id;
          })[0] || {};
        if (curRole.is_admin && !c.isAdmin) fail_('Only an Admin can change an Admin user.', 'FORBIDDEN');
        const active = u.active !== false;
        if (cur.user_id === c.user.user_id && !active) fail_('You cannot deactivate your own account.', 'VALIDATION');
        if (curRole.is_admin && (!role.is_admin || !active) && activeAdminCount_(users, roles, cur.user_id) < 1)
          fail_('There must be at least one active Admin.', 'VALIDATION');
        const divIds = assigned ? cleanDivisionIds_(u.division_ids) : '';
        assignFields_(cur.user_id, assigned ? u.field_ids : []);
        dbUpdate_(SHEETS.USERS, cur._row, {
          display_name: display,
          role_id: role.role_id,
          active: active,
          division_ids: divIds,
          updated_at: nowIso_()
        });
        if (!active) revokeDevices_(cur.user_id);
        audit_(c.user.username, 'USER_UPDATE', 'user', cur.user_id, {
          role: role.role_name,
          active: active,
          fields: assigned ? (u.field_ids || []).length : 0,
          divisions: divIds
        });
        return { user_id: cur.user_id };
      }
      const uname = String(u.username || '')
        .trim()
        .toLowerCase();
      if (!USERNAME_RE.test(uname)) fail_('Username must be 3-30 characters: lowercase letters, digits, dot, dash, underscore.', 'VALIDATION');
      if (usernameTaken_(users, uname, '')) fail_('That username is taken.', 'VALIDATION'); // also names an older account used to have
      const newDivs = assigned ? cleanDivisionIds_(u.division_ids) : ''; // check the divisions first: nothing is changed until everything is valid
      const temp = tempPassword_(role.is_admin);
      const id = newId_('USR');
      if (assigned && (u.field_ids || []).length) assignFields_(id, u.field_ids);
      dbAppend_(SHEETS.USERS, [
        {
          user_id: id,
          division_ids: newDivs,
          username: uname,
          display_name: display,
          role_id: role.role_id,
          pw_hash: makeHash_(temp),
          active: true,
          must_change: true,
          failed_count: 0,
          locked_until: '',
          last_login: '',
          created_at: nowIso_(),
          created_by: c.user.username,
          updated_at: nowIso_(),
          default_username: uname,
          prev_usernames: ''
        }
      ]);
      audit_(c.user.username, 'USER_CREATE', 'user', id, {
        username: uname,
        role: role.role_name,
        fields: assigned ? (u.field_ids || []).length : 0
      });
      return { user_id: id, username: uname, temp_password: temp };
    });
  });
}

function api_resetPassword(token, userId) {
  return wrap_(function () {
    const c = authz_(token, ['users.manage']);
    return withLock_(function () {
      const cur = dbAll_(SHEETS.USERS).filter(function (x) {
        return x.user_id === userId;
      })[0];
      if (!cur) fail_('User not found.', 'VALIDATION');
      const role =
        dbAll_(SHEETS.ROLES).filter(function (r) {
          return r.role_id === cur.role_id;
        })[0] || {};
      if (role.is_admin && !c.isAdmin) fail_('Only an Admin can reset an Admin password.', 'FORBIDDEN');
      const temp = tempPassword_(!!role.is_admin);
      dbUpdate_(SHEETS.USERS, cur._row, { pw_hash: makeHash_(temp), must_change: true, failed_count: 0, locked_until: '', updated_at: nowIso_() });
      revokeDevices_(cur.user_id); // the phones of this account sign in again with the new password
      audit_(c.user.username, 'PASSWORD_RESET', 'user', cur.user_id, { username: cur.username });
      return { username: cur.username, temp_password: temp };
    });
  });
}

/** A person changes their own username (the current password confirms it). The old one stays on the account for the entries made under it. */
function api_changeUsername(token, newName, password) {
  return wrap_(function () {
    const c = authz_(token, null);
    const uname = String(newName || '')
      .trim()
      .toLowerCase();
    if (!USERNAME_RE.test(uname)) fail_('Username must be 3-30 characters: lowercase letters, digits, dot, dash, underscore.', 'VALIDATION');
    return withLock_(function () {
      const users = dbAll_(SHEETS.USERS),
        user = users.filter(function (u) {
          return u.user_id === c.user.user_id;
        })[0];
      if (!verifyHash_(String(password || ''), user.pw_hash)) {
        audit_(user.username, 'USERNAME_CHANGE_FAIL', 'user', user.user_id, 'wrong password');
        fail_('Your password is not correct.', 'VALIDATION');
      }
      if (uname === user.username) fail_('That is already your username.', 'VALIDATION');
      if (usernameTaken_(users, uname, user.user_id)) fail_('That username is taken.', 'VALIDATION');
      renameUser_(user, uname);
      audit_(uname, 'USERNAME_CHANGE', 'user', user.user_id, { from: user.username, to: uname });
      return meOf_(ctx_(token));
    });
  });
}

/** Admin: put an account's username back to the one it was created with (its password is reset separately). */
function api_resetUsername(token, userId) {
  return wrap_(function () {
    const c = authz_(token, ['users.manage']);
    return withLock_(function () {
      const users = dbAll_(SHEETS.USERS),
        cur = users.filter(function (x) {
          return x.user_id === userId;
        })[0];
      if (!cur) fail_('User not found.', 'VALIDATION');
      const role =
        dbAll_(SHEETS.ROLES).filter(function (r) {
          return r.role_id === cur.role_id;
        })[0] || {};
      if (role.is_admin && !c.isAdmin) fail_('Only an Admin can reset an Admin username.', 'FORBIDDEN');
      const def = userDefaultName_(cur);
      if (cur.username === def) fail_('This account still has its original username.', 'VALIDATION');
      if (usernameTaken_(users, def, cur.user_id)) fail_('The original username (' + def + ') is now used by another account.', 'VALIDATION');
      renameUser_(cur, def);
      audit_(c.user.username, 'USERNAME_RESET', 'user', cur.user_id, { from: cur.username, to: def });
      return { username: def, was: cur.username };
    });
  });
}

function api_unlockUser(token, userId) {
  return wrap_(function () {
    const c = authz_(token, ['users.manage']);
    return withLock_(function () {
      const cur = dbAll_(SHEETS.USERS).filter(function (x) {
        return x.user_id === userId;
      })[0];
      if (!cur) fail_('User not found.', 'VALIDATION');
      dbUpdate_(SHEETS.USERS, cur._row, { failed_count: 0, locked_until: '' });
      audit_(c.user.username, 'USER_UNLOCK', 'user', cur.user_id, '');
      return true;
    });
  });
}

// ---------------------------------------------------------------- audit & home checklist
function api_listAudit(token, limit) {
  return wrap_(function () {
    authz_(token, ['audit.view']);
    const n = Math.min(Math.max(parseInt(limit, 10) || 200, 1), 500);
    return dbAll_(SHEETS.AUDIT)
      .slice(-n)
      .reverse()
      .map(function (a) {
        return { ts: a.ts, user: a.user, action: a.action, entity: a.entity, entity_id: a.entity_id, details: a.details };
      });
  });
}

function api_setupStatus(token) {
  return wrap_(function () {
    const c = authz_(token, null);
    if (!c.isAdmin) return { admin: false, steps: [] };
    const estate = dbAll_(SHEETS.ESTATE)[0];
    const fields = fieldsAll_().filter(function (f) {
      return f.active;
    });
    const complete = fields.filter(fieldComplete_).length;
    const divTaken = {};
    dbAll_(SHEETS.USERS).forEach(function (u) {
      if (u.active)
        userDivisions_(u).forEach(function (d) {
          divTaken[d] = true;
        });
    });
    const unassigned = fields.filter(function (f) {
      return !f.afo_user_id && !divTaken[f.division_id];
    }).length;
    const pk = {};
    dbAll_(SHEETS.PARAMS).forEach(function (p) {
      pk[p.key] = true;
    });
    const ratesSet = ['worker_rate_per_hour', 'drone_rate_per_hour', 'drone_rate_per_ha', 'gl_price_per_kg'].filter(function (k) {
      return pk[k];
    }).length;
    const roleById = {};
    dbAll_(SHEETS.ROLES).forEach(function (r) {
      roleById[r.role_id] = r;
    });
    const staff = dbAll_(SHEETS.USERS).filter(function (u) {
      return u.active && roleById[u.role_id] && !roleById[u.role_id].is_admin;
    });
    const afos = staff.filter(function (u) {
      return (
        (roleById[u.role_id].field_scope || 'assigned') === 'assigned' &&
        fields.some(function (f) {
          return f.afo_user_id === u.user_id;
        })
      );
    }).length;
    const today = todayStr_();
    const drone = fields.filter(function (f) {
        return partHa_(f, 'Drone') > 0;
      }),
      manual = fields.filter(function (f) {
        return partHa_(f, 'Manual') > 0;
      });
    const haOf = function (l, m) {
      return round_(
        l.reduce(function (s, f) {
          return s + partHa_(f, m);
        }, 0),
        1
      );
    };
    const plToday = dbAll_(SHEETS.LEAF).filter(function (r) {
      return r.status === 'Active' && r.date === today;
    });
    const stats = {
      fields: fields.length,
      complete: complete,
      drone: drone.length,
      drone_ha: haOf(drone, 'Drone'),
      manual: manual.length,
      manual_ha: haOf(manual, 'Manual'),
      mixed: fields.filter(function (f) {
        return f.method === 'Mixed';
      }).length,
      ha: round_(
        fields.reduce(function (t, f) {
          return t + (f.area_ha || 0);
        }, 0),
        1
      ),
      users: staff.length,
      spray_today: dbAll_(SHEETS.SPRAY).filter(function (r) {
        return r.status === 'Active' && r.date === today;
      }).length,
      pluck_today: plToday.length,
      kg_today: round_(
        plToday.reduce(function (s, r) {
          return s + r.kg;
        }, 0),
        0
      )
    };
    return {
      admin: true,
      stats: stats,
      steps: [
        { label: 'Estate and divisions created', done: !!estate, detail: estate ? estate.name : '', task: 'estate.manage' },
        {
          label: 'Fields created and filled in (area, last prune date, VP/SD, drone area)',
          done: fields.length > 0 && complete === fields.length,
          detail: complete + ' of ' + fields.length + ' fields complete',
          task: 'estate.manage'
        },
        { label: 'Cost rates and green leaf price entered', done: ratesSet >= 3, detail: ratesSet + ' of 4 key rates set', task: 'params.manage' },
        {
          label: 'Foliar plan entered (kg per ha of Zinc, Epsom salt, Wuxal and Urea or T65)',
          done: plansAll_().length > 0,
          detail: plansAll_().length + ' plans',
          task: 'foliar.plan'
        },
        { label: 'Staff accounts created (AFOs, pilots, viewers)', done: staff.length > 0, detail: staff.length + ' accounts', task: 'users.manage' },
        {
          label: 'Every field assigned to an AFO',
          done: fields.length > 0 && unassigned === 0,
          detail: unassigned + ' fields have no AFO; ' + afos + ' AFOs have fields',
          task: 'users.manage'
        }
      ]
    };
  });
}

// ---- Util.gs
/** Shared helpers: errors, dates, ids, validation, sheet access. */

function AppError(message, code) {
  this.name = 'AppError';
  this.message = message;
  this.code = code || 'ERROR';
}
AppError.prototype = Object.create(Error.prototype);

function fail_(message, code) {
  throw new AppError(message, code);
}

/** Wrap every client-callable function: never leak stack traces, always return {ok,...}. */
function wrap_(fn) {
  try {
    return { ok: true, data: withRequest_(fn) };
  } catch (e) {
    if (e instanceof AppError || (e && e.name === 'AppError')) return { ok: false, error: e.message, code: e.code };
    console.error('Unexpected: ' + (e && e.stack ? e.stack : e));
    return { ok: false, error: 'Unexpected error. Please try again or contact the Admin.', code: 'UNEXPECTED' };
  }
}
/** Run fn as one request: the sheets it reads are read once (see REQ_ below). */
function withRequest_(fn) {
  REQ_.depth++;
  try {
    return fn();
  } finally {
    if (REQ_.depth === 1) flushChanges_(); // what this call wrote reaches the change feed (see below)
    if (--REQ_.depth === 0) reqEnd_();
  }
}

// ---------- what one API call has already read
/** Inside one API call (wrap_) a sheet is read once: dbAll_ keeps what it read and hands out COPIES (callers sort and change what they get). A write to a sheet (dbAppend_, dbUpdate_,
 *  dbUpdateMany_, dbDeleteRow_) drops that sheet's copy, and taking the lock drops all of them (another request may have written while this one waited). Outside wrap_ (the owner's
 *  functions in the editor) nothing is kept. */
const REQ_ = { depth: 0, rows: {}, book: null, sheets: {} };
/** Forget what was read: of one sheet (name), or of all of them. */
function reqClear_(name) {
  if (name) delete REQ_.rows[name];
  else REQ_.rows = {};
}
/** The call is over: forget everything, also the open spreadsheet. */
function reqEnd_() {
  REQ_.rows = {};
  REQ_.book = null;
  REQ_.sheets = {};
}

// ---------- dates & ids ----------
function tz_() {
  return Session.getScriptTimeZone() || 'Asia/Colombo';
}
function nowIso_() {
  return Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd'T'HH:mm:ss");
}
function todayStr_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}

function isDate_(s) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const d = new Date(s + 'T00:00:00Z');
  return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === s;
}
function addDays_(s, n) {
  const d = new Date(s + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}
function daysBetween_(a, b) {
  return Math.round((new Date(b + 'T00:00:00Z') - new Date(a + 'T00:00:00Z')) / 86400000);
}
function isTime_(s) {
  return typeof s === 'string' && /^([01]\d|2[0-3]):[0-5]\d$/.test(s);
}
function minutes_(s) {
  return parseInt(s.slice(0, 2), 10) * 60 + parseInt(s.slice(3, 5), 10);
}

function newId_(prefix) {
  return prefix + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
}

/** Cryptographically random string from an unambiguous alphabet (256 divisible by 32 = unbiased). */
function randomString_(n) {
  const alphabet = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789!';
  let out = '';
  while (out.length < n) {
    const hex = Utilities.getUuid().replace(/-/g, '');
    for (let i = 0; i + 1 < hex.length && out.length < n; i += 2) {
      out += alphabet[parseInt(hex.substr(i, 2), 16) % 32];
    }
  }
  return out;
}

// ---------- value cleaning ----------
function cleanText_(v, max, label, required) {
  let s = v === null || v === undefined ? '' : String(v);
  s = s
    .replace(/[\u0000-\u001F\u007F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (required && !s) fail_((label || 'Value') + ' is required.', 'VALIDATION');
  if (s.length > (max || 200)) fail_((label || 'Value') + ' is too long (max ' + (max || 200) + ').', 'VALIDATION');
  return s;
}
/** Parse optional/required number; returns '' when blank and not required. */
function cleanNum_(v, label, opts) {
  opts = opts || {};
  if (v === null || v === undefined || String(v).trim() === '') {
    if (opts.required) fail_(label + ' is required.', 'VALIDATION');
    return '';
  }
  const n = Number(String(v).trim());
  if (!isFinite(n)) fail_(label + ' must be a number.', 'VALIDATION');
  if (opts.int && Math.floor(n) !== n) fail_(label + ' must be a whole number.', 'VALIDATION');
  if (opts.min !== undefined && n < opts.min) fail_(label + ' must be at least ' + opts.min + '.', 'VALIDATION');
  if (opts.gt !== undefined && n <= opts.gt) fail_(label + ' must be greater than ' + opts.gt + '.', 'VALIDATION');
  if (opts.max !== undefined && n > opts.max) fail_(label + ' must be at most ' + opts.max + '.', 'VALIDATION');
  return n;
}
function round_(n, d) {
  const f = Math.pow(10, d);
  return Math.round(n * f) / f;
}
function asBool_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

// ---------- sheet access ----------
function ssId_() {
  const id = PropertiesService.getScriptProperties().getProperty('SSID');
  if (!id) fail_('App is not set up. Run setupApp() once from the script editor.', 'NOT_SETUP');
  return id;
}
function sheet_(name) {
  if (REQ_.depth && REQ_.sheets[name]) return REQ_.sheets[name];
  const ss = REQ_.depth ? REQ_.book || (REQ_.book = SpreadsheetApp.openById(ssId_())) : SpreadsheetApp.openById(ssId_());
  let sh = ss.getSheetByName(name);
  if (!sh && LAZY_SHEETS.indexOf(name) >= 0) {
    // added after the first release: create it on first use, so nobody has to run setupApp() just for this
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SCHEMA[name].length).setValues([SCHEMA[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    formatTextColumns_(sh, name);
  }
  if (!sh) fail_('Sheet "' + name + '" is missing. Run setupApp().', 'NOT_SETUP');
  if (REQ_.depth) REQ_.sheets[name] = sh;
  return sh;
}

/** Sheets that were added after the first release: made on first use, so nobody has to run setupApp() just for them. */
const LAZY_SHEETS = [SHEETS.PLANS, SHEETS.TEMPLATE, SHEETS.STAMPS, SHEETS.CHANGES, SHEETS.DEVICES, SHEETS.OPLOG];

/** Plain-text format for the text columns of a sheet: stops Sheets turning dates, months and ids into other types or evaluating '=...'. */
function formatTextColumns_(sh, name) {
  const maxRows = Math.max(sh.getMaxRows() - 1, 1);
  SCHEMA[name].forEach(function (c, i) {
    if (NUMERIC_COLS.indexOf(c) < 0 && BOOL_COLS.indexOf(c) < 0) sh.getRange(2, i + 1, maxRows, 1).setNumberFormat('@');
  });
}

function normCell_(col, v) {
  if (v instanceof Date) {
    if (col === 'month') return Utilities.formatDate(v, tz_(), 'yyyy-MM'); // "2026-11" typed into a sheet that is not text-formatted
    if (DATE_COLS.indexOf(col) >= 0) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
    if (TIME_COLS.indexOf(col) >= 0) return Utilities.formatDate(v, tz_(), 'HH:mm');
    return Utilities.formatDate(v, tz_(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  if (BOOL_COLS.indexOf(col) >= 0) return asBool_(v);
  if (NUMERIC_COLS.indexOf(col) >= 0) return v === '' || v === null ? '' : Number(v);
  return v === null || v === undefined ? '' : typeof v === 'string' ? v : String(v);
}

/** Read all rows of a sheet as objects (with _row). Inside an API call a second read of the same sheet costs nothing (see REQ_); the rows are always copies. */
function dbAll_(name) {
  if (REQ_.depth && REQ_.rows[name]) return REQ_.rows[name].map(copyRow_);
  const sh = sheet_(name);
  const cols = SCHEMA[name];
  const n = sh.getLastRow();
  if (n < 2) return [];
  const vals = sh.getRange(2, 1, n - 1, cols.length).getValues();
  const rows = vals.map(function (r, i) {
    const o = { _row: i + 2 };
    for (let j = 0; j < cols.length; j++) o[cols[j]] = normCell_(cols[j], r[j]);
    return o;
  });
  if (REQ_.depth) {
    REQ_.rows[name] = rows;
    return rows.map(copyRow_);
  }
  return rows;
}
function copyRow_(r) {
  return Object.assign({}, r);
}
function toRow_(name, obj) {
  return SCHEMA[name].map(function (c) {
    const v = obj[c];
    return v === undefined || v === null ? '' : v;
  });
}
/** Append many objects in one write. */
function dbAppend_(name, objs) {
  if (!objs.length) return;
  const sh = sheet_(name);
  const rows = objs.map(function (o) {
    return toRow_(name, o);
  });
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, SCHEMA[name].length).setValues(rows);
  reqClear_(name);
  rows.forEach(function (r) {
    logChange_(name, r[0], 'u');
  });
}
/** Update many rows of one sheet with ONE read and ONE write: updates = [{row: sheet row (_row), patch: {column: value}}]. Call it inside the lock. (A loop of dbUpdate_ costs a read and a write per row.) */
function dbUpdateMany_(name, updates) {
  if (!updates.length) return;
  const sh = sheet_(name),
    cols = SCHEMA[name];
  const rows = updates.map(function (u) {
      return u.row;
    }),
    lo = Math.min.apply(null, rows),
    hi = Math.max.apply(null, rows);
  const range = sh.getRange(lo, 1, hi - lo + 1, cols.length),
    vals = range.getValues();
  updates.forEach(function (u) {
    const row = vals[u.row - lo];
    Object.keys(u.patch).forEach(function (k) {
      const j = cols.indexOf(k);
      if (j >= 0) row[j] = u.patch[k] === undefined || u.patch[k] === null ? '' : u.patch[k];
    });
  });
  range.setValues(vals);
  reqClear_(name);
  updates.forEach(function (u) {
    logChange_(name, vals[u.row - lo][0], 'u');
  });
}
/** Update one row (by _row) with a patch. */
function dbUpdate_(name, rowNumber, patch) {
  const sh = sheet_(name);
  const cols = SCHEMA[name];
  const cur = sh.getRange(rowNumber, 1, 1, cols.length).getValues()[0];
  const obj = {};
  for (let j = 0; j < cols.length; j++) obj[cols[j]] = cur[j];
  Object.keys(patch).forEach(function (k) {
    obj[k] = patch[k];
  });
  sh.getRange(rowNumber, 1, 1, cols.length).setValues([toRow_(name, obj)]);
  reqClear_(name);
  logChange_(name, obj[cols[0]], 'u');
}
/** Delete one row (by _row). The rows below it move up: every _row read before is out of date. */
function dbDeleteRow_(name, rowNumber) {
  const key = SYNC_TABLES.indexOf(name) >= 0 ? sheet_(name).getRange(rowNumber, 1, 1, 1).getValues()[0][0] : '';
  sheet_(name).deleteRow(rowNumber);
  reqClear_(name);
  logChange_(name, key, 'd');
}

// ---------- the change feed
/** Rows written in a synced table wait here until the call is over (or the lock is let go), then go to the Changes sheet in ONE write. A phone asks for what changed since it last asked (Sync.gs). */
let CHANGES_ = [];
function logChange_(name, key, op) {
  if (SYNC_TABLES.indexOf(name) < 0 || key === '' || key === undefined || key === null) return;
  CHANGES_.push([Date.now(), name, String(key), op]);
  if (!REQ_.depth && !LOCK_HELD_) flushChanges_(); // the owner's own functions (setup, imports) write outside a call
}
function flushChanges_() {
  if (!CHANGES_.length) return;
  const rows = CHANGES_;
  CHANGES_ = [];
  try {
    const sh = sheet_(SHEETS.CHANGES);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    pruneChanges_(sh);
  } catch (e) {
    console.error('change feed failed: ' + e);
    bumpEpoch_(); // the phones cannot trust what they have: they take everything again
  }
}

/** Run fn while holding the script lock (serialises all writes). */
let LOCK_HELD_ = false;
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    fail_('The system is busy. Please try again in a moment.', 'BUSY');
  }
  reqClear_(); // another request may have written while this one waited: read again
  LOCK_HELD_ = true;
  try {
    return fn();
  } finally {
    LOCK_HELD_ = false;
    flushChanges_(); // inside the lock, so the order of the feed is the order of the writes
    lock.releaseLock();
  }
}

function audit_(user, action, entity, entityId, details) {
  try {
    dbAppend_(SHEETS.AUDIT, [
      {
        ts: nowIso_(),
        user: user || '',
        action: action,
        entity: entity || '',
        entity_id: entityId || '',
        details: typeof details === 'string' ? details.slice(0, 1000) : JSON.stringify(details || {}).slice(0, 1000)
      }
    ]);
  } catch (e) {
    console.error('audit failed: ' + e);
  }
}

// The local database engine of the phone. It is placed after the server code in one script (tools/build-engine.js), so it can use the server's own functions (api_*,
// dbAll_, ctxOfUser_ ...) and variables. It keeps a copy of the tables the account may see in the sheets of shims.js, takes in what the server sends (a snapshot, then
// the changes since), and runs the server's own api_* functions on that copy: the answer is the one the server would give, with no network.
//
// The engine itself does no networking: the page (or the service worker) asks the server and hands the answers in. The engine has
//   init(plan)                    open the stored copy (or start empty)       plan = { uid, tz, schema, fmt }
//   beginSnapshot / snapshotTable / finishSnapshot    take every table again (first copy, new epoch, new scope)
//   applyPull(res)                take in the changes of api_syncPull
//   run(fn, args, opId)           run an api_* function; a write (opId) is remembered until the server has done it
//   pendingKeys(opId) / ack(opId, server)    the server did (or refused) the operation: swap the phone's guess for the server's rows
//   status()                      { ready, since, epoch, scope, pending }
//   IO.store { load() -> {meta, tables}|null, saveTable(name, rows), saveMeta(meta), clear() }
var __ENG_FMT = 1;
var __ENG_READS = [
  'api_me',
  'api_listTasks',
  'api_getEstate',
  'api_listPlans',
  'api_foliarCalendar',
  'api_historyFields',
  'api_fieldHistory',
  'api_homeSummary',
  'api_pluckRounds',
  'api_sprayRounds',
  'api_expCharts',
  'api_listParams',
  'api_paramHistory',
  'api_entryFields',
  'api_dayRecords',
  'api_listRecords',
  'api_sprayDay'
];
var __ENG_WRITES = ['api_saveSprayEvent', 'api_deleteSprayEvent', 'api_savePluck', 'api_voidRecord', 'api_editRecord', 'api_moveFoliarStamp'];
var __ENG = { uid: '', touched: null, adding: false };
// What the owner's tools and the plan file loader need but the phone does not carry (they stay on the server).
var FOLIAR_PLAN_FILE = null;
function requireOwner_() {
  throw new Error('Not available on the phone.');
}

// There are no sessions on the phone: the account is the one the copy was made for. The day and the time come from the phone's clock in the estate's time zone.
ctx_ = function (token) {
  var c = ctxOfUser_(__ENG.uid, token);
  if (!c) fail_('Account is not active.', 'SESSION_EXPIRED');
  return c;
};
// The phone keeps no change feed and no audit trail of its own (the server writes both when it does the operation); it only notes which rows an operation touched.
flushChanges_ = function () {};
audit_ = function () {};
logChange_ = function (name, key, op) {
  if (!__ENG.touched || SYNC_TABLES.indexOf(name) < 0 || key === '' || key === undefined || key === null) return;
  __ENG.touched.push({ t: name, k: String(key), kind: op === 'd' ? 'del' : __ENG.adding ? 'add' : 'upd' });
};
var __dbAppendReal = dbAppend_;
dbAppend_ = function (name, objs) {
  __ENG.adding = true;
  try {
    __dbAppendReal(name, objs);
  } finally {
    __ENG.adding = false;
  }
};

function __makeEngine(IO) {
  var E = { ready: false };
  var meta = null,
    dirty = {},
    saveTimer = null;
  var sheetOf = function (name) {
    return __SS.getSheetByName(name);
  };

  function initSheets() {
    __SS.sheets = [];
    Object.keys(SCHEMA).forEach(function (n) {
      __SS.insertSheet(n).setHeaders(SCHEMA[n]);
    });
  }
  function freshMeta(plan) {
    return { fmt: __ENG_FMT, schema: plan.schema, uid: plan.uid, epoch: '', scope: '', since: 0, loaded: {}, ready: false, pending: {} };
  }
  function markDirty(name) {
    dirty[name] = true;
    if (!saveTimer) saveTimer = setTimeout(saveNow, 300);
  }
  function touchMeta() {
    if (!saveTimer) saveTimer = setTimeout(saveNow, 300);
  }
  function saveNow() {
    saveTimer = null;
    if (!meta) return Promise.resolve();
    var names = Object.keys(dirty);
    dirty = {};
    var work = names.map(function (n) {
      return IO.store.saveTable(n, sheetOf(n).exportRows(SCHEMA[n].length));
    });
    work.push(IO.store.saveMeta(meta));
    return Promise.all(work).catch(function () {
      /* the storage refused (full): the copy still works from memory; the next start takes it again */
    });
  }
  function posMap(sh) {
    var pos = {},
      last = sh.getLastRow();
    for (var r = 2; r <= last; r++) {
      var row = sh.d[r - 1];
      if (row && row[0] !== '' && row[0] !== undefined) pos[String(row[0])] = r;
    }
    return pos;
  }
  function pad(row, n) {
    var o = new Array(n);
    for (var i = 0; i < n; i++) o[i] = row[i] === undefined || row[i] === null ? '' : row[i];
    return o;
  }
  /** Put rows in by key and take keys out. skip = { 'Table|key': true }: keys an operation of the phone is still changing (the server does not know that yet). */
  function applyRows(name, rows, deleted, skip) {
    var sh = sheetOf(name);
    if (!sh) return;
    var n = SCHEMA[name].length,
      pos = posMap(sh),
      changed = false;
    (rows || []).forEach(function (x) {
      var key = String(x[0]);
      if (skip && skip[name + '|' + key]) return;
      var at = pos[key],
        o = pad(x, n);
      if (at) sh.d[at - 1] = o;
      else {
        var r2 = sh.getLastRow() + 1;
        sh.d[r2 - 1] = o;
        sh.lastRow = r2;
        pos[key] = r2;
      }
      changed = true;
    });
    var gone = false;
    (deleted || []).forEach(function (key) {
      key = String(key);
      if (skip && skip[name + '|' + key]) return;
      var at = pos[key];
      if (at) {
        sh.d[at - 1] = null;
        delete pos[key];
        gone = true;
      }
    });
    if (gone) {
      sh.d = sh.d.filter(function (row, i) {
        return i === 0 || row;
      });
      sh.lastRow = sh.d.length;
      sh.stale = true;
      changed = true;
    }
    if (changed) {
      sh.maxRows = Math.max(sh.maxRows, sh.getLastRow() + 500);
      sh.version++;
      markDirty(name);
    }
  }
  function pendingSkip(exceptOp) {
    var skip = {};
    Object.keys(meta.pending).forEach(function (id) {
      if (id === exceptOp) return;
      meta.pending[id].touched.forEach(function (x) {
        skip[x.t + '|' + x.k] = true;
      });
    });
    return skip;
  }
  function applyServer(rows, deleted, skip) {
    Object.keys(rows || {}).forEach(function (t) {
      applyRows(t, rows[t], null, skip);
    });
    Object.keys(deleted || {}).forEach(function (t) {
      applyRows(t, null, deleted[t], skip);
    });
  }

  E.init = function (plan) {
    initSheets();
    // The server's tables have another layout than the code on this phone knows (the server was updated first): do nothing with the stored copy; the app updates itself
    // the next time it is opened with signal, and until then the screens ask the server.
    if (plan.schema !== syncSchemaSig_()) {
      E.outdated = true;
      E.ready = false;
      meta = freshMeta(plan);
      return Promise.resolve(E.status());
    }
    E.outdated = false;
    return IO.store.load().then(function (data) {
      meta = freshMeta(plan);
      if (data && data.meta && data.meta.fmt === __ENG_FMT && data.meta.schema === plan.schema && data.meta.uid === plan.uid) {
        meta = data.meta;
        Object.keys(data.tables || {}).forEach(function (n) {
          if (sheetOf(n) && SCHEMA[n]) sheetOf(n).importRows(SCHEMA[n], data.tables[n]);
        });
      } else if (data) return IO.store.clear();
    }).then(function () {
      __tz = plan.tz || 'Asia/Colombo';
      __ENG.uid = meta.uid;
      E.ready = !!meta.ready;
      return E.status();
    });
  };
  E.reset = function () {
    clearTimeout(saveTimer);
    saveTimer = null;
    dirty = {};
    meta = null;
    E.ready = false;
    __ENG.uid = '';
    initSheets();
    return IO.store.clear();
  };
  E.flush = function () {
    clearTimeout(saveTimer);
    return meta ? saveNow() : Promise.resolve();
  };

  // ---------- taking the server's rows in
  E.beginSnapshot = function () {
    meta.ready = false;
    meta.loaded = {};
    E.ready = false;
    SYNC_TABLES.forEach(function (n) {
      sheetOf(n).setHeaders(SCHEMA[n]);
      markDirty(n);
    });
  };
  E.snapshotTable = function (name, rows, done) {
    applyRows(name, rows, null, null);
    if (done) meta.loaded[name] = true;
  };
  E.finishSnapshot = function (info) {
    meta.since = info.since;
    meta.epoch = info.epoch;
    meta.scope = info.scope;
    meta.ready = SYNC_TABLES.every(function (n) {
      return meta.loaded[n];
    });
    E.ready = meta.ready;
    touchMeta();
    return saveNow().then(E.status);
  };
  E.applyPull = function (res) {
    applyServer(res.rows, res.deleted, pendingSkip(''));
    meta.since = res.now;
    meta.epoch = res.epoch;
    meta.scope = res.scope;
    touchMeta();
    return E.status();
  };

  // ---------- running the server's functions on the copy
  E.allowed = function (fn) {
    return E.ready && (__ENG_READS.indexOf(fn) >= 0 || __ENG_WRITES.indexOf(fn) >= 0);
  };
  E.run = function (fn, args, opId) {
    if (!E.ready) throw new Error('LOCAL_NOT_READY');
    if (!E.allowed(fn)) throw new Error('LOCAL_UNSUPPORTED: ' + fn);
    var isWrite = __ENG_WRITES.indexOf(fn) >= 0;
    if (isWrite && !opId) throw new Error('LOCAL_NEEDS_OPERATION_ID');
    var versions = {};
    SYNC_TABLES.forEach(function (n) {
      versions[n] = sheetOf(n).version;
    });
    __ENG.touched = [];
    var touched, out;
    try {
      out = globalThis[fn].apply(null, Array.isArray(args) ? args : []);
    } finally {
      touched = __ENG.touched;
      __ENG.touched = null;
      SYNC_TABLES.forEach(function (n) {
        if (sheetOf(n).version !== versions[n]) markDirty(n);
      });
    }
    out = JSON.parse(JSON.stringify(out === undefined ? null : out));
    if (!isWrite) {
      // a screen that only reads but had to make rows (a month's stamps are made the first time the month is opened, once, on the server): nothing of it stays on the phone
      if (touched.length) {
        removeAdds(touched);
        throw new Error('LOCAL_NEEDS_SERVER: ' + fn);
      }
      return { result: out };
    }
    if (out && out.ok) meta.pending[opId] = { fn: fn, touched: touched };
    else if (touched.length) removeAdds(touched); // the function refused the operation after changing something: nothing of it stays
    touchMeta();
    return { result: out, touched: touched.length };
  };
  function removeAdds(touched) {
    var by = {};
    touched.forEach(function (x) {
      if (x.kind === 'add') (by[x.t] = by[x.t] || []).push(x.k);
    });
    Object.keys(by).forEach(function (t) {
      applyRows(t, null, by[t], null);
    });
  }
  /** The rows an operation of the phone changed or deleted (not the ones it added): the server's version of them is asked for when the operation is answered. */
  E.pendingKeys = function (opId) {
    var out = {};
    ((meta.pending[opId] || { touched: [] }).touched || []).forEach(function (x) {
      if (x.kind === 'add') return;
      var list = out[x.t] || (out[x.t] = []);
      if (list.indexOf(x.k) < 0) list.push(x.k);
    });
    return out;
  };
  /** The server has done the operation (or refused it): the rows the phone added for it go, the server's own rows (the changes since the last pull, and the version of every row the
   *  operation touched) come in, in ONE step, so nothing flickers. server = { pull: api_syncPull answer | null, rows: {table: [rows]}, missing: {table: [keys]} } */
  E.ack = function (opId, server) {
    var op = meta.pending[opId];
    delete meta.pending[opId];
    if (op) removeAdds(op.touched);
    var skip = pendingSkip('');
    server = server || {};
    if (server.pull && !server.pull.snapshot) applyServer(server.pull.rows, server.pull.deleted, skip);
    applyServer(server.rows, server.missing, skip);
    if (server.pull && !server.pull.snapshot) {
      meta.since = server.pull.now;
      meta.epoch = server.pull.epoch;
      meta.scope = server.pull.scope;
    }
    touchMeta();
    return E.status();
  };
  E.status = function () {
    return { ready: E.ready, outdated: !!E.outdated, since: meta ? meta.since : 0, epoch: meta ? meta.epoch : '', scope: meta ? meta.scope : '', pending: meta ? Object.keys(meta.pending).length : 0 };
  };
  return E;
}

// The messages the page sends to the engine (the worker glue uses this inside the worker; the tests use it too).
//   state: { engine }   makeIO() builds the engine's IO when it starts   post(reply) sends an answer back
function __engineMessage(state, d, makeIO, post) {
  var id = d.id;
  function ok(v) {
    post({ id: id, ok: v === undefined ? null : v });
  }
  function err(e) {
    post({ id: id, err: String((e && e.message) || e) });
  }
  try {
    if (d.type === 'init') {
      state.engine = __makeEngine(makeIO());
      state.engine.init(d.plan).then(ok, err);
    } else if (!state.engine) err('LOCAL_NOT_READY');
    else if (d.type === 'beginSnapshot') {
      state.engine.beginSnapshot();
      ok(true);
    } else if (d.type === 'snapshotTable') {
      state.engine.snapshotTable(d.name, d.rows, d.done);
      ok(true);
    } else if (d.type === 'finishSnapshot') state.engine.finishSnapshot(d.info).then(ok, err);
    else if (d.type === 'applyPull') ok(state.engine.applyPull(d.res));
    else if (d.type === 'run') ok(state.engine.run(d.fn, d.args, d.opId));
    else if (d.type === 'pendingKeys') ok(state.engine.pendingKeys(d.opId));
    else if (d.type === 'ack') ok(state.engine.ack(d.opId, d.server));
    else if (d.type === 'flush') state.engine.flush().then(ok, err);
    else if (d.type === 'status') ok(state.engine.status());
    else if (d.type === 'reset')
      state.engine.reset().then(function () {
        state.engine = null;
        ok(true);
      }, err);
    else err('unknown message');
  } catch (e) {
    err(e);
  }
}

// Worker glue: runs the engine inside a Web Worker. The page talks to it with messages (handled by __engineMessage in core.js). The worker keeps the copy of the tables in
// IndexedDB; it does no networking (the page asks the server and hands the answers in).
if (typeof importScripts === 'function' && typeof self !== 'undefined' && typeof postMessage === 'function') {
  (function () {
    var state = { engine: null };
    var store = (function () {
      var dbp = null;
      function db() {
        if (!dbp)
          dbp = new Promise(function (res, rej) {
            var q = indexedDB.open('fsx-engine', 1);
            q.onupgradeneeded = function () {
              q.result.createObjectStore('kv');
            };
            q.onsuccess = function () {
              res(q.result);
            };
            q.onerror = function () {
              rej(q.error);
            };
          });
        return dbp;
      }
      function tx(mode, fn) {
        return db().then(function (d) {
          return new Promise(function (res, rej) {
            var t = d.transaction('kv', mode),
              out = fn(t.objectStore('kv'));
            t.oncomplete = function () {
              res(out);
            };
            t.onerror = t.onabort = function () {
              rej(t.error);
            };
          });
        });
      }
      return {
        load: function () {
          return tx('readonly', function (os) {
            var o = { tables: {}, meta: null, got: false },
              r = os.openCursor();
            r.onsuccess = function () {
              var c = r.result;
              if (!c) return;
              o.got = true;
              if (c.key === 'meta') o.meta = c.value;
              else if (String(c.key).indexOf('T:') === 0) o.tables[String(c.key).slice(2)] = c.value;
              c.continue();
            };
            return o;
          })
            .then(function (o) {
              return o.got ? o : null;
            })
            .catch(function () {
              return null;
            });
        },
        saveTable: function (n, rows) {
          return tx('readwrite', function (os) {
            os.put(rows, 'T:' + n);
          });
        },
        saveMeta: function (m) {
          return tx('readwrite', function (os) {
            os.put(m, 'meta');
          });
        },
        clear: function () {
          return tx('readwrite', function (os) {
            os.clear();
          }).catch(function () {});
        }
      };
    })();
    self.onmessage = function (ev) {
      __engineMessage(
        state,
        ev.data,
        function () {
          return { store: store };
        },
        function (m) {
          postMessage(m);
        }
      );
    };
  })();
}
