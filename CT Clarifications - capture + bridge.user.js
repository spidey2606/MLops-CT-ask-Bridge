// ==UserScript==
// @name         CT Clarifications - capture + bridge
// @namespace    ct-clarifications
// @version      0.5.0
// @homepageURL  https://ct-ask.beta.harmony.a2z.com/
// @updateURL    https://ct-ask.beta.harmony.a2z.com/ct-bridge.user.js
// @downloadURL  https://ct-ask.beta.harmony.a2z.com/ct-bridge.user.js
// @description  On abc-mlops: the floating "Clarify" button captures the conversation and opens the CT Clarifications site. On the site: the SharePoint/Slack bridge (contract 0.3.0). Also sends any-auditor reminders from abc-mlops.
// @match        https://abc-mlops.harmony.a2z.com/*
// @match        https://ct-ask.beta.harmony.a2z.com/*
// @match        https://ct-ask.harmony.a2z.com/*
// @noframes
// @sandbox      JavaScript
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_listValues
// @grant        unsafeWindow
// @connect      amazon.sharepoint.com
// @connect      hooks.slack.com
// ==/UserScript==

(function () {
  'use strict';

  // ---------- config ----------
  const CONTRACT_VERSION = '0.3.0';
  const SCRIPT_VERSION = '0.5.0';
  // Where the Clarify button opens the site. Harmony beta is the live site for now; switch to
  // https://ct-ask.harmony.a2z.com/ only if prod is ever deployed.
  const CT_SITE_URL = 'https://ct-ask.beta.harmony.a2z.com/';
  // URL prefixes (origin + path) that get the bridge. Each must also be covered by an @match above.
  const BRIDGE_PAGES = ['https://ct-ask.beta.harmony.a2z.com/', 'https://ct-ask.harmony.a2z.com/'];
  const SP = 'https://amazon.sharepoint.com/sites/CTClarifications';
  const API = SP + '/_api';
  const READ_LISTS = ['CT-Open', 'CT-Questions', 'CT-Roles', 'CT-Settings'];
  // Doubt lists: create, update, delete, recycle.
  const WRITE_LISTS = ['CT-Open', 'CT-Questions'];
  // Admin tab (0.5.0): create/update on CT-Roles and CT-Settings too. SharePoint itself keeps these two lists
  // read-only for Members, so only site owners (Admins) can actually write them; the bridge also limits the columns.
  const ADMIN_LISTS = ['CT-Roles', 'CT-Settings'];
  const ROLE_COLUMNS = ['Title', 'PersonName', 'Role', 'SlackMemberId', 'Active'];
  // The only CT-Settings rows the page may read. Webhook rows are returned as Id + Title only (so the Admin
  // Settings tab can replace them); their values stay inside the script. Unknown rows are never returned.
  const PUBLIC_SETTINGS = ['SlackChannel', 'AnyAuditorLabel', 'AppVersionMin', 'ReopenWindowHours',
    'ReminderIntervalMin', 'ReminderMaxCount', 'RejectReasons', 'FieldValueRules'];
  // Which webhook each Slack message goes to (0.5.0). 'dm' = the direct-message workflow.
  const HOOK_KEYS = { dm: 'SlackWebhookUrl', any: 'SlackChannelAnyWebhookUrl', accepted: 'SlackChannelAcceptedWebhookUrl' };
  const SECRET_SETTINGS = Object.values(HOOK_KEYS);
  const SLACK_FIELDS = { dm: ['asker', 'recipient', 'headline', 'details', 'link'], any: ['asker', 'message', 'link'], accepted: ['auditor', 'asker', 'link'] };
  const DRAFT_PREFIX = 'ctc_draft:';
  const CT_TAB_NAME = 'ct-clarifications';
  const DRAFT_TTL_MS = 24 * 60 * 60 * 1000;

  // ---------- HTTP ----------
  // Shown when SharePoint answers with its sign-in flow instead of data (expired session).
  const SIGN_IN_MSG = 'Your SharePoint sign-in has expired. Open ' + SP + ' in this browser, let it load, then click Retry.';
  // `label` names the request in error messages instead of the URL (the Slack webhook URL is a secret).
  function req(method, url, { headers = {}, body, label } = {}) {
    const what = label || url.split('?')[0];
    const toSP = url.startsWith(SP + '/');
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method, url, headers, data: body, withCredentials: true, anonymous: false, timeout: 25000,
        onload: (r) => {
          // An expired session redirects to the Microsoft login page (or answers 401) instead of JSON.
          const final = String(r.finalUrl || url);
          if (toSP && (r.status === 401 || !final.startsWith('https://amazon.sharepoint.com/'))) return reject(new Error(SIGN_IN_MSG));
          if (r.status >= 200 && r.status < 300) return resolve(r);
          reject(new Error(method + ' ' + what + ' -> HTTP ' + r.status + ' ' + (label ? '' : (r.responseText || '').slice(0, 300))));
        },
        // A blocked redirect to login.microsoftonline.com (not in @connect) also lands here.
        onerror: () => reject(new Error(toSP ? SIGN_IN_MSG : 'Network error: ' + what)),
        ontimeout: () => reject(new Error('Timeout: ' + what)),
      });
    });
  }
  const NOMETA = { Accept: 'application/json;odata=nometadata' };
  // Reads that need each item's ETag (0.5.0): minimal metadata adds "odata.etag" per item.
  const MINMETA = { Accept: 'application/json;odata=minimalmetadata' };
  const parse = (r) => (r.responseText ? JSON.parse(r.responseText) : {});
  const L = (t) => "/web/lists/getbytitle('" + t + "')";

  let digestCache = null;
  async function digest(force) {
    if (!force && digestCache && Date.now() - digestCache.at < 20 * 60 * 1000) return digestCache.value;
    const r = await req('POST', API + '/contextinfo', { headers: NOMETA });
    digestCache = { value: parse(r).FormDigestValue, at: Date.now() };
    return digestCache.value;
  }
  // POST with a digest; retry once with a fresh digest if SharePoint rejects it.
  async function write(url, headers, body) {
    const send = async (d) => req('POST', url, { headers: Object.assign({}, headers, { 'X-RequestDigest': d }), body });
    try { return await send(await digest(false)); }
    catch (e) {
      if (!/HTTP 403/.test(e.message)) throw e;
      return send(await digest(true));
    }
  }
  const entityTypes = {};
  async function entityType(list) {
    if (!entityTypes[list]) {
      entityTypes[list] = parse(await req('GET', API + L(list) + '?$select=ListItemEntityTypeFullName', { headers: NOMETA }))
        .ListItemEntityTypeFullName;
    }
    return entityTypes[list];
  }

  // ---------- bridge ops ----------
  function checkList(list, allowed) {
    if (!allowed.includes(list)) throw new Error('List not allowed: ' + String(list));
  }
  function checkId(id) {
    if (!Number.isInteger(id) || id <= 0) throw new Error('Invalid item id: ' + String(id));
  }
  function checkFields(fields) {
    if (!fields || typeof fields !== 'object' || Array.isArray(fields)) throw new Error('fields must be an object');
    if ('__metadata' in fields) throw new Error('fields must not contain __metadata');
  }

  async function opWhoami() {
    const u = parse(await req('GET', API + '/web/currentuser?$select=LoginName,Email', { headers: NOMETA }));
    const src = String(u.Email || u.LoginName || '');
    const alias = src.split('|').pop().replace(/@.*$/, '').trim().toLowerCase();
    if (!alias) throw new Error('Could not read your alias from SharePoint');
    return { alias };
  }

  // Item as the page sees it: SharePoint's odata.* keys dropped, the ETag kept as __etag (0.5.0).
  function cleanItem(it) {
    const out = {};
    for (const k of Object.keys(it)) if (!k.startsWith('odata.')) out[k] = it[k];
    if (typeof it['odata.etag'] === 'string' && it['odata.etag']) out.__etag = it['odata.etag'];
    return out;
  }
  // GET every page of a list query (no 5000 cap since 0.5.0; `want` stops early when the page asked for `top`).
  async function getItems(list, qs, want) {
    let url = API + L(list) + '/items' + (qs.length ? '?' + qs.join('&') : '');
    let items = [];
    for (let page = 0; url && items.length < want && page < 500; page++) {
      const body = parse(await req('GET', url, { headers: MINMETA }));
      items = items.concat(body.value || []);
      url = body['odata.nextLink'] || null;
    }
    return items.slice(0, want).map(cleanItem);
  }
  // Titles named by a filter like "Title eq 'A' or Title eq 'B'" (any other filter text is ignored).
  function filterTitles(filter) {
    const out = [];
    const re = /Title eq '((?:[^']|'')*)'/g;
    let m;
    while ((m = re.exec(String(filter)))) out.push(m[1].replace(/''/g, "'"));
    return out;
  }
  // CT-Settings: the page's select/filter/order are never sent to SharePoint (a crafted filter could probe a
  // webhook value). Read everything here, then keep the rows the filter names by Title.
  async function settingsQuery(a) {
    const rows = await getItems('CT-Settings', ['$select=Id,Title,SettingValue', '$top=5000'], Infinity);
    const titles = a.filter ? filterTitles(a.filter) : null;
    const out = [];
    for (const r of rows) {
      const t = String(r.Title || '');
      if (titles && !titles.includes(t)) continue;
      if (PUBLIC_SETTINGS.includes(t)) out.push({ Id: r.Id, Title: t, SettingValue: r.SettingValue, __etag: r.__etag });
      else if (SECRET_SETTINGS.includes(t)) out.push({ Id: r.Id, Title: t, __etag: r.__etag });
    }
    return out;
  }
  async function opSpQuery(a) {
    checkList(a.list, READ_LISTS);
    if (a.list === 'CT-Settings') return settingsQuery(a);
    const q = [];
    if (Array.isArray(a.select) && a.select.length) q.push('$select=' + encodeURIComponent(a.select.join(',')));
    if (a.filter) q.push('$filter=' + encodeURIComponent(a.filter));
    if (a.orderby) q.push('$orderby=' + encodeURIComponent(a.orderby));
    const want = Number.isInteger(a.top) && a.top > 0 ? a.top : Infinity;
    q.push('$top=' + Math.min(want, 5000));
    return getItems(a.list, q, want);
  }

  // Admin lists: only known columns, and only known setting rows.
  function checkAdminWrite(a, creating) {
    const keys = Object.keys(a.fields);
    if (a.list === 'CT-Roles') {
      const bad = keys.find((k) => !ROLE_COLUMNS.includes(k));
      if (bad) throw new Error('Column not allowed on CT-Roles: ' + bad);
    } else if (a.list === 'CT-Settings') {
      const allowed = creating ? ['Title', 'SettingValue'] : ['SettingValue'];
      const bad = keys.find((k) => !allowed.includes(k));
      if (bad) throw new Error('Column not allowed on CT-Settings: ' + bad);
      if (creating && !PUBLIC_SETTINGS.concat(SECRET_SETTINGS).includes(String(a.fields.Title))) {
        throw new Error('Setting not allowed: ' + String(a.fields.Title));
      }
    }
  }
  async function opSpCreate(a) {
    checkList(a.list, WRITE_LISTS.concat(ADMIN_LISTS)); checkFields(a.fields);
    if (ADMIN_LISTS.includes(a.list)) checkAdminWrite(a, true);
    if (a.list === 'CT-Settings') {
      // One row per setting: a second row would never be read (lookups use the first match).
      const t = String(a.fields.Title).replace(/'/g, "''");
      const dup = parse(await req('GET', API + L('CT-Settings') + "/items?$select=Id&$filter=Title eq '" + t + "'", { headers: NOMETA })).value || [];
      if (dup.length) throw new Error('Setting already exists: ' + String(a.fields.Title));
    }
    const body = Object.assign({ __metadata: { type: await entityType(a.list) } }, a.fields);
    const r = await write(API + L(a.list) + '/items',
      { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=verbose' }, JSON.stringify(body));
    const item = parse(r);
    if (a.list === 'CT-Settings') delete item.SettingValue; // never echo a webhook back
    return item;
  }

  // ETags look like "3" or W/"3". Anything else is refused rather than sent as a header.
  const ETAG_RE = /^(W\/)?"[\x21\x23-\x7e]{1,64}"$/;
  async function opSpUpdate(a) {
    checkList(a.list, WRITE_LISTS.concat(ADMIN_LISTS)); checkId(a.id); checkFields(a.fields);
    if (ADMIN_LISTS.includes(a.list)) checkAdminWrite(a, false);
    let ifMatch = '*';
    if (a.etag !== undefined && a.etag !== null && a.etag !== '') {
      if (typeof a.etag !== 'string' || !ETAG_RE.test(a.etag)) throw new Error('Invalid etag');
      ifMatch = a.etag; // conditional write: SharePoint answers HTTP 412 if someone changed the item first
    }
    const body = Object.assign({ __metadata: { type: await entityType(a.list) } }, a.fields);
    await write(API + L(a.list) + '/items(' + a.id + ')', {
      Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=verbose',
      'X-HTTP-Method': 'MERGE', 'IF-MATCH': ifMatch,
    }, JSON.stringify(body));
    return {};
  }

  async function opSpDelete(a) {
    checkList(a.list, WRITE_LISTS); checkId(a.id);
    await write(API + L(a.list) + '/items(' + a.id + ')', { Accept: 'application/json;odata=nometadata', 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
    return {};
  }

  // ---------- recycle bin (0.5.0, Admin delete/restore) ----------
  const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  async function opSpRecycle(a) {
    checkList(a.list, WRITE_LISTS); checkId(a.id);
    const r = await write(API + L(a.list) + '/items(' + a.id + ')/recycle()', { Accept: 'application/json;odata=nometadata' });
    const body = parse(r);
    const id = String(body.value || (body.d && body.d.Recycle) || '');
    return { recycleId: id };
  }
  const listDirs = {};
  async function listDir(list) {
    if (!listDirs[list]) {
      const u = parse(await req('GET', API + L(list) + '/RootFolder?$select=ServerRelativeUrl', { headers: NOMETA })).ServerRelativeUrl;
      listDirs[list] = String(u || '').replace(/^\/+/, '').toLowerCase();
    }
    return listDirs[list];
  }
  async function opSpRecycleBin() {
    const dirs = {};
    for (const l of WRITE_LISTS) dirs[await listDir(l)] = l;
    const sel = '$select=Id,Title,LeafName,DirName,DeletedByEmail,DeletedByName,DeletedDate';
    const rows = parse(await req('GET', API + '/web/RecycleBin?' + sel, { headers: NOMETA })).value || [];
    const out = [];
    for (const r of rows) {
      const list = dirs[String(r.DirName || '').replace(/^\/+/, '').toLowerCase()];
      if (!list) continue; // only CT doubts, nothing else from the site's bin
      out.push({
        recycleId: String(r.Id || ''), title: String(r.Title || r.LeafName || ''), list,
        deletedBy: String(r.DeletedByEmail || r.DeletedByName || '').replace(/@.*$/, '').toLowerCase(),
        deletedAt: String(r.DeletedDate || ''),
      });
    }
    out.sort((x, y) => (x.deletedAt < y.deletedAt ? 1 : x.deletedAt > y.deletedAt ? -1 : 0));
    return out;
  }
  async function opSpRestore(a) {
    const id = String(a.recycleId || '');
    if (!GUID_RE.test(id)) throw new Error('Invalid recycle bin id');
    // Only restore what opSpRecycleBin would list (a CT doubt), never other site content.
    const bin = await opSpRecycleBin();
    if (!bin.some((b) => b.recycleId.toLowerCase() === id.toLowerCase())) throw new Error('Not a CT doubt in the recycle bin');
    await write(API + "/web/RecycleBin('" + id + "')/restore()", { Accept: 'application/json;odata=nometadata' });
    return {};
  }

  // ---------- Slack (0.5.0: three webhooks) ----------
  const hookCache = {};
  async function webhook(kind, optional) {
    const key = HOOK_KEYS[kind];
    const c = hookCache[kind];
    if (c && Date.now() - c.at < 5 * 60 * 1000) return c.url;
    const rows = parse(await req('GET', API + L('CT-Settings') + "/items?$select=Title,SettingValue&$filter=Title eq '" + key + "'", { headers: NOMETA })).value || [];
    const url = rows.length ? String(rows[0].SettingValue || '').trim() : '';
    if (!/^https:\/\/hooks\.slack\.com\/(services|triggers|workflows)\//.test(url)) {
      if (optional) return '';
      throw new Error('No valid ' + key + ' in CT-Settings');
    }
    hookCache[kind] = { url, at: Date.now() };
    return url;
  }
  // Channel posts (any / accepted) go to their own workflow. If that workflow is not set up yet, the message
  // falls back to the direct-message workflow (the page always fills recipient/headline/details for it).
  async function sendSlack(a) {
    let kind = a.channel === 'any' || a.channel === 'accepted' ? a.channel : 'dm';
    let url = kind === 'dm' ? await webhook('dm') : await webhook(kind, true);
    if (!url) { kind = 'dm'; url = await webhook('dm'); }
    const payload = {};
    for (const k of SLACK_FIELDS[kind]) payload[k] = String(a[k] == null ? '' : a[k]);
    try {
      await req('POST', url, { headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload), label: 'Slack webhook' });
    } catch (e) {
      delete hookCache[kind]; // a replaced webhook is picked up on the next send, not 5 minutes later
      throw e;
    }
    return {};
  }
  async function opSlack(a) {
    return sendSlack(a);
  }

  function cleanupDrafts() {
    for (const k of GM_listValues()) {
      if (!k.startsWith(DRAFT_PREFIX)) continue;
      const v = GM_getValue(k, null);
      if (!v || !v.savedAt || Date.now() - v.savedAt > DRAFT_TTL_MS) GM_deleteValue(k);
    }
  }
  async function opGetDraft(a) {
    const key = String(a.key || '');
    if (!/^[A-Za-z0-9_-]{8,64}$/.test(key)) return null;
    // Clarify may reuse an already-open CT tab; the draft saved in the abc-mlops tab can take a
    // moment to reach this tab's storage, so wait briefly for it before giving up.
    let v = GM_getValue(DRAFT_PREFIX + key, null);
    for (let i = 0; !v && i < 10; i++) {
      await new Promise((r) => setTimeout(r, 200));
      v = GM_getValue(DRAFT_PREFIX + key, null);
    }
    GM_deleteValue(DRAFT_PREFIX + key); // one-time read
    if (!v || !v.draft || Date.now() - v.savedAt > DRAFT_TTL_MS) return null;
    return v.draft;
  }

  const OPS = {
    whoami: opWhoami, getDraft: opGetDraft, spQuery: opSpQuery, spCreate: opSpCreate, spUpdate: opSpUpdate, spDelete: opSpDelete,
    slack: opSlack, spRecycle: opSpRecycle, spRecycleBin: opSpRecycleBin, spRestore: opSpRestore,
  };

  function isBridgePage() {
    const here = location.origin + location.pathname;
    return BRIDGE_PAGES.some((p) => here.startsWith(p));
  }
  // On a shared origin, another same-origin page that frames or opens us could script our window
  // and drive the bridge. Refuse in a frame, or when the opener is same-origin (readable).
  function safeContext() {
    // Compare real windows: in the Tampermonkey sandbox `window` is a wrapper, so window.top !== window is always true.
    if (window.self !== window.top) return false;
    if (window.opener) {
      try { void window.opener.location.href; return false; } catch (e) { /* cross-origin opener, e.g. abc-mlops: fine */ }
    }
    return true;
  }

  // Page messages can carry either the real window or the sandbox wrapper as source, depending on browser.
  function isSelf(s) {
    return s != null && (s === window || (typeof unsafeWindow !== 'undefined' && s === unsafeWindow) || s === window.self);
  }
  let ponged = false;
  function sameOriginWindow(s) {
    try { return s.location.href.startsWith(location.origin + location.pathname); } catch (e) { return false; }
  }

  function startBridge() {
    cleanupDrafts();
    window.addEventListener('message', async (ev) => {
      const m0 = ev.data;
      if (!m0 || m0.ns !== 'ctc' || m0.type !== 'req') return;
      // Firefox's sandbox wraps windows differently, so identity checks on ev.source fail there.
      // Accept same-origin messages only; safeContext() already refused frames and same-origin openers,
      // so the only same-origin sender that can reach us is this page.
      if (ev.origin !== location.origin || !(isSelf(ev.source) || ev.source == null || sameOriginWindow(ev.source))) {
        console.warn('CT bridge: ignored request from', ev.origin); return;
      }
      const m = m0;
      if (typeof m.id !== 'string') return;
      ponged = true;
      console.info('CT bridge: request', m.op, m.id);
      const reply = (x) => window.postMessage(Object.assign({ ns: 'ctc', type: 'res', id: m.id }, x), location.origin);
      const fn = Object.prototype.hasOwnProperty.call(OPS, m.op) ? OPS[m.op] : null;
      if (!fn) return reply({ ok: false, error: 'Unknown op: ' + String(m.op) });
      try { reply({ ok: true, data: await fn(m.args || {}) }); console.info('CT bridge: ok', m.op, m.id); }
      catch (e) { console.warn('CT bridge: failed', m.op, (e && e.message) || e); reply({ ok: false, error: (e && e.message) || String(e) }); }
    });
    // Announce repeatedly until the site's first request arrives (or 10 s): the site may start listening after
    // our first post. The site keeps the first valid ready, so repeats are harmless.
    const ready = () => window.postMessage({ ns: 'ctc', type: 'ready', version: CONTRACT_VERSION }, location.origin);
    let tries = 0;
    const tick = () => { if (ponged || tries++ >= 40) return; ready(); setTimeout(tick, 250); };
    tick();
    console.info('CT bridge ' + CONTRACT_VERSION + ': started on ' + location.pathname);
  }

  // ---------- abc-mlops capture (same reading logic as v0.1.0) ----------
  const FIELD_LABELS = ['Tool Invoked Accurate', 'HVA Category', 'Interaction Type', 'Static Response Type',
    'Customer Service Routing', 'Response Content Accurate', 'Negative Customer Feedback Classification',
    'Expected Response', 'Observations'];

  function myAlias() {
    try {
      const c = document.cookie.split(';').map((s) => s.trim()).find((s) => s.startsWith('User='));
      if (!c) return '';
      return (JSON.parse(atob(decodeURIComponent(c.slice(5)))).login || '').toLowerCase();
    } catch (e) { return ''; }
  }
  const bodyText = () => document.body.innerText || '';
  const convId = (t = bodyText()) => (t.match(/ID:\s*([0-9a-f-]{36})/i) || [])[1] || '';
  // MP code from "You are annotating DE", or a "Marketplace: DE" line as a fallback. Also accepts 3-letter codes,
  // combined codes like "MX-ES" / "MX_ES" (the MP is the first part, MX) and locales like "en_GB" (the MP is GB).
  // Codes must be upper case so ordinary words ("for", "the") are never read as an MP.
  function marketplace(t = bodyText()) {
    const code = '([A-Z]{2,3}(?:[-_][A-Z]{2,4})?|[a-z]{2}[-_][A-Z]{2})\\b';
    const m = t.match(new RegExp('You are annotating\\s+' + code)) || t.match(new RegExp('\\bMarketplace:\\s*' + code));
    if (!m) return '';
    const v = m[1];
    return /^[a-z]{2}[-_][A-Z]{2}$/.test(v) ? v.slice(3) : v.split(/[-_]/)[0];
  }
  // Why the page counts as a conversation ('' = it does not). The usual sign is the "ID: <uuid>" line; the
  // query header or Customer messages under "Progress:" also count, so a missing or reformatted ID line
  // does not hide the button.
  function conversationSign(t = bodyText()) {
    if (convId(t)) return 'conversation ID';
    if (/Annotating customer query #\d+/.test(t)) return '"Annotating customer query #"';
    if (t.includes('Progress:') && /(^|\n)\s*Customer\s*\n\s*\d{4}-\d\d-\d\d \d\d:\d\d:\d\d UTC/.test(t)) return 'Customer messages';
    return '';
  }
  const currentTurn = (t = bodyText()) => +((t.match(/Annotating customer query #(\d+)/) || [])[1] || 0);
  function accountFlags() {
    const m = bodyText().match(/ID:\s*[0-9a-f-]{36}\s*\n([\s\S]*?)\n\s*Progress:/i);
    return m ? m[1].split('\n').map((s) => s.trim()).filter(Boolean).join(', ') : '';
  }
  function messagesFromReact() {
    const host = [...document.querySelectorAll('div')].find((d) => {
      const k = Object.keys(d).find((x) => x.startsWith('__reactFiber'));
      if (!k) return false;
      const p = d[k] && d[k].memoizedProps;
      const ch = p && p.children;
      if (!Array.isArray(ch)) return false;
      const f = ch.flat();
      return f.length && f[0] && f[0].props && f[0].props.message && 'messageIndex' in f[0].props.message;
    });
    if (!host) return null;
    const k = Object.keys(host).find((x) => x.startsWith('__reactFiber'));
    return host[k].memoizedProps.children.flat().map((c) => c && c.props && c.props.message).filter(Boolean);
  }
  function messagesFromText() {
    const t = bodyText();
    const start = t.indexOf('Progress:'), end = start < 0 ? -1 : t.indexOf('Submit Annotation', start);
    if (start < 0) return [];
    const lines = t.slice(start, end > 0 ? end : undefined).split('\n').map((s) => s.trim());
    const out = []; let cur = null; let idx = 0;
    const ts = /^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d UTC$/;
    const map = { responseSource: 'responseSource', llmIntent: 'llmIntent', llmWorkflow: 'llmWorkflow',
      primary_offline_sub_intent: 'primaryOfflineSubIntent', 'Tool Invoked': 'toolInvocations', 'Origin Page': 'originPage',
      'Origin Sub-Page': 'originSubPage', 'Is Pill Message': 'isPillMessage', 'Page Source URL': 'pageSourceUrl' };
    for (let i = 0; i < lines.length; i++) {
      const ln = lines[i];
      if ((ln === 'Customer' || ln === 'Bot') && ts.test(lines[i + 1] || '')) {
        if (ln === 'Customer') idx++;
        cur = { speaker: ln === 'Customer' ? 'user' : 'bot', timestamp: lines[i + 1], messageIndex: idx, message: '' };
        out.push(cur); i++; continue;
      }
      if (!cur || !ln) continue;
      if (/^(Not annotated|Annotated|← Annotating|LLM Metadata:|Message Metadata:)/.test(ln)) continue;
      const kv = ln.match(/^(responseSource|llmIntent|llmWorkflow|primary_offline_sub_intent|Tool Invoked|Origin Page|Origin Sub-Page|Is Pill Message|Page Source URL):\s*(.*)$/);
      if (kv && cur.speaker === 'bot') { cur[map[kv[1]]] = kv[2]; continue; }
      if (/^Progress:|^\d+%$/.test(ln)) continue;
      cur.message += (cur.message ? '\n' : '') + ln;
    }
    return out;
  }
  const s = (v) => (v == null ? '' : String(v));
  // Turn = customer question + the bot reply that follows. A greeting before turn 1 is dropped.
  function buildTurns() {
    let msgs = null;
    try { msgs = messagesFromReact(); } catch (e) { msgs = null; }
    if (!msgs || !msgs.length) msgs = messagesFromText();
    const turns = [];
    for (const m of msgs) {
      if (m.speaker === 'user') {
        turns.push({ turn: turns.length + 1, customerAt: s(m.timestamp),
          customer: s(m.message), bot: '', pageSourceUrl: '', toolInvoked: '', subIntent: '',
          meta: { responseSource: '', llmIntent: '', llmWorkflow: '', originPage: '', originSubPage: '', isPillMessage: 'false' } });
      } else if (m.speaker === 'bot' && turns.length) {
        const t = turns[turns.length - 1];
        t.bot = t.bot ? t.bot + '\n\n' + s(m.message) : s(m.message);
        t.pageSourceUrl = s(m.pageSourceUrl) || t.pageSourceUrl;
        t.toolInvoked = s(m.toolInvocations) || t.toolInvoked;
        t.subIntent = s(m.primaryOfflineSubIntent) || t.subIntent;
        t.meta = { responseSource: s(m.responseSource), llmIntent: s(m.llmIntent), llmWorkflow: s(m.llmWorkflow),
          originPage: s(m.originPage), originSubPage: s(m.originSubPage), isPillMessage: String(m.isPillMessage) === 'true' ? 'true' : 'false' };
      }
    }
    return turns;
  }
  // The form block that holds one field's input (radio group, dropdown or text box), or null.
  function fieldBox(label, els) {
    const l = (els || [...document.querySelectorAll('label,span,div')])
      .find((e) => e.textContent.replace(/\*/g, '').trim() === label && e.textContent.length < label.length + 6);
    if (!l) return null;
    let c = l;
    for (let i = 0; i < 6 && c; i++) { c = c.parentElement; if (c && c.querySelector('input,textarea,button')) break; }
    return c || null;
  }
  function radioText(r) {
    const lab = r.closest('label') || (r.parentElement && r.parentElement.parentElement);
    return (lab && lab.textContent.trim()) || r.value || '';
  }
  // Cloudscape shows placeholder text in an empty dropdown; that is not an answer.
  const PLACEHOLDER = /^(Select|Choose|No value|Pick)\b/i;
  function boxValue(c) {
    const chk = c.querySelector('input[type=radio]:checked');
    if (chk) return radioText(chk);
    const ta = c.querySelector('textarea');
    if (ta) return ta.value.trim();
    const b = c.querySelector('button');
    const v = b ? b.textContent.trim() : '';
    return PLACEHOLDER.test(v) ? '' : v;
  }
  function fieldValue(label, els) {
    const c = fieldBox(label, els);
    return c ? boxValue(c) : '';
  }

  // ---------- per-turn values + field options (0.5.0) ----------
  // abc-mlops shows one turn's form at a time. While the annotator works, the bridge remembers what each turn
  // had marked (per conversation, 3 days), so Clarify can send every turn's values, not only the one on screen.
  // It also remembers the choices each field offers (radio labels always; dropdown items when the list is
  // open), so the site can offer the same values when an auditor disagrees.
  const TURNVALS_PREFIX = 'ctc_turnvals:';
  const TURNVALS_TTL_MS = 3 * 24 * 60 * 60 * 1000;
  const OPTS_KEY = 'ctc_fieldopts';
  const MAX_OPTS = 200;
  function optionText(o) {
    const lab = o.querySelector('[class*="awsui_label"]');
    return String(o.getAttribute('aria-label') || (lab || o).textContent || '').trim();
  }
  function boxOptions(c) {
    const radios = [...c.querySelectorAll('input[type=radio]')];
    if (radios.length) return radios.map(radioText);
    const b = c.querySelector('[aria-expanded="true"]');
    if (!b) return [];
    const id = b.getAttribute('aria-controls') || b.getAttribute('aria-owns');
    let host = id ? document.getElementById(id) : null;
    if (!host) host = c.querySelector('[role=listbox]');
    if (!host) {
      const all = document.querySelectorAll('[role=listbox]');
      if (all.length === 1) host = all[0];
    }
    return host ? [...host.querySelectorAll('[role=option]')].map(optionText) : [];
  }
  // label -> {value, options} for every field found on the page.
  function readForm() {
    const els = [...document.querySelectorAll('label,span,div')];
    const out = {};
    for (const f of FIELD_LABELS) {
      const c = fieldBox(f, els);
      if (c && c.querySelector('input,textarea,button')) out[f] = { value: boxValue(c), options: boxOptions(c) };
    }
    return out;
  }
  // Options are kept per marketplace (ctc_fieldopts:<MP>): DE values must not show up on an MX doubt.
  function rememberOptions(form, mp) {
    const key = mp ? OPTS_KEY + ':' + String(mp).toUpperCase() : '';
    if (!key) return {};
    const all = GM_getValue(key, null);
    const opts = all && typeof all === 'object' ? all : {};
    let changed = false;
    for (const f of Object.keys(form)) {
      const have = Array.isArray(opts[f]) ? opts[f] : [];
      for (const v of form[f].options) {
        if (!v || PLACEHOLDER.test(v) || have.includes(v) || have.length >= MAX_OPTS) continue;
        have.push(v); changed = true;
      }
      opts[f] = have;
    }
    if (changed) GM_setValue(key, opts);
    return opts;
  }
  function trackTurn() {
    try {
      const text = bodyText();
      const conv = convId(text), turn = currentTurn(text);
      if (!conv || !turn) return;
      const form = readForm();
      const labels = Object.keys(form);
      if (!labels.length) return;
      rememberOptions(form, marketplace(text));
      const k = TURNVALS_PREFIX + conv;
      const rec = GM_getValue(k, null);
      const turns = rec && rec.turns && typeof rec.turns === 'object' ? rec.turns : {};
      const old = turns[String(turn)] || {};
      // A blank read never replaces a saved value: the page may clear the form for the next turn before the
      // "query #N" header changes. The turn on screen is always read fresh at Clarify, so nothing is lost.
      const row = Object.assign({}, old);
      for (const f of labels) if (form[f].value !== '' || !(f in row)) row[f] = form[f].value;
      if (rec && Date.now() - rec.at < 60 * 60 * 1000 && labels.every((f) => old[f] === row[f])) return;
      turns[String(turn)] = row;
      GM_setValue(k, { at: Date.now(), turns });
    } catch (e) { /* best effort: Clarify still reads the turn on screen */ }
  }
  // Values change only when the annotator clicks, picks or types (moving to another turn is a click too), so
  // read the form shortly after that, and once more a little later in case the next turn rendered slowly.
  function watchForm(pill) {
    let t1 = null, t2 = null;
    const later = (e) => {
      if (e && e.target && pill.contains(e.target)) return;
      clearTimeout(t1); clearTimeout(t2);
      t1 = setTimeout(trackTurn, 800);
      t2 = setTimeout(trackTurn, 2500);
    };
    for (const ev of ['click', 'change', 'input', 'keyup']) document.addEventListener(ev, later, true);
    later();
  }
  function cleanupTurnVals() {
    try {
      for (const k of GM_listValues()) {
        if (!String(k).startsWith(TURNVALS_PREFIX)) continue;
        const v = GM_getValue(k, null);
        if (!v || typeof v.at !== 'number' || Date.now() - v.at > TURNVALS_TTL_MS) GM_deleteValue(k);
      }
    } catch (e) { /* storage not listable */ }
  }
  // Every turn's values for Clarify: remembered turns plus the turn on screen (always read fresh).
  function valuesByTurn(conv, cur, curValues, turnCount) {
    const out = {};
    const rec = conv ? GM_getValue(TURNVALS_PREFIX + conv, null) : null;
    if (rec && rec.turns && typeof rec.turns === 'object') {
      for (const t of Object.keys(rec.turns)) {
        const n = +t, v = rec.turns[t];
        if (Number.isInteger(n) && n >= 1 && n <= turnCount && v && typeof v === 'object') out[String(n)] = Object.assign({}, v);
      }
    }
    out[String(cur)] = Object.assign({}, curValues);
    return out;
  }

  function newKey() {
    const a = new Uint8Array(16);
    crypto.getRandomValues(a);
    return [...a].map((x) => x.toString(16).padStart(2, '0')).join('');
  }
  function clarify() {
    const turns = buildTurns();
    if (!turns.length) { alert('CT Clarifications: could not read the conversation on this page.'); return; }
    const text = bodyText();
    const conv = convId(text);
    const cur = currentTurn(text) || 1;
    const els = [...document.querySelectorAll('label,span,div')];
    const fieldValues = Object.fromEntries(FIELD_LABELS.map((f) => [f, fieldValue(f, els)]));
    let fieldOptions = {};
    try { fieldOptions = rememberOptions(readForm(), marketplace(text)); } catch (e) { /* options are optional */ }
    const draft = {
      sourceTool: 'abc-mlops', conversationId: conv, marketplace: marketplace(text), accountFlags: accountFlags(),
      currentTurn: cur, annotatorAlias: myAlias(), fieldValues, turns,
      fieldValuesByTurn: valuesByTurn(conv, cur, fieldValues, turns.length), fieldOptions, bridgeVersion: SCRIPT_VERSION,
    };
    console.info('CT bridge: Clarify captured ' + turns.length + ' turn(s), MP ' + (draft.marketplace || '(not found)') +
      ', conversation ' + (draft.conversationId || '(no ID found)'));
    const key = newKey();
    GM_setValue(DRAFT_PREFIX + key, { draft, savedAt: Date.now() });
    // Named target: reuse the CT tab an earlier Clarify opened (new tab if there is none).
    const win = window.open(CT_SITE_URL + '#/new?draft=' + key, CT_TAB_NAME);
    if (win) { try { win.focus(); } catch (e) { /* browser decides */ } }
    else alert('CT Clarifications: your browser blocked the new tab. Allow pop-ups for abc-mlops.harmony.a2z.com, then click Clarify again.');
  }
  // ---------- floating "Clarify" button (v0.4.2) ----------
  // Dark pill: the word "Clarify" followed by a small orange speech-bubble icon. No hover motion.
  // Single click = capture + open the site. Drag the pill to move it (position is remembered).
  const POS_KEY = 'ctc_btn_pos';
  // The pill's look lives here (and only here), set straight on the elements with !important. abc-mlops has a
  // Content Security Policy (style-src 'self' + nonce) that can block a <style> sheet from a userscript; styles
  // set from script are not blocked, and !important keeps the page's own CSS from overriding them.
  const PILL_WRAP_STYLE = { position: 'fixed', right: '32px', bottom: '32px', 'z-index': '99999' };
  const PILL_STYLE = {
    'box-sizing': 'border-box', display: 'inline-flex', 'align-items': 'center', gap: '9px', height: '44px',
    width: 'auto', margin: '0', padding: '0 18px 0 22px', border: '0', 'border-radius': '999px',
    background: '#0f1115', color: '#fff', 'font-family': '"Amazon Ember",system-ui,"Segoe UI",sans-serif',
    'font-weight': '700', 'font-size': '16px', 'line-height': '1', 'letter-spacing': '.2px',
    'text-transform': 'none', 'white-space': 'nowrap', cursor: 'pointer', outline: 'none',
    'box-shadow': '0 6px 18px rgba(0,0,0,.25)', 'user-select': 'none', '-webkit-user-select': 'none',
    'touch-action': 'none',
  };
  const ICON_STYLE = { width: '19px', height: '19px', flex: 'none', color: '#ff9900' };
  const setStyle = (el, props) => { for (const k in props) el.style.setProperty(k, props[k], 'important'); };
  // Optional extra: the keyboard focus ring (a style can't target :focus-visible directly). If the sheet is
  // blocked, only the ring is lost.
  const FOCUS_CSS = '#ctc-clarify .ctc-pill:focus-visible{outline:2px solid #ff9900 !important;outline-offset:3px !important}';
  function addSheet(css) {
    try { if (typeof GM_addStyle === 'function') { GM_addStyle(css); return; } } catch (e) { /* fall back */ }
    const st = document.createElement('style');
    st.textContent = css;
    document.head.appendChild(st);
  }

  // ---------- reminders (0.5.0) ----------
  // Same rules as the site (domain/reminders.ts), so doubts still get reminded when no CT tab is open: an
  // any-auditor doubt nobody took gets a channel reminder every ReminderIntervalMin minutes, up to
  // ReminderMaxCount times, then one DM to the first active Admin. Each reminder is claimed first with a
  // conditional write (If-Match), so only one tab, site or abc-mlops, ever sends it.
  const REMIND_EVERY_MS = 5 * 60 * 1000;
  const REMIND_FIRST_MS = 20 * 1000;
  const REMIND_LOCK = 'ctc_remind_last';
  const sv = (v) => (v == null ? '' : String(v));
  function numSetting(v, dflt, min, max) {
    const n = sv(v).trim() === '' ? NaN : Number(v);
    return Number.isFinite(n) && n >= min && n <= max ? n : dflt;
  }
  function dueReminder(d, set, now) {
    const st = sv(d.Status);
    if (st !== 'Open' && st !== 'Reopened') return null;
    if (sv(d.AssignedTo).trim().toLowerCase() !== 'any' || sv(d.TakenBy).trim() !== '') return null;
    let base = NaN;
    for (const k of ['LastReminderAt', 'Created']) {
      base = typeof d[k] === 'string' ? Date.parse(d[k]) : NaN;
      if (Number.isFinite(base)) break;
    }
    if (!Number.isFinite(base) || now - base < set.intervalMin * 60e3) return null;
    const n = Number(d.ReminderCount);
    const sent = Number.isFinite(n) ? n : 0;
    if (sent < set.maxCount) return { kind: 'remind', count: sent + 1, max: set.maxCount };
    if (sent === set.maxCount) return { kind: 'remindAdmin', count: sent };
    return null;
  }
  function threadWith(raw, entry) {
    let arr = [];
    try { const v = JSON.parse(sv(raw)); if (Array.isArray(v)) arr = v; } catch (e) { /* unreadable Thread: start a new one, as the site does */ }
    const out = arr.filter((e) => e && typeof e === 'object' && typeof e.type === 'string')
      .map((e) => ({ type: e.type, by: e.by, at: e.at, text: e.text }));
    out.push(entry);
    return JSON.stringify(out);
  }
  // Slack message for one reminder (same wording as the site's notifier), or null when nobody can be named.
  // Same labels and order as the site's field table (columns.ts FIELD_MAP).
  const TOPIC_LABELS = [['ToolInvoked', 'Tool Invoked Accurate'], ['HVACategory', 'HVA Category'],
    ['InteractionType', 'Interaction Type'], ['StaticResponseType', 'Static Response Type'],
    ['CSRouting', 'Customer Service Routing'], ['ResponseContent', 'Response Content Accurate'],
    ['NegFeedback', 'Negative Customer Feedback Classification'], ['ExpectedResponse', 'Expected Response'],
    ['Observations', 'Observations'], ['General', 'General / other']];
  const splitList = (v) => sv(v).split(';').map((x) => x.trim()).filter((x) => x !== '');
  function doubtLabel(d) {
    const want = new Set(splitList(d.Topics));
    const multi = TOPIC_LABELS.filter(([t]) => want.has(t));
    if (multi.length > 1) return multi.map(([, l]) => l).join(', ');
    const one = multi.length === 1 ? multi[0][0] : sv(d.Topic);
    const hit = TOPIC_LABELS.find(([t]) => t === one);
    return hit ? hit[1] : one;
  }
  function doubtTurnText(d) {
    const turns = [...new Set(splitList(d.Turns).map(Number).filter((n) => Number.isInteger(n)))].sort((a, b) => a - b);
    return turns.length > 1 ? 'Turns ' + turns.join(', ') : 'Turn ' + sv(d.QueryNumber);
  }
  function reminderPayload(d, due, roles) {
    const active = roles.filter((r) => r.Active === true && sv(r.SlackMemberId).trim() !== '');
    const admin = active.find((r) => r.Role === 'Admin');
    if (!admin) return null; // same rule as the site: without an Admin nothing is claimed or sent
    const askerA = sv(d.AskedBy).trim().toLowerCase();
    const m = askerA ? active.find((r) => sv(r.Title).trim().toLowerCase() === askerA) : null;
    const asker = m ? m.SlackMemberId : admin ? admin.SlackMemberId : '';
    if (!asker) return null;
    const suffix = m ? '' : ' (for ' + askerA + ')';
    const link = CT_SITE_URL + '#/doubt/' + sv(d.DoubtId);
    const details = sv(d.ConversationId) + ' · ' + doubtTurnText(d) + ' of ' + sv(d.TurnCount) + ' · ' + sv(d.Marketplace) + ' · ' + doubtLabel(d);
    if (due.kind === 'remind') {
      // Channel post; recipient/headline/details are only used if the channel workflow is not set up (DM fallback).
      return { channel: 'any', asker, link, details,
        message: 'asked a doubt that is still waiting. Any available auditor can take it. (Reminder ' + due.count + ' of ' + due.max + ')',
        recipient: admin.SlackMemberId,
        headline: askerA + "'s doubt is still waiting for an auditor (reminder " + due.count + ' of ' + due.max + ').' + suffix };
    }
    return { asker, recipient: admin.SlackMemberId, link, details,
      headline: 'No auditor took ' + askerA + "'s doubt after " + due.count + ' reminders.' + suffix };
  }
  // One pass. Returns the number of reminders sent. Several abc-mlops tabs share one pass per 5 minutes.
  async function remindPass(now = Date.now()) {
    const last = Number(GM_getValue(REMIND_LOCK, 0)) || 0;
    if (now - last < REMIND_EVERY_MS - 30 * 1000) return 0;
    GM_setValue(REMIND_LOCK, now);
    const set = {};
    for (const r of await settingsQuery({ filter: "Title eq 'ReminderIntervalMin' or Title eq 'ReminderMaxCount'" })) set[r.Title] = r.SettingValue;
    const rules = { intervalMin: numSetting(set.ReminderIntervalMin, 30, 5, 24 * 60), maxCount: Math.floor(numSetting(set.ReminderMaxCount, 3, 0, 20)) };
    if (rules.maxCount <= 0) return 0;
    const items = await getItems('CT-Open', ['$filter=' + encodeURIComponent("AssignedTo eq 'ANY'"), '$top=5000'], Infinity);
    const due = items.map((d) => [d, dueReminder(d, rules, now)]).filter((x) => x[1]);
    if (!due.length) return 0;
    const roles = await getItems('CT-Roles', ['$select=Id,Title,Role,SlackMemberId,Active', '$top=5000'], Infinity);
    let sent = 0;
    for (const [d, r] of due) {
      try {
        const p = reminderPayload(d, r, roles);
        // Without an ETag the claim is not safe (two tabs could both send), so skip.
        if (!p || !d.__etag) continue;
        const at = new Date(now).toISOString();
        const text = r.kind === 'remind' ? 'Reminder ' + r.count + ' of ' + r.max + ' sent to #ct-clarifications' : 'No one took it: Admin pinged';
        try {
          await opSpUpdate({ list: 'CT-Open', id: d.Id, etag: d.__etag, fields: {
            ReminderCount: r.kind === 'remind' ? r.count : r.count + 1, LastReminderAt: at,
            Thread: threadWith(d.Thread, { type: 'remind', by: 'ct-ask', at, text }) } });
        } catch (e) {
          if (/HTTP 412|precondition/i.test(e.message)) continue; // another tab claimed it first
          throw e;
        }
        await sendSlack(p);
        sent += 1;
      } catch (e) { console.warn('CT bridge: reminder for doubt ' + sv(d.DoubtId) + ' failed: ' + (e && e.message)); }
    }
    return sent;
  }
  function startReminders() {
    const run = () => remindPass()
      .then((n) => { if (n) console.info('CT bridge: sent ' + n + ' reminder(s)'); })
      .catch((e) => console.warn('CT bridge: reminder pass skipped: ' + (e && e.message)));
    setTimeout(() => { run(); setInterval(run, REMIND_EVERY_MS); }, REMIND_FIRST_MS);
  }

  function startCapture() {
    addSheet(FOCUS_CSS);

    const w = document.createElement('div');
    w.id = 'ctc-clarify';
    w.style.display = 'none';
    w.innerHTML =
      '<button class="ctc-pill" type="button" title="Ask an auditor about this conversation">Clarify' +
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' +
      '<path d="M21 12a8 8 0 0 1-11.6 7.1L4 20l1-4.6A8 8 0 1 1 21 12z"/></svg>' +
      '</button>';
    setStyle(w, PILL_WRAP_STYLE);
    document.body.appendChild(w);
    const orb = w.querySelector('.ctc-pill');
    setStyle(orb, PILL_STYLE);
    setStyle(orb.querySelector('svg'), ICON_STYLE);

    // Remembered position (right/bottom offsets), clamped to the viewport.
    function place(pos) {
      const max = (v, lim, size) => Math.max(8, Math.min(v, lim - (size || 56) - 8));
      setStyle(w, {
        right: max(pos.right, window.innerWidth, w.offsetWidth) + 'px',
        bottom: max(pos.bottom, window.innerHeight, w.offsetHeight) + 'px',
      });
    }
    const saved = GM_getValue(POS_KEY, null);
    if (saved && typeof saved.right === 'number' && typeof saved.bottom === 'number') place(saved);
    window.addEventListener('resize', () => {
      const p = GM_getValue(POS_KEY, null);
      if (p) place(p);
    });

    // Drag vs click: moving more than 5px turns the press into a drag and swallows the click.
    let drag = null;
    let suppressClick = false;
    orb.addEventListener('pointerdown', (e) => {
      if (e.button !== 0) return;
      suppressClick = false;
      const r = w.getBoundingClientRect();
      drag = { x: e.clientX, y: e.clientY, right: window.innerWidth - r.right, bottom: window.innerHeight - r.bottom, moved: false };
      orb.setPointerCapture(e.pointerId);
    });
    orb.addEventListener('pointermove', (e) => {
      if (!drag) return;
      const dx = e.clientX - drag.x, dy = e.clientY - drag.y;
      if (!drag.moved && Math.hypot(dx, dy) < 5) return;
      drag.moved = true;
      w.classList.add('ctc-dragging');
      setStyle(orb, { cursor: 'grabbing' });
      place({ right: drag.right - dx, bottom: drag.bottom - dy });
    });
    function endDrag() {
      if (!drag) return;
      if (drag.moved) {
        suppressClick = true;
        GM_setValue(POS_KEY, { right: parseFloat(w.style.right), bottom: parseFloat(w.style.bottom) });
      }
      w.classList.remove('ctc-dragging');
      setStyle(orb, { cursor: 'pointer' });
      drag = null;
    }
    orb.addEventListener('pointerup', endDrag);
    orb.addEventListener('pointercancel', endDrag);

    orb.addEventListener('click', (e) => {
      if (suppressClick) { suppressClick = false; return; }
      clarify();
    });

    // Show the pill only on a conversation. Skip work while the tab is hidden, touch the DOM only on change,
    // and re-place it when it appears (its real width is unknown while hidden).
    // Logs each change once, so "no Clarify button" can be diagnosed from the browser console.
    let shown = null;
    const check = () => {
      if (document.hidden) return;
      const sign = conversationSign();
      const want = !!sign;
      if (want === shown) return;
      shown = want;
      w.style.display = want ? '' : 'none';
      if (want) console.info('CT bridge: Clarify shown (found ' + sign + ')');
      else console.info('CT bridge: Clarify hidden: no conversation found on this page ' +
        '(looked for "ID: <conversation id>", "Annotating customer query #", or Customer messages under "Progress:")');
      const p = GM_getValue(POS_KEY, null);
      if (want && p && typeof p.right === 'number' && typeof p.bottom === 'number') place(p);
    };
    console.info('CT bridge ' + SCRIPT_VERSION + ': Clarify capture started on ' + location.pathname);
    check();
    setInterval(check, 1500);
    cleanupTurnVals();
    watchForm(w);
    startReminders();
  }

  // ---------- entry ----------
  if (isBridgePage()) { if (safeContext()) startBridge(); else console.warn('CT bridge: refused (framed or same-origin opener)'); }
  else if (location.hostname === 'abc-mlops.harmony.a2z.com') startCapture();
  else console.info('CT bridge: not a bridge page: ' + location.origin + location.pathname);
})();
