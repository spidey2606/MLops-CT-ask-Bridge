// ==UserScript==
// @name         CT Clarifications - capture + bridge
// @namespace    ct-clarifications
// @version      0.4.8
// @homepageURL  https://ct-ask.beta.harmony.a2z.com/
// @updateURL    https://ct-ask.beta.harmony.a2z.com/ct-bridge.user.js
// @downloadURL  https://ct-ask.beta.harmony.a2z.com/ct-bridge.user.js
// @description  On abc-mlops: the floating "Clarify" button captures the conversation and opens the CT Clarifications site. On the site: the SharePoint/Slack bridge (contract 0.2.0).
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
  const CONTRACT_VERSION = '0.2.0';
  const SCRIPT_VERSION = '0.4.8';
  // Where the Clarify button opens the site. Harmony beta is the live site for now; switch to
  // https://ct-ask.harmony.a2z.com/ only if prod is ever deployed.
  const CT_SITE_URL = 'https://ct-ask.beta.harmony.a2z.com/';
  // URL prefixes (origin + path) that get the bridge. Each must also be covered by an @match above.
  const BRIDGE_PAGES = ['https://ct-ask.beta.harmony.a2z.com/', 'https://ct-ask.harmony.a2z.com/'];
  const SP = 'https://amazon.sharepoint.com/sites/CTClarifications';
  const API = SP + '/_api';
  const READ_LISTS = ['CT-Open', 'CT-Questions', 'CT-Roles', 'CT-Settings'];
  const WRITE_LISTS = ['CT-Open', 'CT-Questions'];
  // The only CT-Settings rows the page may read. Everything else (e.g. SlackWebhookUrl) stays inside the script.
  const PUBLIC_SETTINGS = ['SlackChannel', 'AnyAuditorLabel', 'AppVersionMin'];
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

  async function opSpQuery(a) {
    checkList(a.list, READ_LISTS);
    // CT-Settings: ignore the page's select/filter/order; read Title+SettingValue and return only public rows.
    if (a.list === 'CT-Settings') a = { list: a.list, select: ['Title', 'SettingValue'] };
    const q = [];
    if (Array.isArray(a.select) && a.select.length) q.push('$select=' + encodeURIComponent(a.select.join(',')));
    if (a.filter) q.push('$filter=' + encodeURIComponent(a.filter));
    if (a.orderby) q.push('$orderby=' + encodeURIComponent(a.orderby));
    const want = Number.isInteger(a.top) && a.top > 0 ? a.top : 5000;
    q.push('$top=' + Math.min(want, 5000));
    let url = API + L(a.list) + '/items' + (q.length ? '?' + q.join('&') : '');
    let items = [];
    // Follow paging until we have what was asked for.
    for (let page = 0; url && items.length < want && page < 50; page++) {
      const body = parse(await req('GET', url, { headers: NOMETA }));
      items = items.concat(body.value || []);
      url = body['odata.nextLink'] || null;
    }
    items = items.slice(0, want);
    if (a.list === 'CT-Settings') items = items.filter((it) => PUBLIC_SETTINGS.includes(it.Title)).map((it) => ({ Title: it.Title, SettingValue: it.SettingValue }));
    return items;
  }

  async function opSpCreate(a) {
    checkList(a.list, WRITE_LISTS); checkFields(a.fields);
    const body = Object.assign({ __metadata: { type: await entityType(a.list) } }, a.fields);
    const r = await write(API + L(a.list) + '/items',
      { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=verbose' }, JSON.stringify(body));
    return parse(r);
  }

  async function opSpUpdate(a) {
    checkList(a.list, WRITE_LISTS); checkId(a.id); checkFields(a.fields);
    const body = Object.assign({ __metadata: { type: await entityType(a.list) } }, a.fields);
    await write(API + L(a.list) + '/items(' + a.id + ')', {
      Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=verbose',
      'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*',
    }, JSON.stringify(body));
    return {};
  }

  async function opSpDelete(a) {
    checkList(a.list, WRITE_LISTS); checkId(a.id);
    await write(API + L(a.list) + '/items(' + a.id + ')', { Accept: 'application/json;odata=nometadata', 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
    return {};
  }

  let hookCache = null;
  async function webhook() {
    if (hookCache && Date.now() - hookCache.at < 5 * 60 * 1000) return hookCache.url;
    const rows = parse(await req('GET', API + L('CT-Settings') + "/items?$select=Title,SettingValue&$filter=Title eq 'SlackWebhookUrl'", { headers: NOMETA })).value || [];
    const url = rows.length ? String(rows[0].SettingValue || '').trim() : '';
    if (!/^https:\/\/hooks\.slack\.com\//.test(url)) throw new Error('No valid SlackWebhookUrl in CT-Settings');
    hookCache = { url, at: Date.now() };
    return url;
  }
  async function opSlack(a) {
    const payload = {};
    for (const k of ['asker', 'recipient', 'headline', 'details', 'link']) payload[k] = String(a[k] == null ? '' : a[k]);
    await req('POST', await webhook(), { headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload), label: 'Slack webhook' });
    return {};
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

  const OPS = { whoami: opWhoami, getDraft: opGetDraft, spQuery: opSpQuery, spCreate: opSpCreate, spUpdate: opSpUpdate, spDelete: opSpDelete, slack: opSlack };

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
  const currentTurn = () => +((bodyText().match(/Annotating customer query #(\d+)/) || [])[1] || 0);
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
  function fieldValue(label) {
    const els = [...document.querySelectorAll('label,span,div')];
    const l = els.find((e) => e.textContent.replace(/\*/g, '').trim() === label && e.textContent.length < label.length + 6);
    if (!l) return '';
    let c = l;
    for (let i = 0; i < 6 && c; i++) { c = c.parentElement; if (c && c.querySelector('input,textarea,button')) break; }
    if (!c) return '';
    const chk = c.querySelector('input[type=radio]:checked');
    if (chk) {
      const lab = chk.closest('label') || (chk.parentElement && chk.parentElement.parentElement);
      return (lab && lab.textContent.trim()) || chk.value || '';
    }
    const ta = c.querySelector('textarea');
    if (ta) return ta.value.trim();
    const b = c.querySelector('button');
    const v = b ? b.textContent.trim() : '';
    // Cloudscape shows placeholder text in an empty dropdown; that is not an answer.
    return /^(Select|Choose|No value|Pick)\b/i.test(v) ? '' : v;
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
    const draft = {
      sourceTool: 'abc-mlops', conversationId: convId(text), marketplace: marketplace(text), accountFlags: accountFlags(),
      currentTurn: currentTurn() || 1, annotatorAlias: myAlias(),
      fieldValues: Object.fromEntries(FIELD_LABELS.map((f) => [f, fieldValue(f)])), turns,
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
  }

  // ---------- entry ----------
  if (isBridgePage()) { if (safeContext()) startBridge(); else console.warn('CT bridge: refused (framed or same-origin opener)'); }
  else if (location.hostname === 'abc-mlops.harmony.a2z.com') startCapture();
  else console.info('CT bridge: not a bridge page: ' + location.origin + location.pathname);
})();
