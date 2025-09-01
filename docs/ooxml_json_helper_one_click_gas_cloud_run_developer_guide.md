# OOXML JSON Helper – One‑Click GAS + Cloud Run (Developer Guide)

A single artifact that lets any Google Apps Script (GAS) project deploy a tiny Cloud Run service to **unwrap** OOXML files (PPTX/DOCX/XLSX/THMX) into a JSON manifest, let you edit the manifest, and **rewrap** it back — with an optional **/process** endpoint for server‑side ops. Includes **preflight** (billing/APIs), **budget setup**, **per‑session helper** with signed URLs, and **GAS callers**.

> Tone: pragmatic. Copy → paste → run. No fluff.

---

## TL;DR

- Run `showGcpPreflight()` in GAS → enable billing/APIs, create budget.
- Run `initAndDeploy()` → Cloud Run service gets deployed; `CF_BASE` saved to Script Properties.
- Use `callService_unwrap` / `callService_rewrap` for small files, or **session flow** for 100MB+.
- Optional: `callService_process` to run ops server‑side in one shot.

---

## What You Get

- **Cloud Run service** (Node 18, Express, fflate) with endpoints:
  - `GET /ping`
  - `POST /session/new` → signed URLs & GCS paths for a temporary session
  - `POST /unwrap` → JSON manifest (XML inline, binaries referenced)
  - `POST /rewrap` → rebuild zip (to GCS or base64)
  - `POST /process` → replace/upsert/remove parts (zipB64 or GCS in/out)
- **GAS preflight sidebar**: verify billing, enable required APIs, create/update budget.
- **One‑click deploy** from GAS: uploads service source to GCS, runs Cloud Build → Cloud Run.
- **GAS wrappers** to call the service, including per‑session large‑file flow.

---

## 1) Apps Script – Project Setup

### 1.1 `appsscript.json` scopes

```json
{
  "timeZone": "Europe/Amsterdam",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "oauthScopes": [
    "https://www.googleapis.com/auth/cloud-platform",
    "https://www.googleapis.com/auth/cloud-billing",
    "https://www.googleapis.com/auth/cloud-billing.readonly",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/devstorage.full_control",
    "https://www.googleapis.com/auth/drive"
  ]
}
```

### 1.2 GAS Code (preflight + deploy + callers)

Paste the whole block into a single script file (e.g., `Main.gs`).

```javascript
/**
 * OOXML JSON Helper – Apps Script side
 * - Preflight: billing, APIs, budget
 * - One‑click deploy to Cloud Run (uploads source, Cloud Build → Run)
 * - Script Property CF_BASE set to deployed base URL
 * - Callers: unwrap/rewrap/process + large‑file session helpers
 */

const CONFIG = {
  PROJECT_ID: 'your-gcp-project-id',
  REGION:     'europe-west4',       // or us-central1 if you want Always Free egress
  SERVICE:    'ooxml-json',
  BUCKET:     null,                 // null => auto: "<PROJECT_ID>-ooxml-src"
  PUBLIC:     true,                 // false => protected; GAS fetches ID token
  RUN_SA:     null,                 // optional service account email for Cloud Run

  // Budget defaults (EUR 5, thresholds 10/50/90)
  BUDGET_NAME: 'OOXML Helper Budget',
  BUDGET_CURRENCY: 'EUR',
  BUDGET_AMOUNT_UNITS: '5',
  BUDGET_THRESHOLDS: [0.10, 0.50, 0.90]
};

const REQUIRED_APIS = [
  'run.googleapis.com',
  'cloudbuild.googleapis.com',
  'artifactregistry.googleapis.com',
  'iam.googleapis.com',
  'serviceusage.googleapis.com',
  'cloudbilling.googleapis.com',
  'cloudresourcemanager.googleapis.com',
  'billingbudgets.googleapis.com',
  'storage.googleapis.com'
];

// ===== Entry points =====

function showGcpPreflight() {
  const html = HtmlService.createHtmlOutput(_preflightHtml_(CONFIG.PROJECT_ID))
    .setTitle('GCP Preflight')
    .setWidth(460);
  // Use any editor container you like:
  SpreadsheetApp.getUi().showSidebar(html);
}

function initAndDeploy() {
  const p = CONFIG.PROJECT_ID, r = CONFIG.REGION, s = CONFIG.SERVICE;
  if (!p) throw new Error('Set CONFIG.PROJECT_ID');

  // Billing sanity – hard fail early
  const bill = _pf_checkBilling();
  if (!bill.enabled) throw new Error('Billing not enabled. Use showGcpPreflight() first.');

  // Enable APIs
  _pf_enableRequiredApis();

  // Bucket for source zip
  const bucket = CONFIG.BUCKET || (p + '-ooxml-src');
  ensureBucket_(p, bucket, r);

  // Build source zip for Cloud Run service
  const srcZip = makeSourceZip_();
  const obj = uploadToGCS_(p, bucket, 'src/' + Date.now() + '/' + srcZip.name, srcZip.bytes);

  // Cloud Build: gcloud run deploy --source .
  const buildId = startCloudBuildDeploy_(p, r, obj.bucket, obj.name, s, CONFIG.PUBLIC, CONFIG.RUN_SA);
  Logger.log('Cloud Build started: ' + buildId);

  // Poll build (re‑run if GAS times out)
  waitForBuild_(p, buildId, 300);

  // Get Cloud Run URL and store
  const url = getRunUrl_(p, r, s);
  PropertiesService.getScriptProperties().setProperty('CF_BASE', url);
  Logger.log('CF_BASE = ' + url);
}

// ===== GAS callers =====

function callService_unwrap(fileIdOrName) {
  const base = mustProp_('CF_BASE');
  const file = getFileByIdOrName_(fileIdOrName);
  const zipB64 = Utilities.base64Encode(file.getBlob().getBytes());
  const headers = maybeAuthHeader_(base);
  const resp = UrlFetchApp.fetch(base + '/unwrap', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ zipB64 }),
    headers,
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) throw new Error(resp.getContentText());
  return JSON.parse(resp.getContentText());
}

function callService_rewrap(manifest, outName) {
  const base = mustProp_('CF_BASE');
  const headers = maybeAuthHeader_(base);
  const resp = UrlFetchApp.fetch(base + '/rewrap', {
    method: 'post', contentType: 'application/json', headers,
    payload: JSON.stringify({ manifest })
  });
  if (resp.getResponseCode() >= 300) throw new Error(resp.getContentText());
  const data = JSON.parse(resp.getContentText());
  const bytes = Utilities.base64Decode(data.zipB64);
  const mime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  return DriveApp.createFile(Utilities.newBlob(bytes, mime, outName || 'out.pptx')).getId();
}

function callService_process(fileIdOrName, ops, outName) {
  const base = mustProp_('CF_BASE');
  const headers = maybeAuthHeader_(base);
  const file = getFileByIdOrName_(fileIdOrName);
  const zipB64 = Utilities.base64Encode(file.getBlob().getBytes());
  const resp = UrlFetchApp.fetch(base + '/process', {
    method: 'post', contentType: 'application/json', headers,
    payload: JSON.stringify({ zipB64: zipB64, ops: ops || [] }),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) throw new Error(resp.getContentText());
  const data = JSON.parse(resp.getContentText());
  const bytes = data.zipB64 ? Utilities.base64Decode(data.zipB64) : null;
  const mime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  const id = bytes ? DriveApp.createFile(Utilities.newBlob(bytes, mime, outName || 'out.pptx')).getId() : null;
  Logger.log(JSON.stringify(data.report || {}, null, 2));
  return id || data; // if gcsOut flow used
}

// ===== Large‑file session helpers (GCS signed URLs) =====

function helper_newSession() {
  const base = mustProp_('CF_BASE');
  const headers = maybeAuthHeader_(base);
  const r = UrlFetchApp.fetch(base + '/session/new', { method: 'post', payload: '{}', headers });
  return JSON.parse(r.getContentText());
}

function helper_uploadToSignedUrl(uploadUrl, fileIdOrName) {
  const file = getFileByIdOrName_(fileIdOrName);
  const resp = UrlFetchApp.fetch(uploadUrl, {
    method: 'put',
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    payload: file.getBlob().getBytes(),
    followRedirects: true,
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) throw new Error(resp.getResponseCode() + ' ' + resp.getContentText());
}

function helper_unwrap_fromGCS(gcsIn) {
  const base = mustProp_('CF_BASE');
  const headers = maybeAuthHeader_(base);
  const resp = UrlFetchApp.fetch(base + '/unwrap', {
    method: 'post', contentType: 'application/json', headers,
    payload: JSON.stringify({ gcsIn: gcsIn })
  });
  if (resp.getResponseCode() >= 300) throw new Error(resp.getContentText());
  return JSON.parse(resp.getContentText());
}

function helper_rewrap_toGCS(manifest, gcsIn, gcsOut) {
  const base = mustProp_('CF_BASE');
  const headers = maybeAuthHeader_(base);
  const resp = UrlFetchApp.fetch(base + '/rewrap', {
    method: 'post', contentType: 'application/json', headers,
    payload: JSON.stringify({ manifest: manifest, gcsIn: gcsIn, gcsOut: gcsOut })
  });
  if (resp.getResponseCode() >= 300) throw new Error(resp.getContentText());
  return JSON.parse(resp.getContentText()); // { gcsOut } or { zipB64 }
}

function demo_session_roundtrip() {
  var ses = helper_newSession();
  helper_uploadToSignedUrl(ses.uploadUrl, 'in.pptx');
  var manifest = helper_unwrap_fromGCS(ses.gcsIn);
  manifest.entries.forEach(function(e){ if (e.type === 'xml') e.text = e.text.replace(/ACME/g, 'DeltaQuad'); });
  var out = helper_rewrap_toGCS(manifest, ses.gcsIn, ses.gcsOut);
  Logger.log(out.gcsOut || JSON.stringify(out));
}

// ===== Preflight UI (billing/APIs/budget) =====

function _preflightHtml_(projectId) {
  const apis = REQUIRED_APIS.map(a => `<li><code>${a}</code> — <span data-api="${a}">checking…</span></li>`).join('');
  return `
  <style>
    body{font:13px/1.4 system-ui,Segoe UI,Roboto,Arial;margin:12px}
    h3{margin:8px 0 6px} button{padding:8px 10px;border-radius:8px;border:1px solid #dadce0;background:#fff;cursor:pointer}
    button.primary{background:#0b57d0;color:#fff;border-color:#0b57d0} .ok{color:#137333}.bad{color:#c5221f}.warn{color:#b06000}
    input[type=text]{width:100%;padding:6px;border:1px solid #dadce0;border-radius:6px}
    .grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
  </style>
  <h3>GCP Preflight</h3>
  <div>Project: <code>${projectId}</code></div>
  <div id="billingRow">Billing: <span id="billingStatus">checking…</span></div>
  <h4>Required APIs</h4>
  <ul>${apis}</ul>
  <div><button id="enableApis">Enable missing APIs</button></div>
  <h4>Budget (emails to billing recipients)</h4>
  <div class="grid">
    <div><label>Name</label><input id="budgetName" type="text" value="${CONFIG.BUDGET_NAME}"></div>
    <div><label>Currency</label><input id="currency" type="text" value="${CONFIG.BUDGET_CURRENCY}"></div>
    <div><label>Amount</label><input id="amount" type="text" value="${CONFIG.BUDGET_AMOUNT_UNITS}"></div>
    <div><label>Thresholds</label><input id="thresholds" type="text" value="10,50,90"></div>
  </div>
  <div><button id="createBudget">Create/Update budget</button></div>
  <div><button class="primary" id="recheck">Re-check</button></div>
  <small id="msg"></small>
  <script>
    function setMsg(t, cls){ const m=document.getElementById('msg'); m.textContent=t||''; m.className=cls||''; }
    function setBilling(s, cls){ const el=document.getElementById('billingStatus'); el.textContent=s; el.className=cls||''; }
    function setApi(a, s, cls){ const el=document.querySelector('[data-api="'+a+'"]'); if(el){ el.textContent=s; el.className=cls||''; } }
    function refresh(){
      setMsg(''); setBilling('checking…','');
      google.script.run.withSuccessHandler(function(ok){ setBilling(ok.enabled ? 'ENABLED ('+ok.accountId+')' : 'NOT ENABLED', ok.enabled ? 'ok':'bad'); })._pf_checkBilling();
      ${JSON.stringify(REQUIRED_APIS)}.forEach(function(api){ setApi(api,'checking…','');
        google.script.run.withSuccessHandler(function(st){ setApi(api, st.enabled ? 'ENABLED' : 'DISABLED', st.enabled ? 'ok':'bad'); })._pf_checkApi(api);
      });
    }
    document.getElementById('enableApis').onclick = function(){ setMsg('Enabling APIs…');
      google.script.run.withSuccessHandler(function(){ setMsg('APIs enabled. Some take ~1–2 min to settle.','ok'); refresh(); })._pf_enableRequiredApis(); };
    document.getElementById('createBudget').onclick = function(){
      const name=document.getElementById('budgetName').value.trim();
      const currency=document.getElementById('currency').value.trim();
      const amount=document.getElementById('amount').value.trim();
      const ths=document.getElementById('thresholds').value.split(',').map(s=>parseFloat(s)/100).filter(n=>!isNaN(n));
      setMsg('Ensuring budget…');
      google.script.run.withSuccessHandler(function(info){ setMsg('Budget ready: '+(info.name||name),'ok'); })
        .withFailureHandler(function(err){ setMsg(err && err.message ? err.message : String(err),'bad'); })
        ._pf_ensureBudget(name,currency,amount,ths);
    };
    document.getElementById('recheck').onclick = refresh; refresh();
  </script>`;
}

function _pf_checkBilling() {
  const projectId = CONFIG.PROJECT_ID;
  const info = gapi_(`https://cloudbilling.googleapis.com/v1/projects/${projectId}/billingInfo`, 'get');
  const enabled = !!info.billingEnabled; const acct = info.billingAccountName || '';
  return { enabled: enabled, accountName: acct, accountId: acct.replace('billingAccounts/','') || null };
}
function _pf_checkApi(api) {
  const projectId = CONFIG.PROJECT_ID;
  const svc = gapi_(`https://serviceusage.googleapis.com/v1/projects/${projectId}/services/${api}`, 'get', null, true);
  return { api, enabled: !!svc && svc.state === 'ENABLED' };
}
function _pf_enableRequiredApis() {
  const projectId = CONFIG.PROJECT_ID; const body = { serviceIds: REQUIRED_APIS };
  return gapi_(`https://serviceusage.googleapis.com/v1/projects/${projectId}/services:batchEnable`, 'post', body);
}
function _pf_getProjectNumber() {
  const projectId = CONFIG.PROJECT_ID;
  const r = gapi_(`https://cloudresourcemanager.googleapis.com/v1/projects/${projectId}`, 'get');
  return r.projectNumber;
}
function _pf_listBudgets_(billingAccountName) {
  const out = []; let pageToken = '';
  do {
    const url = `https://billingbudgets.googleapis.com/v1/${billingAccountName}/budgets` + (pageToken ? `?pageToken=${encodeURIComponent(pageToken)}` : '');
    const r = gapi_(url, 'get', null, true) || {}; (r.budgets || []).forEach(b => out.push(b));
    pageToken = r.nextPageToken || '';
  } while (pageToken);
  return out;
}
function _pf_ensureBudget(displayName, currencyCode, amountUnits, thresholds) {
  const bill = _pf_checkBilling(); if (!bill.enabled) throw new Error('Billing not enabled on project.');
  const billingAccountName = bill.accountName; const projectNumber = _pf_getProjectNumber();
  const budget = {
    displayName: displayName || CONFIG.BUDGET_NAME,
    budgetFilter: { projects: [`projects/${projectNumber}`] },
    amount: { specifiedAmount: { currencyCode: currencyCode || CONFIG.BUDGET_CURRENCY, units: String(amountUnits || CONFIG.BUDGET_AMOUNT_UNITS) } },
    thresholdRules: (thresholds && thresholds.length ? thresholds : CONFIG.BUDGET_THRESHOLDS).map(p => ({ thresholdPercent: p })),
    allUpdatesRule: { disableDefaultIamRecipients: false }
  };
  const existing = _pf_listBudgets_(billingAccountName).find(b => b.displayName === budget.displayName);
  if (existing) {
    const name = existing.name; const mask = 'displayName,budgetFilter,amount,thresholdRules,allUpdatesRule';
    const url = `https://billingbudgets.googleapis.com/v1/${name}?updateMask=${encodeURIComponent(mask)}`;
    return gapi_(url, 'patch', budget);
  } else {
    const url = `https://billingbudgets.googleapis.com/v1/${billingAccountName}/budgets`;
    return gapi_(url, 'post', budget);
  }
}

// ===== GCP REST utils =====

function enableApis_(projectId, services) {
  const body = { serviceIds: services };
  gapi_('https://serviceusage.googleapis.com/v1/projects/' + projectId + '/services:batchEnable', 'post', body);
}
function ensureBucket_(projectId, bucketName, location) {
  const exists = gapi_('https://storage.googleapis.com/storage/v1/b/' + encodeURIComponent(bucketName), 'get', null, true);
  if (exists) return;
  const body = { name: bucketName, location: location, iamConfiguration: { uniformBucketLevelAccess: { enabled: true } } };
  gapi_('https://storage.googleapis.com/storage/v1/b?project=' + encodeURIComponent(projectId), 'post', body);
}
function uploadToGCS_(projectId, bucket, objectName, bytes) {
  const url = 'https://storage.googleapis.com/upload/storage/v1/b/' + encodeURIComponent(bucket) + '/o?uploadType=media&name=' + encodeURIComponent(objectName);
  const resp = gapiRaw_(url, 'post', bytes, { 'Content-Type': 'application/zip' });
  const meta = JSON.parse(resp); return { bucket: meta.bucket, name: meta.name };
}
function startCloudBuildDeploy_(projectId, region, bucket, objectName, service, allowUnauth, runSaEmail) {
  const cbUrl = 'https://cloudbuild.googleapis.com/v1/projects/' + projectId + '/builds';
  const args = ['run','deploy', service, '--source','.', '--region', region, '--project', projectId, '--quiet'];
  if (allowUnauth) args.push('--allow-unauthenticated');
  if (runSaEmail) { args.push('--service-account', runSaEmail); }
  // Set env vars for the service
  args.push('--set-env-vars', 'BUCKET=' + (CONFIG.BUCKET || (projectId + '-ooxml-src')) + ',REGION=' + region);
  const body = { source: { storageSource: { bucket: bucket, object: objectName } }, steps: [ { name: 'gcr.io/cloud-builders/gcloud', args: args } ], options: { logging: 'CLOUD_LOGGING_ONLY' } };
  const r = gapi_(cbUrl, 'post', body); return r.id;
}
function waitForBuild_(projectId, buildId, timeoutSec) {
  const url = 'https://cloudbuild.googleapis.com/v1/projects/' + projectId + '/builds/' + buildId;
  const t0 = Date.now();
  while (true) {
    const b = gapi_(url, 'get', null);
    if (b.logUrl) Logger.log('Logs: ' + b.logUrl);
    if (['SUCCESS','FAILURE','CANCELLED','TIMEOUT'].indexOf(b.status) >= 0) {
      if (b.status !== 'SUCCESS') throw new Error('Build ' + b.status);
      return;
    }
    if ((Date.now() - t0) / 1000 > (timeoutSec || 300)) { Logger.log('Build still running; re-run waitForBuild_ later.'); return; }
    Utilities.sleep(4000);
  }
}
function getRunUrl_(projectId, region, service) {
  const url = 'https://run.googleapis.com/v2/projects/' + projectId + '/locations/' + region + '/services/' + service;
  const r = gapi_(url, 'get', null); return r.uri;
}

function gapi_(url, method, body, returnNullOn404) {
  const resp = UrlFetchApp.fetch(url, { method: method || 'get', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'Content-Type': 'application/json' }, payload: body ? JSON.stringify(body) : null, muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (returnNullOn404 && code === 404) return null;
  if (code < 200 || code >= 300) throw new Error(method + ' ' + url + ' => ' + code + ' ' + resp.getContentText());
  const text = resp.getContentText(); return text ? JSON.parse(text) : {};
}
function gapiRaw_(url, method, bytes, extraHeaders) {
  const headers = Object.assign({ Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, extraHeaders || {});
  const resp = UrlFetchApp.fetch(url, { method: method || 'post', headers: headers, payload: bytes, muteHttpExceptions: true });
  const code = resp.getResponseCode(); if (code < 200 || code >= 300) throw new Error('RAW ' + url + ' => ' + code + ' ' + resp.getContentText());
  return resp.getContentText();
}

// ===== Service source packaging =====

function makeSourceZip_() {
  const indexMjs = SERVICE_INDEX_MJS_();
  const pkg = SERVICE_PACKAGE_JSON_();
  const gignore = 'node_modules\n.git\n';
  const files = [ Utilities.newBlob(indexMjs, 'text/plain', 'index.mjs'), Utilities.newBlob(pkg, 'application/json', 'package.json'), Utilities.newBlob(gignore, 'text/plain', '.gcloudignore'), Utilities.newBlob('# OOXML JSON Service\n', 'text/markdown', 'README.md') ];
  const zip = Utilities.zip(files, 'ooxml-json-src.zip');
  return { name: zip.getName(), bytes: zip.getBytes() };
}

function SERVICE_PACKAGE_JSON_() {
  return JSON.stringify({
    type: 'module',
    dependencies: { '@google-cloud/storage': '^7.12.1', express: '^4.19.2', fflate: '^0.8.2' },
    engines: { node: '>=18' },
    scripts: { start: 'node index.mjs' }
  }, null, 2);
}

function SERVICE_INDEX_MJS_() {
  return `import express from "express";
import { unzipSync, zipSync, base64Decode, base64Encode } from "fflate";
import { Storage } from "@google-cloud/storage";
import crypto from "node:crypto";

const app = express();
app.use(express.json({ limit: "100mb" }));
const storage = new Storage();
const BUCKET = process.env.BUCKET;               // set at deploy
const REGION  = process.env.REGION || "europe-west4";

const toU8  = (b64) => base64Decode(b64.replace(/-/g, "+").replace(/_/g, "/"));
const toB64 = (u8)  => base64Encode(u8);
const isXml = (p) => /\.xml$/i.test(p);

// ---- sessions --------------------------------------------------------------
app.post("/session/new", async (req, res) => {
  try {
    if (!BUCKET) return res.status(500).json({ error: "BUCKET not set" });
    const sid = crypto.randomUUID();
    const inKey  = `${sid}/in.pptx`;
    const outKey = `${sid}/out.pptx`;

    const input = storage.bucket(BUCKET).file(inKey);
    const output = storage.bucket(BUCKET).file(outKey);

    const [uploadUrl] = await input.getSignedUrl({ version: "v4", action: "write", expires: Date.now() + 15*60*1000, contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
    const [downloadUrl] = await output.getSignedUrl({ version: "v4", action: "read", expires: Date.now() + 60*60*1000 });

    res.json({ sessionId: sid, gcsIn: `gs://${BUCKET}/${inKey}`, gcsOut: `gs://${BUCKET}/${outKey}`, uploadUrl, downloadUrl });
  } catch (e) { res.status(500).json({ error: String(e?.message || e) }); }
});

// ---- unwrap ---------------------------------------------------------------
app.post("/unwrap", async (req, res) => {
  try {
    const { zipB64, gcsIn } = req.body || {}; let buf;
    if (zipB64) buf = toU8(zipB64);
    else if (gcsIn) { const [, , bucket, ...rest] = gcsIn.split("/"); const key = rest.join("/"); const [data] = await storage.bucket(bucket).file(key).download(); buf = data; }
    else return res.status(400).json({ error: "zipB64 or gcsIn required" });

    const bag = unzipSync(buf, { filter: () => true });
    const paths = Object.keys(bag).sort();
    const entries = paths.map((path) => isXml(path) ? { path, type: "xml", text: new TextDecoder().decode(bag[path]) } : { path, type: "bin" });
    res.json({ version: "ooxml-json-1", kind: detectKind(entries), entries });
  } catch (e) { res.status(500).json({ error: String(e?.message || e) }); }
});

// ---- rewrap ---------------------------------------------------------------
app.post("/rewrap", async (req, res) => {
  try {
    const { manifest, gcsIn, gcsOut } = req.body || {};
    if (!manifest || !Array.isArray(manifest.entries)) return res.status(400).json({ error: "manifest.entries missing" });
    if (!gcsIn && !manifest.zipB64) return res.status(400).json({ error: "gcsIn or manifest.zipB64 required" });

    let raw;
    if (gcsIn) { const [, , bucket, ...rest] = gcsIn.split("/"); const key = rest.join("/"); const [data] = await storage.bucket(bucket).file(key).download(); raw = data; }
    else { raw = toU8(manifest.zipB64); }

    const bag = unzipSync(raw, { filter: () => true });
    const enc = new TextEncoder();
    for (const e of manifest.entries) {
      if (!e || !e.path) continue;
      if (e.type === "xml" && typeof e.text === "string") bag[e.path] = enc.encode(e.text);
      else if (e.type === "bin" && e.dataB64) bag[e.path] = toU8(e.dataB64);
    }
    const ordered = Object.fromEntries(Object.keys(bag).sort().map(p => [p, bag[p]]));
    const outZip = zipSync(ordered, { level: 6, zip64: true });

    if (gcsOut) {
      const [, , outBucket, ...outRest] = gcsOut.split("/"); const outKey = outRest.join("/");
      await storage.bucket(outBucket).file(outKey).save(Buffer.from(outZip), { contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
      return res.json({ gcsOut });
    }
    res.json({ zipB64: base64Encode(outZip) });
  } catch (e) { res.status(500).json({ error: String(e?.message || e) }); }
});

// ---- process (server-side ops) -------------------------------------------
app.post("/process", async (req, res) => {
  try {
    const { zipB64, gcsIn, gcsOut, ops } = req.body || {};
    if (!Array.isArray(ops)) return res.status(400).json({ error: "ops[] required" });
    if (!zipB64 && !gcsIn) return res.status(400).json({ error: "zipB64 or gcsIn required" });

    let buf;
    if (zipB64) buf = toU8(zipB64); else { const [, , bucket, ...rest] = gcsIn.split("/"); const key = rest.join("/"); const [data] = await storage.bucket(bucket).file(key).download(); buf = data; }

    const bag = unzipSync(buf, { filter: () => true });
    const enc = new TextEncoder(), dec = new TextDecoder();
    const getText = (p) => dec.decode(bag[p]);
    const setText = (p, s) => { bag[p] = enc.encode(s); };
    const setBin  = (p, b64) => { bag[p] = toU8(b64); };
    const del     = (p) => { delete bag[p]; };
    const exists  = (p) => Object.prototype.hasOwnProperty.call(bag, p);

    const ctPath = "[Content_Types].xml";
    let ctXml = exists(ctPath) ? getText(ctPath) : `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`;
    const report = { replaced: 0, upserted: 0, removed: 0, renamed: 0, errors: [] };

    for (const op of ops) {
      try {
        switch (op.type) {
          case "replaceText": {
            const scope = op.scope || ""; // prefix
            const re = op.regex ? new RegExp(op.find, op.flags || "g") : new RegExp(String(op.find).replace(/[.*+?^${}()|[\\]\\\\]/g, "\\$&"), "g");
            for (const p of Object.keys(bag)) {
              if (!/\.xml$/i.test(p)) continue; if (scope && !p.startsWith(scope)) continue;
              const before = getText(p); const after = before.replace(re, String(op.replace ?? ""));
              if (after !== before) { setText(p, after); report.replaced++; }
            }
            break;
          }
          case "upsertPart": {
            if (!op.path) throw new Error("upsertPart.path missing");
            if (/\.xml$/i.test(op.path)) { setText(op.path, String(op.text ?? "")); if (op.contentType) ctXml = ensureOverride(ctXml, op.path, op.contentType); }
            else { if (!op.dataB64) throw new Error("upsertPart.dataB64 missing for binary"); setBin(op.path, op.dataB64); if (op.contentType) ctXml = ensureOverride(ctXml, op.path, op.contentType); }
            report.upserted++; break;
          }
          case "removePart": { if (!op.path) throw new Error("removePart.path missing"); if (exists(op.path)) { del(op.path); report.removed++; } ctXml = removeOverride(ctXml, op.path); break; }
          case "renamePart": { if (!op.from || !op.to) throw new Error("renamePart.from/to missing"); if (exists(op.from)) { bag[op.to]=bag[op.from]; delete bag[op.from]; if (op.contentType) { ctXml = removeOverride(ctXml, op.from); ctXml = ensureOverride(ctXml, op.to, op.contentType); } report.renamed++; } break; }
          default: throw new Error("unknown op.type: " + op.type);
        }
      } catch (e) { report.errors.push({ op, message: String(e?.message || e) }); }
    }

    bag[ctPath] = enc.encode(ctXml);
    const ordered = Object.fromEntries(Object.keys(bag).sort().map(p=>[p,bag[p]]));
    const outZip = zipSync(ordered, { level: 6, zip64: true });

    if (gcsOut) { const [, , b, ...rest] = gcsOut.split("/"); const k = rest.join("/"); await storage.bucket(b).file(k).save(Buffer.from(outZip), { contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }); return res.json({ gcsOut, report }); }
    res.json({ zipB64: toB64(outZip), report });
  } catch (e) { res.status(500).json({ error: String(e?.message || e) }); }
});

app.get("/ping", (_req,res)=>res.json({ ok:true, version:"ooxml-json-1" }));

function detectKind(entries) {
  const ct = entries.find(e => e.path === "[Content_Types].xml" && e.type === "xml"); if (!ct) return "unknown";
  const has = (s) => ct.text.includes(`ContentType="${s}"`);
  if (has("application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")) return "pptx";
  if (has("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"))     return "docx";
  if (has("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"))          return "xlsx";
  if (has("application/vnd.openxmlformats-officedocument.theme+xml"))                               return "thmx";
  return "unknown";
}

function ensureOverride(ctXml, partName, contentType) {
  const pn = partName.startsWith('/') ? partName : '/' + partName;
  const rx = new RegExp(`<Override\\b[^>]*PartName="${pn.replace(/[.*+?^${}()|[\\]\\\\]/g,'\\$&')}"[^>]*/>`, 'i');
  if (rx.test(ctXml)) return ctXml;
  return ctXml.replace(/<\/Types>\s*$/i, `<Override PartName="${pn}" ContentType="${contentType}"/><\/Types>`);
}
function removeOverride(ctXml, partName) {
  const pn = partName.startsWith('/') ? partName : '/' + partName;
  const rx = new RegExp(`<Override\\b[^>]*PartName="${pn.replace(/[.*+?^${}()|[\\]\\\\]/g,'\\$&')}"[^>]*/>\\s*`, 'i');
  return ctXml.replace(rx, '');
}

const port = process.env.PORT || 8080;
app.listen(port, () => console.log("listening on", port));
`;
}

// ===== Utilities =====

function getFileByIdOrName_(idOrName) {
  if (/^[a-zA-Z0-9\-_]{20,}$/.test(idOrName)) return DriveApp.getFileById(idOrName);
  const it = DriveApp.getFilesByName(idOrName); if (!it.hasNext()) throw new Error('File not found: ' + idOrName);
  return it.next();
}
function mustProp_(k) { const v = PropertiesService.getScriptProperties().getProperty(k); if (!v) throw new Error('Missing Script Property ' + k); return v; }
function maybeAuthHeader_(base) { return CONFIG.PUBLIC ? {} : { Authorization: 'Bearer ' + getIdTokenForAudience_(base) }; }
function getIdTokenForAudience_(aud) {
  // Uses a helper service account with Token Creator; kept minimal here.
  // In many setups, default credentials on Cloud Run suffice; for GAS, minting an ID token
  // requires an external broker. Keep service PUBLIC=true to avoid that complexity.
  throw new Error('ID token minting from GAS not wired. Set CONFIG.PUBLIC=true or integrate an external broker.');
}
```

> Note: If you need the service **private**, set `PUBLIC: false` and wire an ID‑token flow that fits your environment (IAP, OAuth2 library, or a backend broker). For most internal tooling, public + secret headers + unguessable endpoints are pragmatic; your call.

---

## 2) Cloud Service – Source (Standalone View)

You don’t need this separately — GAS packages it for you — but here for reference.

### `index.mjs`

```js
// Same as SERVICE_INDEX_MJS_() generated by GAS (see above).
// Contains: /session/new, /unwrap, /rewrap, /process endpoints.
```

### `package.json`

```json
{
  "type": "module",
  "dependencies": {
    "@google-cloud/storage": "^7.12.1",
    "express": "^4.19.2",
    "fflate": "^0.8.2"
  },
  "engines": { "node": ">=18" },
  "scripts": { "start": "node index.mjs" }
}
```

Deploy manually if you like:

```bash
# From the source directory
gcloud run deploy ooxml-json --source=. --region=europe-west4 --allow-unauthenticated \
  --set-env-vars BUCKET=<your-bucket>,REGION=europe-west4
```

---

## 3) API Reference (Short)

### `POST /session/new`

**Resp** `{ sessionId, gcsIn, gcsOut, uploadUrl, downloadUrl }`

### `POST /unwrap`

**Req** `{ zipB64 }` **or** `{ gcsIn }` → **Resp** `{ version, kind, entries[] }`

- XML parts: `{ path, type:"xml", text }`
- Binary parts: `{ path, type:"bin" }` (no payload; keeps manifests lean)

### `POST /rewrap`

**Req** `{ manifest, gcsIn?, gcsOut? }` *(or **`{ manifest, manifest.zipB64 }`**)*

- Applies `manifest.entries` over original zip (from GCS or inline).
- **Resp** `{ gcsOut }` if provided, else `{ zipB64 }`.

### `POST /process`

**Req** `{ zipB64|gcsIn, gcsOut?, ops:[ ... ] }` Supported ops:

- `replaceText` `{ type, scope?, find, replace, regex?, flags? }`
- `upsertPart`  `{ type, path, text? | dataB64?, contentType? }`
- `removePart`  `{ type, path }`
- `renamePart`  `{ type, from, to, contentType? }` **Resp** `{ zipB64|gcsOut, report }`

---

## 4) Typical Flows

### A. Small file (inline base64)

```javascript
const m = callService_unwrap('in.pptx');
m.entries.forEach(e => { if (e.type === 'xml') e.text = e.text.replace(/ACME/g,'DeltaQuad'); });
const outId = callService_rewrap(m, 'out.pptx');
```

### B. Large file (100MB)

```javascript
const ses = helper_newSession();
helper_uploadToSignedUrl(ses.uploadUrl, 'in.pptx');
const m = helper_unwrap_fromGCS(ses.gcsIn);
// mutate XML only
const res = helper_rewrap_toGCS(m, ses.gcsIn, ses.gcsOut);
Logger.log(res.gcsOut);
```

### C. One‑shot server‑side ops

```javascript
const ops = [
  { type: 'replaceText', scope: 'ppt/slides/', find: 'ACME', replace: 'DeltaQuad' },
  { type: 'upsertPart', path: 'ppt/customXml/item1.xml', text: '<dq/>' }
];
callService_process('in.pptx', ops, 'out.pptx');
```

---

## 5) Security & Cost Notes

- **Billing required** (Cloud Run/Build/Storage). Preflight checks guide users.
- Keep **PUBLIC=true** to avoid ID‑token plumbing from GAS. If you must lock it down, front with IAP or a broker.
- **Budget** is set via the sidebar; default €5 with 10/50/90% alerts.
- For free‑tier friendliness, deploy in *us‑** regions*\* if acceptable; otherwise cost is still low.
- Add a GCS **lifecycle rule** to auto‑delete `*/in.pptx` & `*/out.pptx` after 1–24h.

---

## 6) Troubleshooting (fast)

- **403 on API calls**: run `showGcpPreflight()` → Enable APIs → Re‑check.
- **Billing not enabled**: link a billing account in Console → Re‑check.
- **Build timeouts**: rerun `waitForBuild_` with the same build ID (or rerun `initAndDeploy`).
- **Huge files via inline**: use session flow; don’t POST 100MB base64 from GAS.
- **Office “repair”**: you edited content types wrong. Use `contentType` in `upsertPart`/`renamePart` ops.

---

## 7) License & Attribution

Use it. Ship it. Responsibility is yours. No warranties.

