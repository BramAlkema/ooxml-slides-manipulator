# OOXML JSON Helper — Contracts, Error Codes, DSL & Acid Test (Artifact)

Ship‑ready addendum to the base **OOXML JSON Helper**. This pins **hard API contracts**, a **routeable error taxonomy**, a **wire format** (don’t drift), a **Brandbook Rules DSL** (declarative + autofix), and an **acid test** that produces artefacts.

> Goal: fewer surprises in prod; faster debugging; repeatable compliance checks.

---

## 0) Versioning & Wire Guardrails

- **Wire version:** `ooxml-json-1`. Bump when breaking the manifest/wire schemas; expose via `/ping.version` and in every response as `version`.
- **Compat rule:** Minor additions (new fields) are fine. Removing/renaming fields → bump.
- **Correlation:** Accept/return `x-correlation-id` header and echo it in JSON (`correlation`). Log it everywhere.

---

## 1) Type Contracts (Layer Interfaces)

### 1.1 Shared Types (TypeScript)
```ts
// types/contracts.d.ts
export type OOXMLKind = 'pptx'|'docx'|'xlsx'|'thmx'|'unknown';

export type PartEntry =
  | { path: string; type: 'xml'; text: string }
  | { path: string; type: 'bin'; dataB64?: string }; // optional to keep manifests lean

export interface Manifest {
  version: 'ooxml-json-1';
  kind: OOXMLKind;
  entries: PartEntry[];
}

export interface UnwrapReq { zipB64?: string; gcsIn?: string }
export interface UnwrapRes { version: 'ooxml-json-1'; kind: OOXMLKind; entries: PartEntry[]; correlation?: string }

export interface RewrapReq { manifest: Manifest; gcsIn?: string; gcsOut?: string }
export interface RewrapRes { version: 'ooxml-json-1'; zipB64?: string; gcsOut?: string; correlation?: string }

export type ProcessOp =
  | { type: 'replaceText'; scope?: string; find: string; replace: string; regex?: boolean; flags?: string }
  | { type: 'upsertPart'; path: string; text?: string; dataB64?: string; contentType?: string }
  | { type: 'removePart'; path: string }
  | { type: 'renamePart'; from: string; to: string; contentType?: string };

export interface ProcessReq { zipB64?: string; gcsIn?: string; gcsOut?: string; ops: ProcessOp[] }
export interface ProcessRes {
  version: 'ooxml-json-1';
  zipB64?: string; gcsOut?: string;
  report: { replaced: number; upserted: number; removed: number; renamed: number; errors: { op: any; message: string }[] };
  correlation?: string;
}

export class OOXMLError extends Error { constructor(public code: string, msg: string, public ctx: Record<string, any> = {}){ super(msg); } }
```

### 1.2 Core (OOXMLCore) contract (JSDoc for GAS)
```js
/**
 * @typedef {Object} OOXMLDoc
 * @property {function(): string[]} listPaths
 * @property {function(string): string} getXml
 * @property {function(string, string): void} setXml
 * @property {function(string): Uint8Array} getBinary
 * @property {function(string, Uint8Array): void} setBinary
 * @property {function(): Promise<Uint8Array>} saveZip
 */
```

### 1.3 Service bridge (CloudPPTXService)
```js
/** Cloud bridge: unzip/zip/process with retry & fallback */
export const CloudPPTXService = {
  /** @returns {Promise<UnwrapRes>} */
  unwrap(zipBytesOrGcs){},
  /** @returns {Promise<RewrapRes>} */
  rewrap(req){},
  /** @returns {Promise<ProcessRes>} */
  process(req){},
  /** health check */
  ping(){}
};
```

---

## 2) Error Code Taxonomy (grep‑able)

| Range | Domain              | Examples                                                  |
|------:|---------------------|-----------------------------------------------------------|
| C00x  | Core ZIP/XML        | **C001** bad_zip, **C002** missing_part, **C003** xml_parse |
| S01x  | Service/Network     | **S010** http_4xx, **S011** http_5xx, **S012** timeout, **S013** retry_exhausted |
| A02x  | App/Slides semantics| **A020** theme_not_found, **A021** layout_map_missing     |
| E03x  | Extensions          | **E030** ext_missing, **E031** ext_failed, **E032** bad_config |
| V04x  | Validation          | **V040** rule_parse, **V041** rule_exec                   |

Emit logs in a trivially grep‑able shape:
```
ERR[C001] bad_zip ctx={"url":"/unwrap","size":104857600,"corr":"ab12"}
```

Throw with context:
```js
throw new OOXMLError('C002', 'missing_part', { path: 'ppt/slides/slide1.xml', corr });
```

---

## 3) Wire Schemas (don’t drift)

### 3.1 OpenAPI (excerpt)
```yaml
openapi: 3.0.3
info: { title: OOXML JSON Service, version: 1.0.0 }
paths:
  /ping:
    get: { responses: { '200': { description: OK, content: { application/json: { schema: { type: object, properties: { ok: {type: boolean}, version: {type: string} } } } } } } }
  /session/new:
    post: { responses: { '200': { description: Session, content: { application/json: { schema: { type: object, properties: { sessionId:{type:string}, gcsIn:{type:string}, gcsOut:{type:string}, uploadUrl:{type:string}, downloadUrl:{type:string} } } } } } } }
  /unwrap:
    post:
      requestBody: { required: true, content: { application/json: { schema: { oneOf: [ {type: object, properties: {zipB64:{type:string}}, required:[zipB64]}, {type: object, properties: {gcsIn:{type:string}}, required:[gcsIn]} ] } } } }
      responses: { '200': { description: Manifest, content: { application/json: { schema: { $ref: '#/components/schemas/Manifest' } } } } }
  /rewrap:
    post:
      requestBody: { required: true, content: { application/json: { schema: { type: object, properties: { manifest: { $ref: '#/components/schemas/Manifest' }, gcsIn:{type:string}, gcsOut:{type:string} }, required: [manifest] } } } }
      responses: { '200': { description: OK, content: { application/json: { schema: { type: object, properties: { zipB64:{type:string}, gcsOut:{type:string} } } } } } }
  /process:
    post:
      requestBody: { required: true, content: { application/json: { schema: { $ref: '#/components/schemas/ProcessReq' } } } }
      responses: { '200': { description: OK, content: { application/json: { schema: { $ref: '#/components/schemas/ProcessRes' } } } } }
components:
  schemas:
    PartEntry:
      oneOf:
        - { type: object, required:[path,type,text], properties:{ path:{type:string}, type:{const: xml}, text:{type:string} } }
        - { type: object, required:[path,type], properties:{ path:{type:string}, type:{const: bin}, dataB64:{type:string} } }
    Manifest:
      type: object
      required: [version, kind, entries]
      properties:
        version: { type: string }
        kind: { type: string, enum: [pptx,docx,xlsx,thmx,unknown] }
        entries: { type: array, items: { $ref: '#/components/schemas/PartEntry' } }
    ProcessReq:
      type: object
      properties:
        zipB64: { type: string }
        gcsIn:  { type: string }
        gcsOut: { type: string }
        ops:
          type: array
          items:
            oneOf:
              - { type: object, required:[type,find,replace], properties:{ type:{const: replaceText}, scope:{type:string}, find:{type:string}, replace:{type:string}, regex:{type:boolean}, flags:{type:string} } }
              - { type: object, required:[type,path], properties:{ type:{const: upsertPart}, path:{type:string}, text:{type:string}, dataB64:{type:string}, contentType:{type:string} } }
              - { type: object, required:[type,path], properties:{ type:{const: removePart}, path:{type:string} } }
              - { type: object, required:[type,from,to], properties:{ type:{const: renamePart}, from:{type:string}, to:{type:string}, contentType:{type:string} } }
    ProcessRes:
      type: object
      properties:
        version: { type: string }
        zipB64: { type: string }
        gcsOut: { type: string }
        report:
          type: object
          properties:
            replaced: { type: integer }
            upserted: { type: integer }
            removed: { type: integer }
            renamed: { type: integer }
            errors: { type: array, items: { type: object } }
```

---

## 4) Brandbook Rules DSL (declarative)

### 4.1 JSON Schema
```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "$id": "https://example.com/brand-rules.schema.json",
  "type": "object",
  "properties": {
    "profile": { "type": "string" },
    "rules": { "type": "array", "items": { "$ref": "#/definitions/rule" } }
  },
  "required": ["rules"],
  "definitions": {
    "rule": {
      "type": "object",
      "properties": {
        "id": { "type": "string" },
        "desc": { "type": "string" },
        "where": { "type": "string" },
        "xpath": { "type": "string" },
        "expect": { "type": "object" },
        "autofix": { "type": "boolean" },
        "weight": { "type": "number", "minimum": 0 }
      },
      "required": ["id","where","xpath","expect"]
    }
  }
}
```

### 4.2 Example rules
```json
{
  "profile": "default",
  "rules": [
    {
      "id": "brand.colours.accent1",
      "desc": "Theme accent1 must equal #005BBB",
      "where": "ppt/theme/theme1.xml",
      "xpath": "/*[local-name()='theme']/*[local-name()='themeElements']/*[local-name()='clrScheme']/*[local-name()='accent1']/*",
      "expect": { "hex": "#005BBB" },
      "autofix": true,
      "weight": 5
    },
    {
      "id": "brand.fonts.major",
      "desc": "Major Latin font is Inter",
      "where": "ppt/theme/theme1.xml",
      "xpath": "//a:majorFont/a:latin/@typeface",
      "expect": { "equals": "Inter" },
      "autofix": false,
      "weight": 3
    }
  ]
}
```

### 4.3 Evaluation (outline)
- Load `where` → XML string.
- Evaluate XPath with **namespace‑agnostic** selectors (`local-name()`), or pre‑map ns.
- Extract actual value: `a:srgbClr/@val` → normalise to `#RRGGBB`.
- Compare to expectation (`hex`, `equals`, `regex`).
- If `autofix` true and mismatch → rewrite subtree to a clean `<a:srgbClr val="..."/>` (or the right attribute), record `before/after`.

### 4.4 Violation record
```ts
export type Violation = {
  ruleId: string; where: string; message: string;
  before?: string; after?: string; autoFixed: boolean; weight: number;
}
```

---

## 5) Acid Test (end‑to‑end)

**Input:** `BrandDeck.pptx` in Drive.  
**Output artefacts:** `BrandDeck.fixed.pptx`, `BrandDeck.report.json`, `screenshots/*.png`.

### 5.1 GAS runner (copy/paste)
```javascript
function runAcidTest() {
  // 1) Load input
  var file = DriveApp.getFilesByName('BrandDeck.pptx').next();
  var zipB64 = Utilities.base64Encode(file.getBlob().getBytes());

  // 2) Server-side ops (replace a token, add a custom part)
  var ops = [
    { type: 'replaceText', scope: 'ppt/slides/', find: 'ACME', replace: 'DeltaQuad' },
    { type: 'upsertPart', path: 'ppt/customXml/brand.json', text: JSON.stringify({ palette: 'default' }), contentType: 'application/json' }
  ];
  var base = PropertiesService.getScriptProperties().getProperty('CF_BASE');
  var resp = UrlFetchApp.fetch(base + '/process', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ zipB64: zipB64, ops: ops }), muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) throw new Error(resp.getContentText());
  var data = JSON.parse(resp.getContentText());

  // 3) Save fixed PPTX + report
  var outBytes = Utilities.base64Decode(data.zipB64);
  var outId = DriveApp.createFile(Utilities.newBlob(outBytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'BrandDeck.fixed.pptx')).getId();
  DriveApp.createFile(Utilities.newBlob(JSON.stringify(data.report, null, 2), 'application/json', 'BrandDeck.report.json'));
  Logger.log('Fixed PPTX: ' + outId);
  return data.report;
}
```

### 5.2 Visual verification (CI idea)
- Use **LibreOffice headless** to render PNGs from `BrandDeck.fixed.pptx`.
- Compare with baseline via pixel diff (tiny threshold).
- Store PNGs + `report.json` as build artefacts.

Example GitHub Action (sketch):
```yaml
name: acid-test
on: [push]
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Install LibreOffice
        run: sudo apt-get update && sudo apt-get install -y libreoffice-impress imagemagick
      - name: Convert PPTX to PNG
        run: |
          mkdir -p out
          libreoffice --headless --convert-to png --outdir out BrandDeck.fixed.pptx
      - name: Pixel diff (example)
        run: |
          compare -metric AE out/Slide1.png baseline/Slide1.png out/diff1.png || true
      - name: Upload artefacts
        uses: actions/upload-artifact@v4
        with: { name: acid-artifacts, path: out }
```

---

## 6) Observability & Hygiene

- **Correlation ID:**
  - If request header `x-correlation-id` present → echo in JSON `correlation` + logs.
  - If absent → generate UUID; return it.
- **Structured logs:** one line JSON per event. Example:
```json
{"level":"error","code":"S011","msg":"http_5xx","corr":"ab12","path":"/rewrap","latency_ms":231}
```
- **Deterministic ZIP:** sort paths, fixed deflate level, ZIP64 on.
- **Lifecycle:** GCS rule: delete session objects after 24h.

---

## 7) Minimal Server Hooks (errors + correlation)

### Express middleware (snippet)
```js
app.use((req,res,next)=>{
  const corr = (req.headers['x-correlation-id']||'').toString() || crypto.randomUUID();
  res.setHeader('x-correlation-id', corr);
  req.corr = corr; next();
});

function boom(code, msg, ctx, res, http=500){
  const err = { code, message: msg, ...ctx, correlation: res.getHeader('x-correlation-id') };
  console.error(`ERR[${code}] ${msg} ctx=${JSON.stringify(ctx||{})}`);
  res.status(http).json(err);
}
```

Use `boom('C001','bad_zip',{size:buf.length}, res, 400)` for clean failures.

---

## 8) Extension Framework – contract pins

```js
export class BaseExtension {
  static kind = 'VALIDATION'; // THEME | VALIDATION | CONTENT
  static id = 'com.example.base';
  static requires = [];
  /** @param {OOXMLDoc} doc @param {{profile?:string, now?:Date}} ctx */
  async execute(doc, ctx) { /* override */ }
}

export const ExtensionFramework = (()=>{
  const reg = new Map();
  return {
    register(id, cls, meta){ reg.set(id, { cls, meta }); },
    getByKind(kind){ return [...reg.values()].filter(x=>x.cls.kind===kind).map(x=>new x.cls()); }
  };
})();
```

---

## 9) Reports (stable shape)

```json
{
  "version": "ooxml-json-1",
  "score": 92,
  "violations": [
    { "ruleId": "brand.colours.accent1", "where": "ppt/theme/theme1.xml", "message": "expected #005BBB, got #2277CC", "autoFixed": true, "weight": 5 }
  ],
  "opsReport": { "replaced": 14, "upserted": 1, "removed": 0, "renamed": 0 },
  "correlation": "ab12"
}
```

---

## 10) Deployment notes (delta from base guide)

- Pass env: `BUCKET`, `REGION`.
- Keep **PUBLIC=true** unless you wire ID token minting from GAS.
- Pin deps: `fflate@^0.8.2`, Node 18.
- Budget: default €5 with 10/50/90% alerts (see preflight panel).

---

## 11) Fast Checklist (print this)

- [ ] Billing linked (+ budget & alerts)
- [ ] Required APIs enabled
- [ ] Cloud Run deployed; `CF_BASE` stored
- [ ] GCS lifecycle rule on session bucket
- [ ] Correlation IDs plumbed through
- [ ] Error codes greppable in logs
- [ ] Acid test green: fixed.pptx + report.json + screenshots
- [ ] Rules DSL validated in CI

---

**Done.** This is the contract & test spine. Build extensions on top and go ship. 

