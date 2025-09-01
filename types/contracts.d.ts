/**
 * OOXML JSON Service - TypeScript Type Contracts
 * 
 * PURPOSE:
 * Provides type-safe contracts for the OOXML JSON Service API.
 * These definitions ensure wire format compatibility and enable
 * proper IDE support and compile-time validation.
 * 
 * VERSION: ooxml-json-1
 * 
 * USAGE:
 * Import these types in TypeScript projects or use for JSDoc
 * type annotations in Google Apps Script.
 */

// Core OOXML Types
export type OOXMLKind = 'pptx' | 'docx' | 'xlsx' | 'thmx' | 'unknown';

export type PartEntry =
  | { path: string; type: 'xml'; text: string }
  | { path: string; type: 'bin'; dataB64?: string }; // optional to keep manifests lean

export interface Manifest {
  version: 'ooxml-json-1';
  kind: OOXMLKind;
  entries: PartEntry[];
  metadata?: ManifestMetadata;
}

export interface ManifestMetadata {
  originalSize?: number;
  extractedAt?: string;
  entryCount?: number;
  processingTime?: number;
  correlation?: string;
}

// Service Request/Response Types
export interface UnwrapRequest {
  zipB64?: string;
  gcsIn?: string;
  correlation?: string;
}

export interface UnwrapResponse {
  version: 'ooxml-json-1';
  kind: OOXMLKind;
  entries: PartEntry[];
  metadata?: ManifestMetadata;
  correlation?: string;
}

export interface RewrapRequest {
  manifest: Manifest;
  gcsIn?: string;
  gcsOut?: string;
  correlation?: string;
}

export interface RewrapResponse {
  version: 'ooxml-json-1';
  zipB64?: string;
  gcsOut?: string;
  metadata?: {
    finalSize?: number;
    compressionRatio?: number;
    processingTime?: number;
  };
  correlation?: string;
}

// Server-Side Operations
export type ProcessOperation =
  | ReplaceTextOperation
  | UpsertPartOperation
  | RemovePartOperation
  | RenamePartOperation;

export interface ReplaceTextOperation {
  type: 'replaceText';
  scope?: string;
  find: string;
  replace: string;
  regex?: boolean;
  flags?: string;
}

export interface UpsertPartOperation {
  type: 'upsertPart';
  path: string;
  text?: string;
  dataB64?: string;
  contentType?: string;
}

export interface RemovePartOperation {
  type: 'removePart';
  path: string;
}

export interface RenamePartOperation {
  type: 'renamePart';
  from: string;
  to: string;
  contentType?: string;
}

export interface ProcessRequest {
  zipB64?: string;
  gcsIn?: string;
  gcsOut?: string;
  ops: ProcessOperation[];
  correlation?: string;
}

export interface ProcessResponse {
  version: 'ooxml-json-1';
  zipB64?: string;
  gcsOut?: string;
  report: ProcessReport;
  correlation?: string;
}

export interface ProcessReport {
  replaced: number;
  upserted: number;
  removed: number;
  renamed: number;
  errors: ProcessError[];
  totalOps: number;
  processingTime: number;
}

export interface ProcessError {
  op: ProcessOperation;
  message: string;
  code?: string;
}

// Session Management
export interface SessionRequest {
  correlation?: string;
}

export interface SessionResponse {
  sessionId: string;
  gcsIn: string;
  gcsOut: string;
  uploadUrl: string;
  downloadUrl: string;
  expiresAt: string;
  correlation?: string;
}

// Health Check
export interface HealthResponse {
  ok: boolean;
  version: string;
  timestamp: string;
  uptime?: number;
  memory?: {
    used: number;
    total: number;
  };
}

// Error Types
export interface OOXMLErrorResponse {
  error: true;
  code: string;
  message: string;
  context?: Record<string, any>;
  correlation?: string;
  timestamp: string;
}

export class OOXMLError extends Error {
  code: string;
  context: Record<string, any>;
  correlationId: string | null;
  timestamp: string;

  constructor(code: string, message: string, context?: Record<string, any>, correlationId?: string | null);
  toLogFormat(): string;
  toJSON(): OOXMLErrorResponse;
}

// Brandbook Rules DSL Types
export interface BrandbookRulesConfig {
  profile?: string;
  version?: string;
  metadata?: RulesMetadata;
  rules: BrandRule[];
}

export interface RulesMetadata {
  name?: string;
  description?: string;
  author?: string;
  created?: string;
  updated?: string;
}

export interface BrandRule {
  id: string;
  desc?: string;
  category?: 'color' | 'font' | 'layout' | 'content' | 'accessibility';
  where: string;
  xpath: string;
  expect: RuleExpectation;
  autofix?: boolean;
  weight?: number;
  enabled?: boolean;
  tags?: string[];
}

export type RuleExpectation = 
  | { hex: string }
  | { equals: string }
  | { regex: string; flags?: string }
  | { range: { min: number; max: number } }
  | { oneOf: string[] }
  | { font: string };

export interface BrandViolation {
  ruleId: string;
  where: string;
  message: string;
  xpath?: string;
  expected?: any;
  actual?: any;
  before?: string;
  after?: string;
  autoFixed: boolean;
  weight: number;
  category?: string;
  timestamp: string;
}

export interface ValidationResult {
  version: 'ooxml-json-1';
  profile: string;
  score: number;
  totalRules: number;
  violations: BrandViolation[];
  autoFixed: number;
  duration: number;
  timestamp: string;
}

// Extension Framework Types
export type ExtensionType = 'THEME' | 'VALIDATION' | 'CONTENT' | 'TEMPLATE' | 'EXPORT';

export interface ExtensionMetadata {
  id: string;
  name: string;
  description?: string;
  version: string;
  type: ExtensionType;
  author?: string;
  requires?: string[];
  tags?: string[];
}

export interface ExtensionContext {
  profile?: string;
  now?: Date;
  correlationId?: string;
  manifest: Manifest;
  adapter?: OOXMLExtensionAdapter;
}

export interface ExtensionResult {
  success: boolean;
  message?: string;
  changes?: number;
  data?: any;
  violations?: BrandViolation[];
  metrics?: {
    duration: number;
    operationsPerformed: number;
  };
}

export interface OOXMLExtensionAdapter {
  getFile(path: string): string | null;
  setFile(path: string, content: string, contentType?: string | null): void;
  removeFile(path: string): boolean;
  replaceText(find: string | RegExp, replace: string, options?: any): Promise<any>;
  upsertPart(path: string, content: string, contentType?: string | null): Promise<any>;
  removePart(path: string): Promise<any>;
  renamePart(fromPath: string, toPath: string, contentType?: string | null): Promise<any>;
  executeOperations(): Promise<any>;
  getFilesByPattern(pattern: string | RegExp): PartEntry[];
  getMetrics(): any;
}

// Configuration Types
export interface OOXMLServiceConfig {
  PROJECT_ID: string;
  REGION: string;
  SERVICE: string;
  BUCKET?: string | null;
  PUBLIC: boolean;
  RUN_SA?: string | null;
  BUDGET_NAME: string;
  BUDGET_CURRENCY: string;
  BUDGET_AMOUNT_UNITS: string;
  BUDGET_THRESHOLDS: number[];
}

export interface DeploymentStatus {
  deployed: boolean;
  serviceUrl: string | null;
  deployedAt: string | null;
  project: string | null;
  region: string | null;
  health: HealthStatus | null;
}

export interface HealthStatus {
  available: boolean;
  statusCode?: number;
  response?: any;
  error?: string;
  checkedAt: string;
}

// Service Interface Definitions
export interface OOXMLJsonService {
  unwrap(fileIdOrBlob: string | Blob, options?: { useSession?: boolean }): Promise<Manifest>;
  rewrap(manifest: Manifest, options?: { filename?: string; gcsIn?: string; saveToGCS?: boolean }): Promise<string | any>;
  process(fileIdOrBlob: string | Blob, operations: ProcessOperation[], options?: { filename?: string; saveToGCS?: boolean }): Promise<{ fileId?: string; report: ProcessReport } | any>;
  createSession(): Promise<SessionResponse>;
  uploadToSession(uploadUrl: string, fileIdOrBlob: string | Blob): Promise<void>;
  healthCheck(): Promise<HealthStatus>;
  getServiceInfo(): any;
}

export interface OOXMLDeployment {
  showGcpPreflight(): void;
  initAndDeploy(options?: Partial<OOXMLServiceConfig>): Promise<string>;
  getDeploymentStatus(): DeploymentStatus;
  cleanup(options?: { removeService?: boolean; removeBucket?: boolean; clearProperties?: boolean }): any;
}

export interface BrandbookRulesEngine {
  validateRulesConfig(rulesConfig: BrandbookRulesConfig): { valid: boolean; errors: string[] };
  executeValidation(manifest: Manifest, rulesConfig: BrandbookRulesConfig, options?: any): Promise<ValidationResult>;
  getRuleTemplates(): Record<string, BrandRule>;
}

// Utility Types
export type CorrelationId = string;

export interface LogEntry {
  timestamp: string;
  level: 'info' | 'warn' | 'error' | 'debug';
  message: string;
  correlation?: CorrelationId;
  context?: Record<string, any>;
}

export interface MetricsEntry {
  timestamp: string;
  operation: string;
  duration_ms: number;
  success: boolean;
  correlation?: CorrelationId;
  context?: Record<string, any>;
}

// JSDoc Type Helpers for Google Apps Script
/**
 * @typedef {import('./contracts').Manifest} Manifest
 * @typedef {import('./contracts').ProcessOperation} ProcessOperation
 * @typedef {import('./contracts').BrandbookRulesConfig} BrandbookRulesConfig
 * @typedef {import('./contracts').ValidationResult} ValidationResult
 * @typedef {import('./contracts').ExtensionResult} ExtensionResult
 */

// API Version Constants
export const API_VERSION = 'ooxml-json-1' as const;
export const SCHEMA_VERSION = '1.0.0' as const;

// Error Code Constants (for type safety)
export const ERROR_CODES = {
  // Core ZIP/XML (C00x)
  C001_BAD_ZIP: 'C001',
  C002_MISSING_PART: 'C002',
  C003_XML_PARSE: 'C003',
  
  // Service/Network (S01x)
  S010_HTTP_4XX: 'S010',
  S011_HTTP_5XX: 'S011',
  S012_TIMEOUT: 'S012',
  
  // Application/Slides (A02x)
  A020_THEME_NOT_FOUND: 'A020',
  A021_LAYOUT_MAP_MISSING: 'A021',
  
  // Extension Framework (E03x)
  E030_EXTENSION_MISSING: 'E030',
  E031_EXTENSION_FAILED: 'E031',
  
  // Validation/Compliance (V04x)
  V040_RULE_PARSE: 'V040',
  V041_RULE_EXEC: 'V041'
} as const;

export type ErrorCode = typeof ERROR_CODES[keyof typeof ERROR_CODES];