import { contextBridge, ipcRenderer } from 'electron';

export type ImportResult = {
  rowCount: number;
  sourceFile: string;
};

export type ImportProgress = {
  stage: 'reading' | 'parsing' | 'importing' | 'finalizing' | 'done';
  message: string;
  current: number;
  total: number;
};

export type RaportMeta = {
  importedAt: string | null;
  sourceFile: string | null;
  rowCount: number;
};

export type DbInfo = {
  filePath: string;
  exists: boolean;
};

export type RaportPage = {
  page: number;
  pageSize: number;
  total: number;
  columns: Array<{ field: string; label: string }>;
  rows: Array<Record<string, string | null>>;
};

export type MrnBatchMeta = {
  scannedAt: string | null;
  groups: number;
  rows: number;
};

export type MrnBatchGroup = {
  numer_mrn: string;
  count: number;
};

export type MrnBatchRows = {
  numer_mrn: string;
  rows: Array<Record<string, string | null>>;
};

export type ValidationGroupKey = {
  odbiorca: string;
  kraj_wysylki: string;
  warunki_dostawy: string;
  waluta: string;
  kurs_waluty: string;
  transport_na_granicy_rodzaj: string;
  kod_towaru: string;
};

export type ValidationGroups = {
  range: { start: string; end: string };
  groups: Array<{ key: ValidationGroupKey; count: number }>;
};

export type ValidationItems = {
  range: { start: string; end: string };
  key: ValidationGroupKey;
  items: Array<{
    rowId: number;
    data_mrn: string | null;
    odbiorca: string | null;
    numer_mrn: string | null;
    coef: number | null;
    verifiedManual: boolean;
    checkable: boolean;
    outlier: boolean;
    outlierSide: 'low' | 'high' | null;
  }>;
};

export type ValidationDashboard = {
  range: { start: string; end: string };
  stats: { outliersHigh: number; outliersLow: number; singles: number; verifiedManual: number };
  days: Array<{ date: string; end: string; outliersHigh: number; outliersLow: number; singles: number; total: number }>;
};

export type ValidationDayFilter = 'all' | 'outliersHigh' | 'outliersLow' | 'singles';

export type ValidationDateGrouping = 'day' | 'days2' | 'days3' | 'week' | 'month' | 'months2';

export type ValidationGroupingOptions = { grouping?: ValidationDateGrouping };

export type ValidationExportFilters = {
  importer?: string;
  agent?: string[];
  dzial?: string;
};

export type ValidationDayItems = {
  date: string;
  totals: { all: number; outliersHigh: number; outliersLow: number; singles: number; verifiedManual: number };
  items: Array<{
    rowId: number;
    data_mrn: string | null;
    numer_mrn: string | null;
    odbiorca: string | null;
    key: ValidationGroupKey;
    coef: number | null;
    verifiedManual: boolean;
    checkable: boolean;
    outlier: boolean;
    outlierSide: 'low' | 'high' | null;
  }>;
};

export type ValidationOutlierError = {
  rowId: number;
  data_mrn: string | null;
  numer_mrn: string | null;
  nr_sad: string | null;
  agent_celny: string | null;
  odbiorca: string | null;
  key: ValidationGroupKey;
  coef: number;
  outlierSide: 'low' | 'high';
  limit: number;
  discrepancyPct: number | null;
};

export type ValidationOutlierErrors = {
  range: { start: string; end: string };
  items: ValidationOutlierError[];
};

export type ValidationExportResult = {
  ok: boolean;
  canceled?: boolean;
  filePath?: string;
  error?: string;
};

export type ValidationExportPreview = {
  period: string;
  grouping: string;
  range: { start: string; end: string };
  availableAgents: string[];
  meta: Array<{ key: string; value: string }>;
  sheets: Array<{
    name: string;
    sections: Array<{
      title: string;
      rows: Array<Record<string, unknown>>;
      totalRows: number;
      truncated: boolean;
    }>;
  }>;
};

export type ValidationExportPreviewResult = {
  ok: boolean;
  preview?: ValidationExportPreview;
  error?: string;
};

export type AgentDzialInfo = {
  filePath: string;
  exists: boolean;
  rowCount: number;
  modifiedAt: string | null;
  error?: string;
};

export type UpdateCheckResult = {
  supported: boolean;
  updateAvailable: boolean;
  currentVersion: string;
  latestVersion: string | null;
  downloadUrl: string | null;
  squirrelFeedUrl: string | null;
  manifestUrl: string | null;
  error: string | null;
};

export type UpdateStatus =
  | { state: 'idle' }
  | { state: 'checking'; message?: string }
  | { state: 'available'; message?: string }
  | { state: 'not-available'; message?: string }
  | { state: 'downloading'; percent?: number; transferred?: number; total?: number; bytesPerSecond?: number }
  | { state: 'downloaded'; message?: string }
  | { state: 'installing'; message?: string }
  | { state: 'error'; message: string };

export type UpdateStartResult = { ok: boolean; error?: string };

contextBridge.exposeInMainWorld('api', {
  onImportProgress: (handler: (p: ImportProgress) => void): (() => void) => {
    const listener = (_event: unknown, payload: ImportProgress) => handler(payload);
    ipcRenderer.on('raport:importProgress', listener);
    return () => ipcRenderer.removeListener('raport:importProgress', listener);
  },
  importRaport: (): Promise<ImportResult> => ipcRenderer.invoke('raport:import'),
  clearRaport: (): Promise<RaportMeta> => ipcRenderer.invoke('raport:clear'),
  getRaportMeta: (): Promise<RaportMeta> => ipcRenderer.invoke('raport:meta'),
  getRaportPage: (page: number, pageSize: number): Promise<RaportPage> =>
    ipcRenderer.invoke('raport:page', { page, pageSize }),
  getDbInfo: (): Promise<DbInfo> => ipcRenderer.invoke('raport:dbInfo'),
  showDbInFolder: (): Promise<boolean> => ipcRenderer.invoke('raport:showDbInFolder'),

  getAgentDzialInfo: (): Promise<AgentDzialInfo> => ipcRenderer.invoke('agentDzial:info'),
  clearAgentDzialMap: (): Promise<AgentDzialInfo> => ipcRenderer.invoke('agentDzial:clear'),
  showAgentDzialInFolder: (): Promise<boolean> => ipcRenderer.invoke('agentDzial:showInFolder'),

  rebuildMrnBatch: (): Promise<{ rowsInserted: number; groups: number; scannedAt: string | null }> =>
    ipcRenderer.invoke('mrnBatch:rebuild'),
  getMrnBatchMeta: (): Promise<MrnBatchMeta> => ipcRenderer.invoke('mrnBatch:meta'),
  getMrnBatchGroups: (limit?: number): Promise<MrnBatchGroup[]> => ipcRenderer.invoke('mrnBatch:groups', { limit }),
  getMrnBatchRows: (numerMrn: string): Promise<MrnBatchRows> => ipcRenderer.invoke('mrnBatch:rows', { numerMrn }),

  getValidationDefaultMonth: (): Promise<{ month: string | null }> => ipcRenderer.invoke('validation:defaultMonth'),
  getValidationGroups: (month: string, mrn?: string, options?: ValidationGroupingOptions): Promise<ValidationGroups> =>
    ipcRenderer.invoke('validation:groups', { month, mrn, grouping: options?.grouping }),
  getValidationItems: (month: string, key: ValidationGroupKey, mrn?: string, options?: ValidationGroupingOptions): Promise<ValidationItems> =>
    ipcRenderer.invoke('validation:items', { month, key, mrn, grouping: options?.grouping }),
  getValidationDashboard: (month: string, mrn?: string, options?: ValidationGroupingOptions): Promise<ValidationDashboard> =>
    ipcRenderer.invoke('validation:dashboard', { month, mrn, grouping: options?.grouping }),
  getValidationDayItems: (month: string, date: string, filter: ValidationDayFilter, mrn?: string, options?: ValidationGroupingOptions): Promise<ValidationDayItems> =>
    ipcRenderer.invoke('validation:dayItems', { month, date, filter, mrn, grouping: options?.grouping }),
  getValidationOutlierErrors: (month: string, mrn?: string, options?: ValidationGroupingOptions): Promise<ValidationOutlierErrors> =>
    ipcRenderer.invoke('validation:outlierErrors', { month, mrn, grouping: options?.grouping }),
  setValidationManualVerified: (rowId: number, verified: boolean): Promise<{ ok: true }> =>
    ipcRenderer.invoke('validation:setManualVerified', { rowId, verified }),
  exportValidationXlsx: (
    period: string,
    mrn?: string,
    options?: ValidationGroupingOptions,
    filters?: ValidationExportFilters,
  ): Promise<ValidationExportResult> => ipcRenderer.invoke('validation:exportXlsx', { period, mrn, grouping: options?.grouping, filters }),
  previewValidationExport: (
    period: string,
    mrn?: string,
    options?: ValidationGroupingOptions,
    filters?: ValidationExportFilters,
  ): Promise<ValidationExportPreviewResult> =>
    ipcRenderer.invoke('validation:exportPreview', { period, mrn, grouping: options?.grouping, filters }),

  getAppVersion: (): Promise<{ version: string }> => ipcRenderer.invoke('app:version'),
  checkForUpdates: (): Promise<UpdateCheckResult> => ipcRenderer.invoke('updates:check'),
  onUpdateStatus: (handler: (s: UpdateStatus) => void): (() => void) => {
    const listener = (_event: unknown, payload: UpdateStatus) => handler(payload);
    ipcRenderer.on('updates:status', listener);
    return () => ipcRenderer.removeListener('updates:status', listener);
  },
  downloadAndInstallUpdate: (feedUrl: string): Promise<UpdateStartResult> =>
    ipcRenderer.invoke('updates:downloadAndInstall', { feedUrl }),
  openExternal: (url: string): Promise<boolean> => ipcRenderer.invoke('updates:open', { url }),
  quitApp: (): Promise<boolean> => ipcRenderer.invoke('app:quit'),
});
