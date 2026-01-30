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

  rebuildMrnBatch: (): Promise<{ rowsInserted: number; groups: number; scannedAt: string | null }> =>
    ipcRenderer.invoke('mrnBatch:rebuild'),
  getMrnBatchMeta: (): Promise<MrnBatchMeta> => ipcRenderer.invoke('mrnBatch:meta'),
  getMrnBatchGroups: (limit?: number): Promise<MrnBatchGroup[]> => ipcRenderer.invoke('mrnBatch:groups', { limit }),
  getMrnBatchRows: (numerMrn: string): Promise<MrnBatchRows> => ipcRenderer.invoke('mrnBatch:rows', { numerMrn }),
});
