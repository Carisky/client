import type {
  DbInfo,
  ImportProgress,
  ImportResult,
  MrnBatchGroup,
  MrnBatchMeta,
  MrnBatchRows,
  RaportMeta,
  RaportPage,
} from './preload';

declare global {
  interface Window {
    api: {
      onImportProgress: (handler: (p: ImportProgress) => void) => () => void;
      importRaport: () => Promise<ImportResult>;
      clearRaport: () => Promise<RaportMeta>;
      getRaportMeta: () => Promise<RaportMeta>;
      getRaportPage: (page: number, pageSize: number) => Promise<RaportPage>;
      getDbInfo: () => Promise<DbInfo>;
      showDbInFolder: () => Promise<boolean>;

      rebuildMrnBatch: () => Promise<{ rowsInserted: number; groups: number; scannedAt: string | null }>;
      getMrnBatchMeta: () => Promise<MrnBatchMeta>;
      getMrnBatchGroups: (limit?: number) => Promise<MrnBatchGroup[]>;
      getMrnBatchRows: (numerMrn: string) => Promise<MrnBatchRows>;
    };
  }
}

export {};
