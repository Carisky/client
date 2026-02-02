import type {
  DbInfo,
  ImportProgress,
  ImportResult,
  MrnBatchGroup,
  MrnBatchMeta,
  MrnBatchRows,
  RaportMeta,
  RaportPage,
  ValidationDashboard,
  ValidationDayFilter,
  ValidationDayItems,
  ValidationGroupKey,
  ValidationGroups,
  ValidationItems,
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

      getValidationDefaultMonth: () => Promise<{ month: string | null }>;
      getValidationGroups: (month: string) => Promise<ValidationGroups>;
      getValidationItems: (month: string, key: ValidationGroupKey) => Promise<ValidationItems>;
      getValidationDashboard: (month: string) => Promise<ValidationDashboard>;
      getValidationDayItems: (month: string, date: string, filter: ValidationDayFilter) => Promise<ValidationDayItems>;
      setValidationManualVerified: (rowId: number, verified: boolean) => Promise<{ ok: true }>;
    };
  }
}

export {};
