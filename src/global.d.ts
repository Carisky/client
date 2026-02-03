import type {
  DbInfo,
  ImportProgress,
  ImportResult,
  MrnBatchGroup,
  MrnBatchMeta,
  MrnBatchRows,
  RaportMeta,
  RaportPage,
  UpdateCheckResult,
  UpdateStartResult,
  UpdateStatus,
  ValidationDashboard,
  ValidationDayFilter,
  ValidationDayItems,
  ValidationGroupingOptions,
  ValidationGroupKey,
  ValidationGroups,
  ValidationItems,
  ValidationOutlierErrors,
  ValidationExportResult,
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
      getValidationGroups: (month: string, mrn?: string, options?: ValidationGroupingOptions) => Promise<ValidationGroups>;
      getValidationItems: (month: string, key: ValidationGroupKey, mrn?: string, options?: ValidationGroupingOptions) => Promise<ValidationItems>;
      getValidationDashboard: (month: string, mrn?: string, options?: ValidationGroupingOptions) => Promise<ValidationDashboard>;
      getValidationDayItems: (month: string, date: string, filter: ValidationDayFilter, mrn?: string, options?: ValidationGroupingOptions) => Promise<ValidationDayItems>;
      getValidationOutlierErrors: (month: string, mrn?: string, options?: ValidationGroupingOptions) => Promise<ValidationOutlierErrors>;
      setValidationManualVerified: (rowId: number, verified: boolean) => Promise<{ ok: true }>;
      exportValidationXlsx: (period: string, mrn?: string, options?: ValidationGroupingOptions) => Promise<ValidationExportResult>;

      getAppVersion: () => Promise<{ version: string }>;
      checkForUpdates: () => Promise<UpdateCheckResult>;
      onUpdateStatus: (handler: (s: UpdateStatus) => void) => () => void;
      downloadAndInstallUpdate: (feedUrl: string) => Promise<UpdateStartResult>;
      openExternal: (url: string) => Promise<boolean>;
      quitApp: () => Promise<boolean>;
    };
  }
}

export {};
