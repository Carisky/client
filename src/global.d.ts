import type {
  AgentDzialInfo,
  DbInfo,
  ImportProgress,
  ImportResult,
  MrnBatchGroup,
  MrnBatchMeta,
  MrnBatchRows,
  RaportMeta,
  RaportPage,
  ResourcesSyncInfo,
  UpdateCheckResult,
  UpdateStartResult,
  UpdateStatus,
  ValidationDashboard,
  ValidationDayFilter,
  ValidationDayItems,
  ValidationExportOptions,
  ValidationExportFilters,
  ValidationGroupingOptions,
  ValidationGroupKey,
  ValidationGroups,
  ValidationItems,
  ValidationOutlierErrors,
  ValidationExportResult,
  ValidationExportPreviewResult,
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

      getAgentDzialInfo: () => Promise<AgentDzialInfo>;
      clearAgentDzialMap: () => Promise<AgentDzialInfo>;
      showAgentDzialInFolder: () => Promise<boolean>;

      getResourcesSyncInfo: () => Promise<ResourcesSyncInfo>;

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
      exportValidationXlsx: (
        period: string,
        mrn?: string,
        options?: ValidationGroupingOptions,
        filters?: ValidationExportFilters,
        exportOptions?: ValidationExportOptions,
      ) => Promise<ValidationExportResult>;
      previewValidationExport: (
        period: string,
        mrn?: string,
        options?: ValidationGroupingOptions,
        filters?: ValidationExportFilters,
        exportOptions?: ValidationExportOptions,
      ) => Promise<ValidationExportPreviewResult>;

      writeClipboardText: (text: string) => Promise<{ ok: boolean; error?: string }>;

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
