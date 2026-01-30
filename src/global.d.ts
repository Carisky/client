import type { DbInfo, ImportProgress, ImportResult, RaportMeta, RaportPage } from './preload';

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
    };
  }
}

export {};
