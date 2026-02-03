import { app, dialog, ipcMain, shell } from 'electron';
import {
  clearRaportData,
  getDbInfo,
  getMrnBatchGroups,
  getMrnBatchMeta,
  getMrnBatchRows,
  getRaportMeta,
  getRaportPage,
  getValidationDashboard,
  getValidationDayItems,
  getValidationDefaultMonth,
  getValidationGroups,
  getValidationItems,
  getValidationOutlierErrors,
  importRaportFromXlsx,
  rebuildMrnBatch,
  setValidationManualVerified,
} from './raportDb';
import { checkForUpdates } from './appUpdate';
import { startSquirrelUpdate } from './squirrelAutoUpdate';

export function registerRaportIpc(): void {
  ipcMain.handle('raport:import', async (event) => {
    const result = await dialog.showOpenDialog({
      title: 'Wybierz plik Excel',
      properties: ['openFile'],
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    });

    if (result.canceled || result.filePaths.length === 0) {
      return { rowCount: 0, sourceFile: '' };
    }

    const filePath = result.filePaths[0];
    return importRaportFromXlsx(filePath, (p) => event.sender.send('raport:importProgress', p));
  });

  ipcMain.handle('raport:meta', async () => getRaportMeta());

  ipcMain.handle('raport:page', async (_evt, args: { page: number; pageSize: number }) =>
    getRaportPage({ page: args.page, pageSize: args.pageSize }),
  );

  ipcMain.handle('raport:clear', async () => {
    await clearRaportData();
    return getRaportMeta();
  });

  ipcMain.handle('raport:dbInfo', async () => getDbInfo());

  ipcMain.handle('raport:showDbInFolder', async () => {
    const info = await getDbInfo();
    shell.showItemInFolder(info.filePath);
    return true;
  });

  ipcMain.handle('mrnBatch:rebuild', async () => rebuildMrnBatch());
  ipcMain.handle('mrnBatch:meta', async () => getMrnBatchMeta());
  ipcMain.handle('mrnBatch:groups', async (_evt, args?: { limit?: number }) => getMrnBatchGroups({ limit: args?.limit }));
  ipcMain.handle('mrnBatch:rows', async (_evt, args: { numerMrn: string }) => getMrnBatchRows(args?.numerMrn));

  ipcMain.handle('validation:defaultMonth', async () => getValidationDefaultMonth());
  ipcMain.handle('validation:groups', async (_evt, args: { month: string; mrn?: string }) =>
    getValidationGroups({ month: args?.month, mrn: args?.mrn }),
  );
  ipcMain.handle('validation:items', async (_evt, args: { month: string; key: unknown; mrn?: string; grouping?: unknown; offsetDays?: unknown }) =>
    getValidationItems({ month: args?.month, key: args?.key as never, mrn: args?.mrn, grouping: args?.grouping as never, offsetDays: args?.offsetDays as never }),
  );

  ipcMain.handle('validation:dashboard', async (_evt, args: { month: string; mrn?: string; grouping?: unknown; offsetDays?: unknown }) =>
    getValidationDashboard({ month: args?.month, mrn: args?.mrn, grouping: args?.grouping as never, offsetDays: args?.offsetDays as never }),
  );
  ipcMain.handle('validation:dayItems', async (_evt, args: { month: string; date: string; filter: unknown; mrn?: string; grouping?: unknown; offsetDays?: unknown }) =>
    getValidationDayItems({
      month: args?.month,
      date: args?.date,
      filter: (args?.filter as never) ?? 'all',
      mrn: args?.mrn,
      grouping: args?.grouping as never,
      offsetDays: args?.offsetDays as never,
    }),
  );
  ipcMain.handle('validation:outlierErrors', async (_evt, args: { month: string; mrn?: string; grouping?: unknown; offsetDays?: unknown }) =>
    getValidationOutlierErrors({ month: args?.month, mrn: args?.mrn, grouping: args?.grouping as never, offsetDays: args?.offsetDays as never }),
  );
  ipcMain.handle('validation:setManualVerified', async (_evt, args: { rowId: number; verified: boolean }) =>
    setValidationManualVerified({ rowId: args?.rowId, verified: Boolean(args?.verified) }),
  );

  ipcMain.handle('app:version', async () => ({ version: app.getVersion() }));
  ipcMain.handle('updates:check', async () => checkForUpdates());
  ipcMain.handle('updates:downloadAndInstall', async (event, args: { feedUrl: string }) => {
    const feedUrl = String(args?.feedUrl ?? '').trim();
    return startSquirrelUpdate(feedUrl, event.sender);
  });
  ipcMain.handle('updates:open', async (_evt, args: { url: string }) => {
    const url = String(args?.url ?? '').trim();
    if (!url) return false;
    await shell.openExternal(url);
    return true;
  });
  ipcMain.handle('app:quit', async () => {
    app.quit();
    return true;
  });
}
