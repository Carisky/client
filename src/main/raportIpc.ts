import { dialog, ipcMain, shell } from 'electron';
import {
  clearRaportData,
  getDbInfo,
  getMrnBatchGroups,
  getMrnBatchMeta,
  getMrnBatchRows,
  getRaportMeta,
  getRaportPage,
  getValidationDefaultMonth,
  getValidationGroups,
  getValidationItems,
  importRaportFromXlsx,
  rebuildMrnBatch,
} from './raportDb';

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
  ipcMain.handle('validation:groups', async (_evt, args: { month: string }) => getValidationGroups({ month: args?.month }));
  ipcMain.handle('validation:items', async (_evt, args: { month: string; key: unknown }) =>
    getValidationItems({ month: args?.month, key: args?.key as never }),
  );
}
