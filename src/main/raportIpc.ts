import { app, clipboard, dialog, ipcMain, shell } from 'electron';
import * as path from 'path';
import * as xlsx from 'xlsx';
import {
  clearAgentDzialMap,
  clearRaportData,
  exportValidationWynikiToXlsx,
  getAgentDzialInfo,
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
  previewValidationWynikiExport,
  rebuildMrnBatch,
  setValidationManualVerified,
} from './raportDb';
import { checkForUpdates } from './appUpdate';
import { startSquirrelUpdate } from './squirrelAutoUpdate';
import { getResourcesSyncInfo } from './resourcesAutoFix';

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

  ipcMain.handle('agentDzial:info', async () => getAgentDzialInfo());
  ipcMain.handle('agentDzial:clear', async () => clearAgentDzialMap());
  ipcMain.handle('agentDzial:showInFolder', async () => {
    const info = await getAgentDzialInfo();
    shell.showItemInFolder(info.filePath);
    return true;
  });

  ipcMain.handle('resources:syncInfo', async () => getResourcesSyncInfo());

  ipcMain.handle('clipboard:writeText', async (_evt, args: { text: string }) => {
    try {
      const text = String(args?.text ?? '');
      clipboard.writeText(text);
      return { ok: true as const };
    } catch (e: unknown) {
      return { ok: false as const, error: e instanceof Error ? e.message : String(e) };
    }
  });

  ipcMain.handle('mrnBatch:rebuild', async () => rebuildMrnBatch());
  ipcMain.handle('mrnBatch:meta', async () => getMrnBatchMeta());
  ipcMain.handle('mrnBatch:groups', async (_evt, args?: { limit?: number }) => getMrnBatchGroups({ limit: args?.limit }));
  ipcMain.handle('mrnBatch:rows', async (_evt, args: { numerMrn: string }) => getMrnBatchRows(args?.numerMrn));

  ipcMain.handle('validation:defaultMonth', async () => getValidationDefaultMonth());
  ipcMain.handle('validation:groups', async (_evt, args: { month: string; mrn?: string; grouping?: unknown }) =>
    getValidationGroups({ month: args?.month, mrn: args?.mrn, grouping: args?.grouping as never }),
  );
  ipcMain.handle('validation:items', async (_evt, args: { month: string; key: unknown; mrn?: string; grouping?: unknown }) =>
    getValidationItems({ month: args?.month, key: args?.key as never, mrn: args?.mrn, grouping: args?.grouping as never }),
  );

  ipcMain.handle('validation:dashboard', async (_evt, args: { month: string; mrn?: string; grouping?: unknown }) =>
    getValidationDashboard({ month: args?.month, mrn: args?.mrn, grouping: args?.grouping as never }),
  );
  ipcMain.handle('validation:dayItems', async (_evt, args: { month: string; date: string; filter: unknown; mrn?: string; grouping?: unknown }) =>
    getValidationDayItems({
      month: args?.month,
      date: args?.date,
      filter: (args?.filter as never) ?? 'all',
      mrn: args?.mrn,
      grouping: args?.grouping as never,
    }),
  );
  ipcMain.handle('validation:outlierErrors', async (_evt, args: { month: string; mrn?: string; grouping?: unknown }) =>
    getValidationOutlierErrors({ month: args?.month, mrn: args?.mrn, grouping: args?.grouping as never }),
  );
  ipcMain.handle('validation:setManualVerified', async (_evt, args: { rowId: number; verified: boolean }) =>
    setValidationManualVerified({ rowId: args?.rowId, verified: Boolean(args?.verified) }),
  );
  ipcMain.handle(
    'validation:exportXlsx',
    async (
      _evt,
      args: {
        period: string;
        mrn?: string;
        grouping?: unknown;
        filters?: { importer?: unknown; agent?: unknown; dzial?: unknown };
        exportOptions?: unknown;
      },
    ) => {
    const period = String(args?.period ?? '').trim();
    const grouping = String(args?.grouping ?? 'day').trim();
    const mrn = String(args?.mrn ?? '').trim();
    const f = args?.filters ?? {};
    const importer = String(f?.importer ?? '').trim();
    const dzial = String(f?.dzial ?? '').trim();
    const agentRaw = f?.agent;
    const agents = (Array.isArray(agentRaw) ? agentRaw : typeof agentRaw === 'string' ? [agentRaw] : [])
      .map((v) => String(v ?? '').trim())
      .filter(Boolean);

    const safePart = (v: string) =>
      v
        .replace(/[\\/:*?"<>|]+/g, '_')
        .replace(/\s+/g, ' ')
        .trim()
        .slice(0, 80);

    const parts = [`Wyniki_${safePart(period || 'okres')}_${safePart(grouping || 'day')}`];
    if (mrn) parts.push(`MRN_${safePart(mrn)}`);
    if (importer) parts.push(`Importer_${safePart(importer)}`);
    if (agents.length === 1) parts.push(`Agent_${safePart(agents[0] ?? '')}`);
    else if (agents.length > 1) parts.push(`Agents_${agents.length}`);
    if (dzial) parts.push(`Dzial_${safePart(dzial)}`);
    const fileName = `${parts.join('_')}.xlsx`.slice(0, 200);
    const defaultPath = path.join(app.getPath('downloads'), fileName);

    const res = await dialog.showSaveDialog({
      title: 'Eksportuj Wyniki (IQR) do Excel',
      defaultPath,
      filters: [{ name: 'Excel', extensions: ['xlsx'] }],
    });

    if (res.canceled || !res.filePath) return { ok: false, canceled: true };
    try {
      await exportValidationWynikiToXlsx({
        period,
        mrn: mrn || undefined,
        grouping: args?.grouping,
        filters: {
          importer: importer || undefined,
          agent: agents.length ? agents : undefined,
          dzial: dzial || undefined,
        },
        exportOptions: args?.exportOptions,
        filePath: res.filePath,
      });
      return { ok: true, filePath: res.filePath };
    } catch (e: unknown) {
      return { ok: false, error: e instanceof Error ? e.message : String(e) };
    }
    },
  );

  ipcMain.handle(
    'validation:exportPreview',
    async (
      _evt,
      args: {
        period: string;
        mrn?: string;
        grouping?: unknown;
        filters?: { importer?: unknown; agent?: unknown; dzial?: unknown };
        exportOptions?: unknown;
      },
    ) => {
      const period = String(args?.period ?? '').trim();
      const mrn = String(args?.mrn ?? '').trim();
      const f = args?.filters ?? {};
      const importer = String(f?.importer ?? '').trim();
      const dzial = String(f?.dzial ?? '').trim();
      const agentRaw = f?.agent;
      const agents = (Array.isArray(agentRaw) ? agentRaw : typeof agentRaw === 'string' ? [agentRaw] : [])
        .map((v) => String(v ?? '').trim())
        .filter(Boolean);

      try {
        const preview = await previewValidationWynikiExport({
          period,
          mrn: mrn || undefined,
          grouping: args?.grouping,
          filters: {
            importer: importer || undefined,
            agent: agents.length ? agents : undefined,
            dzial: dzial || undefined,
          },
          exportOptions: args?.exportOptions,
          limit: 200,
        });
        return { ok: true, preview };
      } catch (e: unknown) {
        return { ok: false, error: e instanceof Error ? e.message : String(e) };
      }
    },
  );

  ipcMain.handle(
    'attention:quickExportXlsx',
    async (
      _evt,
      args: {
        period: string;
        grouping?: unknown;
        range?: { start?: unknown; end?: unknown };
        agentFilter?: unknown;
        rows?: Array<{
          dataMrn?: unknown;
          nrSad?: unknown;
          mrn?: unknown;
          discrepancyPct?: unknown;
          side?: unknown;
          agent?: unknown;
        }>;
      },
    ) => {
      const period = String(args?.period ?? '').trim();
      if (!period) return { ok: false, error: 'missing period' };

      const grouping = String(args?.grouping ?? 'day').trim() || 'day';
      const rangeStart = String(args?.range?.start ?? '').trim();
      const rangeEnd = String(args?.range?.end ?? '').trim();
      const safePart = (v: string) =>
        v
          .replace(/[\\/:*?"<>|]+/g, '_')
          .replace(/\s+/g, ' ')
          .trim()
          .slice(0, 80);
      const toText = (v: unknown): string => String(v ?? '').trim();
      const toPctText = (v: number | null): string => {
        if (v == null || !Number.isFinite(v)) return '';
        if (v < 0.05) return '<0.1%';
        return `${v.toFixed(1)}%`;
      };

      const parsedRows = (Array.isArray(args?.rows) ? args.rows : [])
        .map((r) => {
          const discrepancyRaw = Number(r?.discrepancyPct);
          const discrepancyPct = Number.isFinite(discrepancyRaw) ? discrepancyRaw : null;
          const sideRaw = toText(r?.side).toLowerCase();
          const side = sideRaw === 'high' || sideRaw === 'low' ? sideRaw : '';
          return {
            dataMrn: toText(r?.dataMrn),
            nrSad: toText(r?.nrSad),
            mrn: toText(r?.mrn),
            discrepancyPct,
            side,
            agent: toText(r?.agent),
          };
        })
        .filter((r) => r.dataMrn || r.nrSad || r.mrn || r.agent || r.discrepancyPct != null);

      if (!parsedRows.length) return { ok: false, error: 'no rows to export' };

      const agentFilter = (Array.isArray(args?.agentFilter) ? args.agentFilter : [])
        .map((v) => toText(v))
        .filter(Boolean);

      const parts = [`Uwagi_${safePart(period)}_${safePart(grouping)}`, `Rows_${parsedRows.length}`];
      if (agentFilter.length === 1) parts.push(`Agent_${safePart(agentFilter[0] ?? '')}`);
      else if (agentFilter.length > 1) parts.push(`Agents_${agentFilter.length}`);
      const fileName = `${parts.join('_')}.xlsx`.slice(0, 200);
      const defaultPath = path.join(app.getPath('downloads'), fileName);

      const res = await dialog.showSaveDialog({
        title: 'Szybki eksport (Do Mojej Uwagi) do Excel',
        defaultPath,
        filters: [{ name: 'Excel', extensions: ['xlsx'] }],
      });
      if (res.canceled || !res.filePath) return { ok: false, canceled: true };

      try {
        const rows = parsedRows.map((r, idx) => ({
          Lp: idx + 1,
          DataMRN: r.dataMrn,
          NrSAD: r.nrSad,
          MRN: r.mrn,
          OdchylenieIQR: toPctText(r.discrepancyPct),
          Kierunek: r.side === 'high' ? '↑' : r.side === 'low' ? '↓' : '',
          AgentCelny: r.agent,
        }));

        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet(rows);
        if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
        ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activePane: 'bottomLeft', state: 'frozen' };
        ws['!cols'] = [
          { wch: 7 },
          { wch: 14 },
          { wch: 14 },
          { wch: 28 },
          { wch: 16 },
          { wch: 10 },
          { wch: 28 },
        ];
        xlsx.utils.book_append_sheet(wb, ws, 'DoMojejUwagi');

        const metaRows: Array<[string, string | number]> = [
          ['Raport', 'Do Mojej Uwagi (szybki eksport)'],
          ['ExportedAt', new Date().toISOString()],
          ['Period', period],
          ['Grouping', grouping],
          ['RangeStart', rangeStart],
          ['RangeEnd', rangeEnd],
          ['AgentFilter', agentFilter.length ? agentFilter.join(', ') : 'wszyscy'],
          ['Rows', rows.length],
        ];
        const wsMeta = xlsx.utils.aoa_to_sheet(metaRows);
        wsMeta['!cols'] = [{ wch: 20 }, { wch: 80 }];
        xlsx.utils.book_append_sheet(wb, wsMeta, 'Meta');

        xlsx.writeFile(wb, res.filePath, { bookType: 'xlsx', compression: true });
        return { ok: true, filePath: res.filePath };
      } catch (e: unknown) {
        return { ok: false, error: e instanceof Error ? e.message : String(e) };
      }
    },
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
