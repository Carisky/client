import { app, clipboard, dialog, ipcMain, shell } from 'electron';
import * as fs from 'fs';
import * as path from 'path';
import * as xlsx from 'xlsx';
import {
  assertFeatureAccess,
  bootstrapSuperAdmin,
  completePasswordSetup,
  getAdminPanelData,
  getAuthSessionState,
  loginWithPassword,
  loginWithToken,
  logout,
  rotateAdminUserToken,
  saveAdminUser,
  savePermissionGroup,
} from './adminAuth';
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
import { getResourcesSyncInfo, resolveResourcePath } from './resourcesAutoFix';

function getResourceAsDataUrl(relPath: string, mimeType: string): string | null {
  const filePath = resolveResourcePath(relPath);
  if (!filePath) return null;
  try {
    const content = fs.readFileSync(filePath);
    return `data:${mimeType};base64,${content.toString('base64')}`;
  } catch {
    return null;
  }
}

function getRestrictedVisibleAgentNames(
  effectivePermissions: string[] | undefined,
  fullName: string | null | undefined,
  allPermissionKey: string,
): string[] | null {
  if (effectivePermissions?.includes(allPermissionKey)) return null;
  const normalizedFullName = String(fullName ?? '').trim();
  return [normalizedFullName || '__NO_VISIBLE_AGENT__'];
}

export function registerRaportIpc(): void {
  ipcMain.handle('auth:session', async () => getAuthSessionState());
  ipcMain.handle(
    'auth:bootstrap',
    async (_evt, args: { login?: unknown; fullName?: unknown }) =>
      bootstrapSuperAdmin({
        login: String(args?.login ?? ''),
        fullName: String(args?.fullName ?? ''),
      }),
  );
  ipcMain.handle(
    'auth:loginPassword',
    async (_evt, args: { login?: unknown; password?: unknown }) =>
      loginWithPassword({
        login: String(args?.login ?? ''),
        password: String(args?.password ?? ''),
      }),
  );
  ipcMain.handle(
    'auth:loginToken',
    async (_evt, args: { login?: unknown; token?: unknown }) =>
      loginWithToken({
        login: String(args?.login ?? ''),
        token: String(args?.token ?? ''),
      }),
  );
  ipcMain.handle(
    'auth:completePasswordSetup',
    async (_evt, args: { password?: unknown }) =>
      completePasswordSetup({
        password: String(args?.password ?? ''),
      }),
  );
  ipcMain.handle('auth:logout', async () => logout());

  ipcMain.handle('admin:panel', async () => getAdminPanelData());
  ipcMain.handle(
    'admin:saveUser',
    async (
      _evt,
      args: {
        userId?: unknown;
        login?: unknown;
        fullName?: unknown;
        systemRole?: unknown;
        isActive?: unknown;
        groupIds?: unknown;
      },
    ) =>
      saveAdminUser({
        userId: args?.userId ? String(args.userId) : undefined,
        login: String(args?.login ?? ''),
        fullName: String(args?.fullName ?? ''),
        systemRole: String(args?.systemRole ?? '') as 'ADMIN' | 'USER',
        isActive: Boolean(args?.isActive),
        groupIds: Array.isArray(args?.groupIds)
          ? args.groupIds.map((groupId) => String(groupId ?? ''))
          : [],
      }),
  );
  ipcMain.handle(
    'admin:rotateUserToken',
    async (_evt, args: { userId?: unknown }) =>
      rotateAdminUserToken({
        userId: String(args?.userId ?? ''),
      }),
  );
  ipcMain.handle(
    'admin:savePermissionGroup',
    async (
      _evt,
      args: {
        groupId?: unknown;
        name?: unknown;
        description?: unknown;
        permissionKeys?: unknown;
      },
    ) =>
      savePermissionGroup({
        groupId: args?.groupId ? String(args.groupId) : undefined,
        name: String(args?.name ?? ''),
        description: String(args?.description ?? ''),
        permissionKeys: Array.isArray(args?.permissionKeys)
          ? args.permissionKeys.map((permissionKey) => String(permissionKey ?? '')) as never
          : [],
      }),
  );

  ipcMain.handle('raport:import', async (event) => {
    await assertFeatureAccess('REPORT_IMPORT');
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

  ipcMain.handle('raport:meta', async () => {
    await assertFeatureAccess('AUTHENTICATED');
    return getRaportMeta();
  });

  ipcMain.handle('raport:page', async (_evt, args: { page: number; pageSize: number }) => {
    await assertFeatureAccess('REPORT_VIEW');
    return getRaportPage({ page: args.page, pageSize: args.pageSize });
  });

  ipcMain.handle('raport:clear', async () => {
    await assertFeatureAccess('REPORT_IMPORT');
    await clearRaportData();
    return getRaportMeta();
  });

  ipcMain.handle('raport:dbInfo', async () => {
    await assertFeatureAccess('SETTINGS_VIEW');
    return getDbInfo();
  });

  ipcMain.handle('raport:showDbInFolder', async () => {
    await assertFeatureAccess('SETTINGS_VIEW');
    const info = await getDbInfo();
    shell.showItemInFolder(info.filePath);
    return true;
  });

  ipcMain.handle('agentDzial:info', async () => {
    await assertFeatureAccess('SETTINGS_VIEW');
    return getAgentDzialInfo();
  });
  ipcMain.handle('agentDzial:clear', async () => {
    await assertFeatureAccess('SETTINGS_VIEW');
    return clearAgentDzialMap();
  });
  ipcMain.handle('agentDzial:showInFolder', async () => {
    await assertFeatureAccess('SETTINGS_VIEW');
    const info = await getAgentDzialInfo();
    shell.showItemInFolder(info.filePath);
    return true;
  });

  ipcMain.handle('resources:syncInfo', async () => {
    await assertFeatureAccess('SETTINGS_VIEW');
    return getResourcesSyncInfo();
  });

  ipcMain.handle('clipboard:writeText', async (_evt, args: { text: string }) => {
    try {
      const text = String(args?.text ?? '');
      clipboard.writeText(text);
      return { ok: true as const };
    } catch (e: unknown) {
      return { ok: false as const, error: e instanceof Error ? e.message : String(e) };
    }
  });

  ipcMain.handle('mrnBatch:rebuild', async () => {
    await assertFeatureAccess('REPORT_DUPLICATES_VIEW');
    return rebuildMrnBatch();
  });
  ipcMain.handle('mrnBatch:meta', async () => {
    await assertFeatureAccess('REPORT_DUPLICATES_VIEW');
    return getMrnBatchMeta();
  });
  ipcMain.handle('mrnBatch:groups', async (_evt, args?: { limit?: number }) => {
    await assertFeatureAccess('REPORT_DUPLICATES_VIEW');
    return getMrnBatchGroups({ limit: args?.limit });
  });
  ipcMain.handle('mrnBatch:rows', async (_evt, args: { numerMrn: string }) => {
    await assertFeatureAccess('REPORT_DUPLICATES_VIEW');
    return getMrnBatchRows(args?.numerMrn);
  });

  ipcMain.handle('validation:defaultMonth', async () => {
    await assertFeatureAccess('VALIDATION_VIEW');
    return getValidationDefaultMonth();
  });
  ipcMain.handle('validation:groups', async (_evt, args: { month: string; mrn?: string; grouping?: unknown }) => {
    await assertFeatureAccess('VALIDATION_VIEW');
    return getValidationGroups({ month: args?.month, mrn: args?.mrn, grouping: args?.grouping as never });
  });
  ipcMain.handle('validation:items', async (_evt, args: { month: string; key: unknown; mrn?: string; grouping?: unknown }) => {
    await assertFeatureAccess('VALIDATION_VIEW');
    return getValidationItems({ month: args?.month, key: args?.key as never, mrn: args?.mrn, grouping: args?.grouping as never });
  });

  ipcMain.handle('validation:dashboard', async (_evt, args: { month: string; mrn?: string; grouping?: unknown }) => {
    await assertFeatureAccess('VALIDATION_VIEW');
    return getValidationDashboard({ month: args?.month, mrn: args?.mrn, grouping: args?.grouping as never });
  });
  ipcMain.handle('validation:dayItems', async (_evt, args: { month: string; date: string; filter: unknown; mrn?: string; grouping?: unknown }) => {
    await assertFeatureAccess('VALIDATION_VIEW');
    return getValidationDayItems({
      month: args?.month,
      date: args?.date,
      filter: (args?.filter as never) ?? 'all',
      mrn: args?.mrn,
      grouping: args?.grouping as never,
    });
  });
  ipcMain.handle('validation:outlierErrors', async (_evt, args: { month: string; mrn?: string; grouping?: unknown }) => {
    await assertFeatureAccess('VALIDATION_VIEW');
    return getValidationOutlierErrors({ month: args?.month, mrn: args?.mrn, grouping: args?.grouping as never });
  });
  ipcMain.handle('attention:outlierErrors', async (_evt, args: { month: string; mrn?: string; grouping?: unknown }) => {
    await assertFeatureAccess('ATTENTION_VIEW');
    const session = await getAuthSessionState();
    const visibleAgentNames = getRestrictedVisibleAgentNames(
      session.user?.effectivePermissions,
      session.user?.fullName,
      'ATTENTION_VIEW_ALL',
    );
    return getValidationOutlierErrors({
      month: args?.month,
      mrn: args?.mrn,
      grouping: args?.grouping as never,
      visibleAgentNames,
    });
  });
  ipcMain.handle('validation:setManualVerified', async (_evt, args: { rowId: number; verified: boolean }) => {
    await assertFeatureAccess('VALIDATION_VERIFY_MANUAL');
    return setValidationManualVerified({ rowId: args?.rowId, verified: Boolean(args?.verified) });
  });
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
      await assertFeatureAccess('EXPORT_VIEW');
      const session = await getAuthSessionState();
      const visibleAgentNames = getRestrictedVisibleAgentNames(
        session.user?.effectivePermissions,
        session.user?.fullName,
        'EXPORT_VIEW_ALL',
      );

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
          visibleAgentNames,
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
      await assertFeatureAccess('EXPORT_VIEW');
      const session = await getAuthSessionState();
      const visibleAgentNames = getRestrictedVisibleAgentNames(
        session.user?.effectivePermissions,
        session.user?.fullName,
        'EXPORT_VIEW_ALL',
      );

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
          visibleAgentNames,
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
      await assertFeatureAccess('EXPORT_VIEW');

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
  ipcMain.handle('app:logoDataUrl', async () =>
    getResourceAsDataUrl('logo.png', 'image/png'),
  );
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
