import { app, autoUpdater } from 'electron';
import type { WebContents } from 'electron';

export type UpdateStatus =
  | { state: 'idle' }
  | { state: 'checking'; message?: string }
  | { state: 'available'; message?: string }
  | { state: 'not-available'; message?: string }
  | { state: 'downloading'; percent?: number; transferred?: number; total?: number; bytesPerSecond?: number }
  | { state: 'downloaded'; message?: string }
  | { state: 'installing'; message?: string }
  | { state: 'error'; message: string };

let listenersAttached = false;
let statusTarget: WebContents | null = null;

function send(status: UpdateStatus) {
  try {
    statusTarget?.send('updates:status', status);
  } catch {
    // ignore
  }
}

function attachListenersOnce() {
  if (listenersAttached) return;
  listenersAttached = true;

  autoUpdater.on('checking-for-update', () => send({ state: 'checking' }));
  autoUpdater.on('update-available', () => send({ state: 'available' }));
  autoUpdater.on('update-not-available', () => send({ state: 'not-available' }));
  autoUpdater.on('download-progress', (p) => {
    const percent = typeof p?.percent === 'number' ? p.percent : undefined;
    const transferred = typeof p?.transferred === 'number' ? p.transferred : undefined;
    const total = typeof p?.total === 'number' ? p.total : undefined;
    const bytesPerSecond = typeof p?.bytesPerSecond === 'number' ? p.bytesPerSecond : undefined;
    send({ state: 'downloading', percent, transferred, total, bytesPerSecond });
  });
  autoUpdater.on('update-downloaded', () => {
    send({ state: 'downloaded' });
    setTimeout(() => {
      send({ state: 'installing' });
      try {
        autoUpdater.quitAndInstall();
      } catch (e: unknown) {
        const msg = e instanceof Error ? e.message : String(e);
        send({ state: 'error', message: msg });
      }
    }, 800);
  });
  autoUpdater.on('error', (e) => {
    const msg = e instanceof Error ? e.message : String(e);
    send({ state: 'error', message: msg });
  });
}

export function canAutoUpdate(): { ok: boolean; error?: string } {
  if (process.platform !== 'win32') return { ok: false, error: 'Auto-update supported only on Windows.' };
  if (!app.isPackaged) return { ok: false, error: 'Auto-update works only in packaged app.' };
  return { ok: true };
}

export function startSquirrelUpdate(feedUrl: string, target: WebContents): { ok: boolean; error?: string } {
  const cap = canAutoUpdate();
  if (!cap.ok) return cap;

  const url = String(feedUrl ?? '').trim().replace(/\/+$/, '');
  if (!url) return { ok: false, error: 'Missing feedUrl.' };

  statusTarget = target;
  attachListenersOnce();

  try {
    autoUpdater.setFeedURL({ url });
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    return { ok: false, error: msg };
  }

  try {
    autoUpdater.checkForUpdates();
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    return { ok: false, error: msg };
  }

  return { ok: true };
}

