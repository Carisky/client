import { app } from 'electron';
import * as fs from 'fs';
import * as path from 'path';

export type UpdateCheckResult = {
  supported: boolean;
  updateAvailable: boolean;
  currentVersion: string;
  latestVersion: string | null;
  downloadUrl: string | null;
  manifestUrl: string | null;
  error: string | null;
};

function parseVersion(v: string): number[] {
  const s = String(v ?? '').trim();
  const core = s.split('-')[0] || '';
  return core
    .split('.')
    .slice(0, 3)
    .map((x) => Number.parseInt(x, 10))
    .map((n) => (Number.isFinite(n) && n >= 0 ? n : 0));
}

function compareVersions(a: string, b: string): number {
  const aa = parseVersion(a);
  const bb = parseVersion(b);
  for (let i = 0; i < 3; i++) {
    const d = (aa[i] ?? 0) - (bb[i] ?? 0);
    if (d !== 0) return d;
  }
  return 0;
}

function tryReadJsonFile(filePath: string): unknown | null {
  try {
    const s = fs.readFileSync(filePath, 'utf8');
    return JSON.parse(s) as unknown;
  } catch {
    return null;
  }
}

function resolveUpdateConfigPath(): string | null {
  const candidates: string[] = [];

  // Dev: run from client/
  candidates.push(path.resolve(process.cwd(), 'resources', 'update.json'));

  // Dev: bundled path
  candidates.push(path.resolve(__dirname, '..', '..', 'resources', 'update.json'));

  // Packaged: extraResource copied under resources/resources/
  candidates.push(path.resolve(process.resourcesPath, 'resources', 'update.json'));

  // Packaged (fallback): extraResource copied as a file (or custom layout)
  candidates.push(path.resolve(process.resourcesPath, 'update.json'));

  for (const p of candidates) {
    try {
      if (fs.existsSync(p)) return p;
    } catch {
      // ignore
    }
  }
  return null;
}

function normalizeManifestUrl(url: string): string {
  const u = String(url ?? '').trim();
  if (!u) return u;
  // Historical bug: script used ".../<branch>/client/releases/latest.json" while repo uses "/releases/latest.json".
  return u.replace('/client/releases/', '/releases/');
}

function resolveManifestUrl(): string | null {
  const fromEnv = (process.env.UPDATE_MANIFEST_URL ?? '').trim();
  if (fromEnv) return normalizeManifestUrl(fromEnv);

  const cfgPath = resolveUpdateConfigPath();
  if (!cfgPath) return null;
  const cfg = tryReadJsonFile(cfgPath);
  if (!cfg || typeof cfg !== 'object') return null;
  const url = (cfg as { manifestUrl?: unknown }).manifestUrl;
  return typeof url === 'string' && url.trim() ? normalizeManifestUrl(url) : null;
}

async function fetchJson(url: string, timeoutMs: number): Promise<unknown> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: 'GET',
      headers: { 'accept': 'application/json' },
      signal: controller.signal,
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return (await res.json()) as unknown;
  } finally {
    clearTimeout(timer);
  }
}

function extractLatestInfo(payload: unknown): { version: string | null; downloadUrl: string | null } {
  if (!payload || typeof payload !== 'object') return { version: null, downloadUrl: null };
  const p = payload as { version?: unknown; downloadUrl?: unknown; url?: unknown; win32?: unknown; platforms?: unknown };

  const version = typeof p.version === 'string' && p.version.trim() ? p.version.trim() : null;

  const dl =
    (typeof p.downloadUrl === 'string' && p.downloadUrl.trim() ? p.downloadUrl.trim() : null) ??
    (typeof p.url === 'string' && p.url.trim() ? p.url.trim() : null);

  return { version, downloadUrl: dl };
}

export async function checkForUpdates(): Promise<UpdateCheckResult> {
  const currentVersion = app.getVersion();
  const manifestUrl = resolveManifestUrl();

  if (!manifestUrl) {
    return {
      supported: false,
      updateAvailable: false,
      currentVersion,
      latestVersion: null,
      downloadUrl: null,
      manifestUrl: null,
      error: 'Update manifest URL not configured.',
    };
  }

  try {
    const u = new URL(manifestUrl);
    // Avoid CDN caching issues (GitHub raw is heavily cached).
    u.searchParams.set('_ts', String(Date.now()));
    const payload = await fetchJson(u.toString(), 4500);
    const { version: latestVersion, downloadUrl } = extractLatestInfo(payload);
    if (!latestVersion || !downloadUrl) {
      return {
        supported: true,
        updateAvailable: false,
        currentVersion,
        latestVersion: latestVersion ?? null,
        downloadUrl: downloadUrl ?? null,
        manifestUrl,
        error: 'Invalid update manifest.',
      };
    }

    const isNewer = compareVersions(latestVersion, currentVersion) > 0;
    return {
      supported: true,
      updateAvailable: isNewer,
      currentVersion,
      latestVersion,
      downloadUrl,
      manifestUrl,
      error: null,
    };
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    return {
      supported: true,
      updateAvailable: false,
      currentVersion,
      latestVersion: null,
      downloadUrl: null,
      manifestUrl,
      error: msg,
    };
  }
}
