import { app } from 'electron';
import * as fs from 'fs';
import * as path from 'path';

type ResourcesManifest = {
  formatVersion?: unknown;
  version?: unknown;
  channel?: unknown;
  generatedAt?: unknown;
  files?: unknown;
};

type ResourcesManifestFile = {
  path: string;
  size?: number;
  sha256?: string;
};

function toPosixPath(p: string): string {
  return p.replace(/\\/g, '/');
}

function isSafeRelPath(rel: string): boolean {
  if (!rel) return false;
  if (rel.includes('\0')) return false;
  const s = toPosixPath(rel).trim();
  if (!s) return false;
  if (s.startsWith('/') || s.startsWith('\\')) return false;
  const parts = s.split('/').filter(Boolean);
  if (parts.some((p) => p === '.' || p === '..')) return false;
  return true;
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

  // User cache (downloaded resources)
  candidates.push(path.resolve(app.getPath('userData'), 'resources', 'update.json'));

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
  return u.replace('/client/releases/', '/releases/');
}

function resolveResourcesManifestUrl(): string | null {
  const fromEnv = (process.env.RESOURCES_MANIFEST_URL ?? '').trim();
  if (fromEnv) return fromEnv;

  const cfgPath = resolveUpdateConfigPath();
  if (!cfgPath) return null;
  const cfg = tryReadJsonFile(cfgPath);
  if (!cfg || typeof cfg !== 'object') return null;

  const rcfg = cfg as { resourcesManifestUrl?: unknown; manifestUrl?: unknown };
  const explicit = rcfg.resourcesManifestUrl;
  if (typeof explicit === 'string' && explicit.trim()) return explicit.trim();

  const manifestUrl = rcfg.manifestUrl;
  if (typeof manifestUrl !== 'string' || !manifestUrl.trim()) return null;

  const normalized = normalizeManifestUrl(manifestUrl);
  try {
    const u = new URL(normalized);
    // Remove query; we will add cache-busting ourselves.
    u.search = '';
    if (u.pathname.endsWith('/releases/latest.json')) {
      u.pathname = u.pathname.replace(/\/releases\/latest\.json$/, '/resources/resources-manifest.json');
      return u.toString();
    }
    return null;
  } catch {
    return null;
  }
}

async function fetchJson(url: string, timeoutMs: number): Promise<unknown> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, {
      method: 'GET',
      headers: { accept: 'application/json' },
      signal: controller.signal,
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return (await res.json()) as unknown;
  } finally {
    clearTimeout(timer);
  }
}

async function fetchBinary(url: string, timeoutMs: number): Promise<Uint8Array> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url, { method: 'GET', signal: controller.signal });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buf = await res.arrayBuffer();
    return new Uint8Array(buf);
  } finally {
    clearTimeout(timer);
  }
}

function extractFiles(payload: unknown): ResourcesManifestFile[] {
  if (!payload || typeof payload !== 'object') return [];
  const p = payload as ResourcesManifest;
  if (!Array.isArray(p.files)) return [];

  const out: ResourcesManifestFile[] = [];
  for (const item of p.files) {
    if (!item || typeof item !== 'object') continue;
    const rec = item as Record<string, unknown>;
    const filePath = typeof rec.path === 'string' ? rec.path.trim() : '';
    if (!isSafeRelPath(filePath)) continue;

    const sizeRaw = rec.size;
    const size =
      typeof sizeRaw === 'number' && Number.isFinite(sizeRaw) && sizeRaw >= 0
        ? Math.floor(sizeRaw)
        : undefined;

    const sha = typeof rec.sha256 === 'string' && rec.sha256.trim() ? rec.sha256.trim() : undefined;
    out.push({ path: toPosixPath(filePath), size, sha256: sha });
  }
  return out;
}

function getUserResourcesRoot(): string {
  return path.resolve(app.getPath('userData'), 'resources');
}

function resolveBundledResourcePath(relPath: string): string | null {
  if (!isSafeRelPath(relPath)) return null;
  const relFs = relPath.split('/').join(path.sep);

  const candidates: string[] = [];
  // Dev: repo-root/client
  candidates.push(path.resolve(process.cwd(), 'resources', relFs));
  // Dev/webpack: resolve from compiled main folder
  candidates.push(path.resolve(__dirname, '..', '..', 'resources', relFs));
  // Packaged: extraResource copied under resources/resources/
  candidates.push(path.resolve(process.resourcesPath, 'resources', relFs));
  // Packaged fallback
  candidates.push(path.resolve(process.resourcesPath, relFs));

  for (const p of candidates) {
    try {
      if (fs.existsSync(p)) return p;
    } catch {
      // ignore
    }
  }
  return null;
}

function resolveUserCachedResourcePath(relPath: string): string | null {
  if (!isSafeRelPath(relPath)) return null;
  const relFs = relPath.split('/').join(path.sep);
  const p = path.join(getUserResourcesRoot(), relFs);
  try {
    if (fs.existsSync(p)) return p;
  } catch {
    // ignore
  }
  return null;
}

export function resolveResourcePath(relPath: string): string | null {
  return resolveUserCachedResourcePath(relPath) ?? resolveBundledResourcePath(relPath);
}

function shouldDownloadToCache(params: { relPath: string; size?: number }): boolean {
  const existing = resolveResourcePath(params.relPath);
  if (!existing) return true;
  if (params.size == null) return false;
  try {
    const st = fs.statSync(existing);
    return st.size !== params.size;
  } catch {
    return true;
  }
}

function safeJoin(root: string, relPosix: string): string | null {
  if (!isSafeRelPath(relPosix)) return null;
  const relFs = relPosix.split('/').join(path.sep);
  const full = path.resolve(root, relFs);
  const rel = path.relative(root, full);
  if (!rel || rel === '.') return full;
  if (rel.startsWith('..') || path.isAbsolute(rel)) return null;
  return full;
}

export async function autoFixResourcesOnStartup(): Promise<void> {
  if (!app.isPackaged) return;

  const resourcesManifestUrl = resolveResourcesManifestUrl();
  if (!resourcesManifestUrl) return;

  let baseUrl: URL | null = null;
  try {
    baseUrl = new URL(resourcesManifestUrl);
  } catch {
    return;
  }

  const userRoot = getUserResourcesRoot();
  try {
    fs.mkdirSync(userRoot, { recursive: true });
  } catch {
    // ignore
  }

  const u = new URL(resourcesManifestUrl);
  u.searchParams.set('_ts', String(Date.now()));

  let payload: unknown;
  try {
    payload = await fetchJson(u.toString(), 4500);
  } catch {
    return;
  }

  const files = extractFiles(payload);
  if (!files.length) return;

  // Compute base URL from manifest URL directory.
  const resourcesRootUrl = new URL(baseUrl.toString());
  resourcesRootUrl.search = '';
  resourcesRootUrl.pathname = resourcesRootUrl.pathname.replace(/\\/g, '/').replace(/[^/]+$/, '');

  for (const f of files) {
    if (!shouldDownloadToCache({ relPath: f.path, size: f.size })) continue;

    const outPath = safeJoin(userRoot, f.path);
    if (!outPath) continue;

    try {
      fs.mkdirSync(path.dirname(outPath), { recursive: true });
    } catch {
      continue;
    }

    const fileUrl = new URL(resourcesRootUrl.toString());
    fileUrl.pathname =
      resourcesRootUrl.pathname.replace(/\/$/, '/') +
      f.path
        .split('/')
        .map((seg) => encodeURIComponent(seg))
        .join('/');
    fileUrl.searchParams.set('_ts', String(Date.now()));

    try {
      const bin = await fetchBinary(fileUrl.toString(), 8000);
      if (f.size != null && bin.byteLength !== f.size) {
        continue;
      }

      const tmp = `${outPath}.tmp`;
      fs.writeFileSync(tmp, bin);
      fs.renameSync(tmp, outPath);
    } catch {
      // ignore
    }
  }
}
