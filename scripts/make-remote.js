/* eslint-disable no-console */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const childProcess = require('child_process');

function exec(cmd, opts = {}) {
  return childProcess.execSync(cmd, { stdio: 'pipe', encoding: 'utf8', ...opts }).trim();
}

function execInherit(cmd, opts = {}) {
  childProcess.execSync(cmd, { stdio: 'inherit', ...opts });
}

function execFileInherit(file, args, opts = {}) {
  childProcess.execFileSync(file, args, { stdio: 'inherit', ...opts });
}

function parseGithubRepo(remoteUrl) {
  const s = String(remoteUrl || '').trim();
  // git@github.com:owner/repo.git
  let m = /^git@github\.com:([^/]+)\/(.+?)(?:\.git)?$/.exec(s);
  if (m) return { owner: m[1], repo: m[2] };
  // https://github.com/owner/repo.git
  m = /^https?:\/\/github\.com\/([^/]+)\/(.+?)(?:\.git)?$/.exec(s);
  if (m) return { owner: m[1], repo: m[2] };
  return null;
}

function getGithubToken() {
  const t = (process.env.GITHUB_TOKEN || process.env.GH_TOKEN || '').trim();
  return t || null;
}

async function ghApi(url, { method = 'GET', token, headers = {}, body } = {}) {
  const res = await fetch(url, {
    method,
    headers: {
      accept: 'application/vnd.github+json',
      ...(token ? { authorization: `Bearer ${token}` } : {}),
      'user-agent': 'make-remote',
      ...headers,
    },
    body,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`GitHub API ${method} ${url} -> HTTP ${res.status} ${text}`.slice(0, 1200));
  }
  const ct = res.headers.get('content-type') || '';
  if (ct.includes('application/json')) return res.json();
  return res.text();
}

async function getOrCreateRelease({ owner, repo, tag, token }) {
  // Try get by tag
  try {
    const rel = await ghApi(`https://api.github.com/repos/${owner}/${repo}/releases/tags/${encodeURIComponent(tag)}`, {
      token,
    });
    return rel;
  } catch {
    // ignore
  }

  const rel = await ghApi(`https://api.github.com/repos/${owner}/${repo}/releases`, {
    method: 'POST',
    token,
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      tag_name: tag,
      name: tag,
      draft: false,
      prerelease: false,
      generate_release_notes: false,
    }),
  });
  return rel;
}

async function uploadReleaseAsset({ uploadUrl, filePath, token }) {
  const name = path.basename(filePath);
  const url = String(uploadUrl).replace('{?name,label}', `?name=${encodeURIComponent(name)}`);
  const stat = fs.statSync(filePath);
  const stream = fs.createReadStream(filePath);

  console.log(`Uploading asset (${Math.round(stat.size / (1024 * 1024))} MB): ${name}`);
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      accept: 'application/vnd.github+json',
      authorization: `Bearer ${token}`,
      'user-agent': 'make-remote',
      'content-type': 'application/octet-stream',
      'content-length': String(stat.size),
    },
    body: stream,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Upload failed -> HTTP ${res.status} ${text}`.slice(0, 1200));
  }
  return res.json();
}

async function deleteReleaseAssetByName({ owner, repo, releaseId, name, token }) {
  const assets = await ghApi(`https://api.github.com/repos/${owner}/${repo}/releases/${releaseId}/assets`, { token });
  if (!Array.isArray(assets)) return;
  const existing = assets.find((a) => a && typeof a === 'object' && a.name === name);
  if (!existing) return;
  const assetId = existing.id;
  if (!assetId) return;
  console.log(`Deleting existing asset: ${name}`);
  await ghApi(`https://api.github.com/repos/${owner}/${repo}/releases/assets/${assetId}`, {
    method: 'DELETE',
    token,
  });
}

function walk(dir) {
  const out = [];
  const stack = [dir];
  while (stack.length) {
    const d = stack.pop();
    const entries = fs.readdirSync(d, { withFileTypes: true });
    for (const e of entries) {
      const p = path.join(d, e.name);
      if (e.isDirectory()) stack.push(p);
      else out.push(p);
    }
  }
  return out;
}

function sha256(filePath) {
  const h = crypto.createHash('sha256');
  h.update(fs.readFileSync(filePath));
  return h.digest('hex');
}

function ensureDir(p) {
  fs.mkdirSync(p, { recursive: true });
}

function writeJson(filePath, value) {
  ensureDir(path.dirname(filePath));
  fs.writeFileSync(filePath, JSON.stringify(value, null, 2) + '\n', 'utf8');
}

async function main() {
  const root = path.resolve(__dirname, '..');
  process.chdir(root);

  const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
  const version = String(pkg.version || '').trim();
  if (!version) throw new Error('Missing package.json version');

  const branch = exec('git rev-parse --abbrev-ref HEAD', { cwd: root });
  const origin = exec('git remote get-url origin', { cwd: root });
  if (!origin) {
    throw new Error('No git remote "origin". Add it first: git remote add origin <url>');
  }

  const gh = parseGithubRepo(origin);
  if (!gh) {
    throw new Error('Only GitHub origin URLs are supported for auto manifest URL. Set UPDATE_MANIFEST_URL in resources/update.json manually.');
  }

  const token = getGithubToken();
  if (!token) {
    throw new Error('Missing GitHub token. Set env var GITHUB_TOKEN (or GH_TOKEN) with "repo" scope to upload release assets.');
  }

  console.log(`Building v${version}...`);
  const forgeJs = path.join(root, 'node_modules', '@electron-forge', 'cli', 'dist', 'electron-forge.js');
  if (!fs.existsSync(forgeJs)) {
    throw new Error('Electron Forge not found. Run: npm install');
  }
  execFileInherit(process.execPath, [forgeJs, 'make'], { cwd: root });

  const outDir = path.join(root, 'out', 'make');
  if (!fs.existsSync(outDir)) throw new Error(`Build output not found: ${outDir}`);

  const files = walk(outDir);
  const exe = files.find((p) => /\.exe$/i.test(p) && /setup/i.test(path.basename(p)));
  if (!exe) throw new Error('Could not find Setup.exe in out/make');

  const rawBase = `https://raw.githubusercontent.com/${gh.owner}/${gh.repo}/${branch}`;
  const manifestUrl = `${rawBase}/client/releases/latest.json`;
  const tag = `v${version}`;
  const exeName = path.basename(exe);
  const downloadUrl = `https://github.com/${gh.owner}/${gh.repo}/releases/download/${encodeURIComponent(tag)}/${encodeURIComponent(exeName)}`;

  const manifest = {
    version,
    channel: branch,
    publishedAt: new Date().toISOString(),
    downloadUrl,
    file: exeName,
    sha256: sha256(exe),
    size: fs.statSync(exe).size,
  };

  writeJson(path.join(root, 'releases', `v${version}`, 'manifest.json'), manifest);
  writeJson(path.join(root, 'releases', 'latest.json'), manifest);
  writeJson(path.join(root, 'resources', 'update.json'), { manifestUrl });

  console.log(`Publishing manifests to branch "${branch}"...`);
  execInherit('git add -A releases resources/update.json', { cwd: root });

  const status = exec('git status --porcelain', { cwd: root });
  if (!status) {
    console.log('Nothing to commit.');
  } else {
    execInherit(`git commit -m "release: v${version}"`, { cwd: root });
  }

  execInherit(`git tag -f ${tag}`, { cwd: root });
  execInherit('git push origin HEAD', { cwd: root });
  execInherit('git push origin --tags -f', { cwd: root });

  console.log(`Creating GitHub release ${tag}...`);
  const rel = await getOrCreateRelease({ owner: gh.owner, repo: gh.repo, tag, token });
  await deleteReleaseAssetByName({ owner: gh.owner, repo: gh.repo, releaseId: rel.id, name: exeName, token });
  await uploadReleaseAsset({ uploadUrl: rel.upload_url, filePath: exe, token });

  console.log(`Done. Manifest: ${manifestUrl}`);
}

main().catch((e) => {
  console.error(e instanceof Error ? e.message : e);
  process.exit(1);
});
