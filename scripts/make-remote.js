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

function main() {
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

  console.log(`Building v${version}...`);
  execInherit('npm run make', { cwd: root });

  const outDir = path.join(root, 'out', 'make');
  if (!fs.existsSync(outDir)) throw new Error(`Build output not found: ${outDir}`);

  const files = walk(outDir);
  const exe = files.find((p) => /\.exe$/i.test(p) && /setup/i.test(path.basename(p)));
  if (!exe) throw new Error('Could not find Setup.exe in out/make');

  const releaseDir = path.join(root, 'releases', `v${version}`);
  ensureDir(releaseDir);

  const exeName = path.basename(exe);
  const destExe = path.join(releaseDir, exeName);
  fs.copyFileSync(exe, destExe);

  const relPath = path.posix.join('client', 'releases', `v${version}`, exeName);
  const rawBase = `https://raw.githubusercontent.com/${gh.owner}/${gh.repo}/${branch}`;
  const downloadUrl = `${rawBase}/${relPath}`;
  const manifestUrl = `${rawBase}/client/releases/latest.json`;

  const manifest = {
    version,
    channel: branch,
    publishedAt: new Date().toISOString(),
    downloadUrl,
    file: relPath,
    sha256: sha256(destExe),
    size: fs.statSync(destExe).size,
  };

  writeJson(path.join(root, 'releases', `v${version}`, 'manifest.json'), manifest);
  writeJson(path.join(root, 'releases', 'latest.json'), manifest);
  writeJson(path.join(root, 'resources', 'update.json'), { manifestUrl });

  console.log(`Publishing artifacts to branch "${branch}"...`);
  execInherit('git add -A releases resources/update.json', { cwd: root });

  const status = exec('git status --porcelain', { cwd: root });
  if (!status) {
    console.log('Nothing to commit.');
  } else {
    execInherit(`git commit -m "release: v${version}"`, { cwd: root });
  }

  try {
    execInherit(`git tag -f v${version}`, { cwd: root });
  } catch {
    // ignore
  }

  execInherit('git push origin HEAD', { cwd: root });
  execInherit('git push origin --tags -f', { cwd: root });

  console.log(`Done. Manifest: ${manifestUrl}`);
}

main();

