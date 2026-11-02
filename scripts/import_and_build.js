import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const importer = resolve(root, 'scripts', 'import_statement.py');
const python = process.platform === 'win32' ? 'python' : 'python3';
const imported = spawnSync(python, [importer, ...process.argv.slice(2)], {
  cwd: root,
  stdio: 'inherit',
  env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
});
if (imported.error) throw imported.error;
if (imported.status !== 0) process.exit(imported.status || 1);

const vite = resolve(root, 'node_modules', 'vite', 'bin', 'vite.js');
const built = spawnSync(process.execPath, [vite, 'build'], { cwd: root, stdio: 'inherit' });
if (built.error) throw built.error;
process.exit(built.status || 0);
