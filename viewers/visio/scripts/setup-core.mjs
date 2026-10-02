import { execFileSync, spawnSync } from 'node:child_process';
import { existsSync, readFileSync, copyFileSync, mkdtempSync } from 'node:fs';
import { homedir, tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const core = resolve(root, '..', 'ooxml');
const revision = readFileSync(resolve(root, 'integration/core-revision.txt'), 'utf8').trim();
const patchPath = resolve(root, 'integration/ooxml-visio.patch');
const expectedPatch = readFileSync(patchPath, 'utf8').replace(/\r\n/g, '\n');
const lockPath = resolve(root, 'integration/core-package-lock.json');
const git = (args, cwd = core) => execFileSync('git', args, { cwd, stdio: 'inherit' });
const output = (args) => execFileSync('git', args, { cwd: core, encoding: 'utf8' });
if (!/^[a-f0-9]{40}$/.test(revision)) throw new Error('Invalid pinned core revision.');
if (!existsSync(core)) {
	git(['clone', 'https://github.com/ChristopherVR/ooxml.git', core], root);
	git(['checkout', '--detach', revision]);
	git(['apply', '--check', patchPath]);
	git(['apply', patchPath]);
}
// Refuse to mutate or build an unrelated/partially patched existing checkout.
if (output(['rev-parse', 'HEAD']).trim() !== revision)
	throw new Error(
		'Existing ooxml checkout has a different revision. Use an isolated sibling checkout at integration/core-revision.txt. Nothing was overwritten.',
	);
if (output(['diff', '--cached', '--name-only']).trim())
	throw new Error('Existing ooxml checkout has staged changes. Nothing was overwritten.');
let actualPatch = output(['diff', '--binary']);
for (const file of output(['ls-files', '--others', '--exclude-standard'])
	.trim()
	.split('\n')
	.filter(Boolean)) {
	if (file === 'package-lock.json') continue;
	const result = spawnSync('git', ['diff', '--no-index', '--binary', '--', '/dev/null', file], {
		cwd: core,
		encoding: 'utf8',
	});
	if (result.status !== 0 && result.status !== 1)
		throw new Error(`Cannot verify untracked core file ${file}.`);
	actualPatch += result.stdout;
}
if (actualPatch.replace(/\r\n/g, '\n') !== expectedPatch)
	throw new Error(
		'Existing ooxml changes do not match the included Visio patch. Nothing was overwritten. Use an isolated clean checkout, or rebuild your intentionally modified core manually.',
	);
const installedLock = resolve(core, 'package-lock.json');
if (
	existsSync(installedLock) &&
	readFileSync(installedLock, 'utf8') !== readFileSync(lockPath, 'utf8')
)
	throw new Error(
		'Existing core npm lock differs from the pinned integration lock. Nothing was overwritten.',
	);
if (!existsSync(installedLock)) copyFileSync(lockPath, installedLock);
const cache =
	process.env.VISIO_NPM_CACHE ??
	(existsSync(homedir()) ? undefined : mkdtempSync(resolve(tmpdir(), 'visio-core-cache-')));
execFileSync('npm', ['ci', '--ignore-scripts', ...(cache ? ['--cache', cache] : [])], {
	cwd: core,
	stdio: 'inherit',
});
execFileSync('npm', ['run', 'build:visio'], { cwd: core, stdio: 'inherit' });
console.log('Pinned Visio core ready. Run npm ci and npm ci --prefix packages/bindings.');
