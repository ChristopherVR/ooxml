import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, copyFileSync, mkdtempSync } from 'node:fs';
import { homedir, tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { isDeepStrictEqual } from 'node:util';
import { coreDirectory as core, viewerRoot as root } from './core-paths.mjs';
import { canonicalPatch, coreSnapshot } from './core-snapshot.mjs';

const revision = readFileSync(resolve(root, 'integration/core-revision.txt'), 'utf8').trim();
const patchPath = resolve(root, 'integration/ooxml-visio.patch');
const expectedPatch = existsSync(patchPath) ? canonicalPatch(readFileSync(patchPath, 'utf8')) : '';
const lockPath = resolve(root, 'integration/core-package-lock.json');
const git = (args, cwd = core) =>
	execFileSync('git', args, {
		cwd,
		stdio: 'inherit',
		env: { ...process.env, GIT_TERMINAL_PROMPT: '0' },
	});
const output = (args) => execFileSync('git', args, { cwd: core, encoding: 'utf8' });
if (!/^[a-f0-9]{40}$/.test(revision)) throw new Error('Invalid pinned core revision.');
if (!existsSync(core)) {
	git(['clone', 'https://github.com/ChristopherVR/ooxml.git', core], root);
	git(['checkout', '--detach', revision]);
	if (expectedPatch) {
		git(['apply', '--check', patchPath]);
		git(['apply', patchPath]);
	}
}
// Never reset, patch or build an unrelated/partially modified existing checkout.
if (output(['rev-parse', 'HEAD']).trim() !== revision)
	throw new Error(
		'Existing ooxml checkout has a different revision. Use an isolated checkout at integration/core-revision.txt with VISIO_CORE_DIR. Nothing was overwritten.',
	);
if (output(['diff', '--cached', '--name-only']).trim())
	throw new Error('Existing ooxml checkout has staged changes. Nothing was overwritten.');
if (coreSnapshot(core, revision) !== expectedPatch)
	throw new Error(
		'Existing ooxml changes do not match the optional integration patch (or clean published baseline). Nothing was overwritten. Rebuild intentionally modified core manually, or review and regenerate the patch.',
	);
const installedLock = resolve(core, 'package-lock.json');
if (
	existsSync(installedLock) &&
	!isDeepStrictEqual(
		JSON.parse(readFileSync(installedLock, 'utf8')),
		JSON.parse(readFileSync(lockPath, 'utf8')),
	)
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
console.log(
	'Pinned Visio core ready. Run npm ci and npm ci --prefix packages/bindings. If using VISIO_CORE_DIR, run npm run link:core with the same override after npm ci.',
);
