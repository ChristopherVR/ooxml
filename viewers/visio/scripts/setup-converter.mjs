import { runNpm } from './npm-command.mjs';
import { canonicalPatch, converterSnapshot } from './converter-snapshot.mjs';
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, copyFileSync, mkdtempSync } from 'node:fs';
import { homedir, tmpdir } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { isDeepStrictEqual } from 'node:util';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const core = resolve(root, '..', 'emf-converter-current');
const revision = readFileSync(
	resolve(root, 'integration/emf-converter-revision.txt'),
	'utf8',
).trim();
const patchPath = resolve(root, 'integration/emf-converter-visio.patch');
const expectedPatch = readFileSync(patchPath, 'utf8').replace(/\r\n/g, '\n');
const lockPath = resolve(root, 'integration/emf-converter-package-lock.json');
const git = (args, cwd = core) =>
	execFileSync('git', args, {
		cwd,
		stdio: 'inherit',
		env: { ...process.env, GIT_TERMINAL_PROMPT: '0' },
	});
const output = (args) => execFileSync('git', args, { cwd: core, encoding: 'utf8' });
if (!/^[a-f0-9]{40}$/.test(revision)) throw new Error('Invalid pinned converter revision.');
if (!existsSync(core)) {
	git(['clone', 'https://github.com/ChristopherVR/emf-converter.git', core], root);
	git(['checkout', '--detach', revision]);
	git(['apply', '--check', patchPath]);
	git(['apply', patchPath]);
}
// Refuse to mutate or build an unrelated/partially patched existing checkout.
const patchedRevision = readFileSync(
	resolve(root, 'integration/emf-converter-patched-revision.txt'),
	'utf8',
).trim();
if (!/^[a-f0-9]{40}$/.test(patchedRevision)) throw new Error('Invalid patched converter revision.');
if (![revision, patchedRevision].includes(output(['rev-parse', 'HEAD']).trim()))
	throw new Error(
		'Existing isolated converter checkout has a different revision. Use an isolated sibling checkout at integration/emf-converter-revision.txt. Nothing was overwritten.',
	);
if (output(['diff', '--cached', '--name-only']).trim())
	throw new Error(
		'Existing isolated converter checkout has staged changes. Nothing was overwritten.',
	);
const actualPatch = converterSnapshot(core, revision);
if (actualPatch !== canonicalPatch(expectedPatch))
	throw new Error(
		'Existing isolated converter changes do not match the included Visio patch. Nothing was overwritten. Use an isolated clean checkout, or rebuild your intentionally modified converter manually.',
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
		'Existing converter npm lock differs from the pinned integration lock. Nothing was overwritten.',
	);
if (!existsSync(installedLock)) copyFileSync(lockPath, installedLock);
const cache =
	process.env.VISIO_NPM_CACHE ??
	(existsSync(homedir()) ? undefined : mkdtempSync(resolve(tmpdir(), 'visio-converter-cache-')));
runNpm(['ci', '--ignore-scripts', ...(cache ? ['--cache', cache] : [])], {
	cwd: core,
	stdio: 'inherit',
});
runNpm(['run', 'build'], { cwd: core, stdio: 'inherit' });
console.log(
	'Pinned local converter built. No application dependency or live conversion was changed. Run npm run check:converter for its regression suite.',
);
