import { execFileSync, spawnSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..'),
	core = resolve(root, '..', 'ooxml');
const git = (args) => execFileSync('git', args, { cwd: core, encoding: 'utf8' });
const pinned = readFileSync(resolve(root, 'integration/core-revision.txt'), 'utf8').trim();
if (git(['rev-parse', 'HEAD']).trim() !== pinned)
	throw new Error(
		'Core HEAD differs from pinned revision. Review and update the integration baseline deliberately.',
	);
if (git(['diff', '--cached', '--name-only']).trim())
	throw new Error('Core has staged changes. Leave index ownership with its author.');
let patch = git(['diff', '--binary']);
for (const file of git(['ls-files', '--others', '--exclude-standard'])
	.trim()
	.split('\n')
	.filter(Boolean)) {
	if (file === 'package-lock.json') continue;
	const diff = spawnSync('git', ['diff', '--no-index', '--binary', '--', '/dev/null', file], {
		cwd: core,
		encoding: 'utf8',
	});
	if (diff.status !== 0 && diff.status !== 1) throw new Error(`Could not include ${file}`);
	patch += diff.stdout;
}
writeFileSync(resolve(root, 'integration/ooxml-visio.patch'), patch);
console.log(
	`Updated core patch: ${Buffer.byteLength(patch)} bytes. Validate a clean setup before publishing.`,
);
