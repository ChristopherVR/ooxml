import { converterSnapshot } from './converter-snapshot.mjs';
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..'),
	core = resolve(root, '..', 'emf-converter-current');
const git = (args) => execFileSync('git', args, { cwd: core, encoding: 'utf8' });
const pinned = readFileSync(resolve(root, 'integration/emf-converter-revision.txt'), 'utf8').trim();
const patched = readFileSync(
	resolve(root, 'integration/emf-converter-patched-revision.txt'),
	'utf8',
).trim();
if (![pinned, patched].includes(git(['rev-parse', 'HEAD']).trim()))
	throw new Error(
		'Converter HEAD differs from pinned revision. Review and update the integration baseline deliberately.',
	);
if (git(['diff', '--cached', '--name-only']).trim())
	throw new Error('Converter has staged changes. Leave index ownership with its author.');
const patch = converterSnapshot(core, pinned);
writeFileSync(resolve(root, 'integration/emf-converter-visio.patch'), patch);
console.log(
	`Updated converter patch: ${Buffer.byteLength(patch)} bytes. Validate a clean setup before publishing.`,
);
