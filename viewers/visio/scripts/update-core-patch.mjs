import { execFileSync } from 'node:child_process';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { coreDirectory as core, viewerRoot as root } from './core-paths.mjs';
import { coreSnapshot } from './core-snapshot.mjs';

const git = (args) => execFileSync('git', args, { cwd: core, encoding: 'utf8' });
const pinned = readFileSync(resolve(root, 'integration/core-revision.txt'), 'utf8').trim();
if (git(['rev-parse', 'HEAD']).trim() !== pinned)
	throw new Error(
		'Core HEAD differs from pinned revision. Review and update the integration baseline deliberately.',
	);
if (git(['diff', '--cached', '--name-only']).trim())
	throw new Error('Core has staged changes. Leave index ownership with its author.');
const patch = coreSnapshot(core, pinned);
const target = resolve(root, 'integration/ooxml-visio.patch');
if (patch) writeFileSync(target, patch);
else rmSync(target, { force: true });
console.log(
	`Updated optional core patch: ${Buffer.byteLength(patch)} bytes. Validate a clean setup before publishing.`,
);
