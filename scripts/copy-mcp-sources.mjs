// Copy an MCP package into a clean consumer without Bun's workspace dependency links. A dependency
// on a sibling released from this repository (`workspace:*` or `*`) is installed from a tarball of
// the local package, which is what that release publishes alongside it, so an MCP change can use
// core API that is not on npm yet. Everything else comes from the registry, as a user gets it.
import { execFileSync } from 'node:child_process';
import { cpSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

import { PACKAGES, isWorkspaceRange } from './release-plan.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const FIELDS = ['dependencies', 'peerDependencies', 'optionalDependencies'];

const [source, target] = process.argv.slice(2);
if (!source || !target) throw new Error('Expected source and target directories');
const root = resolve(source);
const out = resolve(target);
cpSync(root, out, {
	recursive: true,
	filter: (file) => !relative(root, file).split(sep).includes('node_modules'),
});

const manifestPath = join(out, 'package.json');
const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));
const dirs = new Map(Object.values(PACKAGES).map((meta) => [meta.npm, meta.dir]));
const tarballs = join(out, '.workspace');
let rewritten = false;
for (const field of FIELDS) {
	for (const [dep, range] of Object.entries(manifest[field] ?? {})) {
		const dir = dirs.get(dep);
		if (!isWorkspaceRange(range) || !dir) continue;
		mkdirSync(tarballs, { recursive: true });
		// Built already (CI builds the core before the viewers); packing must not rebuild it.
		const [packed] = JSON.parse(
			execFileSync(
				'npm',
				['pack', join(ROOT, dir), '--pack-destination', tarballs, '--ignore-scripts', '--json'],
				{ encoding: 'utf8', shell: process.platform === 'win32' },
			),
		);
		manifest[field][dep] = `file:./.workspace/${packed.filename}`;
		rewritten = true;
	}
}
if (rewritten) writeFileSync(manifestPath, `${JSON.stringify(manifest, null, '\t')}\n`);
