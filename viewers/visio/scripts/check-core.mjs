import { execFileSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const core = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..', 'ooxml');
// The viewer relies on unreleased local format code. Verify that code in CI too,
// rather than treating a successful viewer bundle as parser conformance evidence.
for (const script of ['typecheck:strict', 'typecheck:pptx']) {
	execFileSync('npm', ['run', script], { cwd: core, stdio: 'inherit' });
}
execFileSync(
	process.execPath,
	[resolve(core, 'node_modules/vitest/vitest.mjs'), 'run', 'src/visio'],
	{
		cwd: core,
		stdio: 'inherit',
	},
);
