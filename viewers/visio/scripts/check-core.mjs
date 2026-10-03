import { runNpm } from './npm-command.mjs';
import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';
import { realpathSync } from 'node:fs';
import { coreDirectory as core, viewerRoot } from './core-paths.mjs';

if (realpathSync(resolve(viewerRoot, 'node_modules/ooxml-core')) !== realpathSync(core))
	throw new Error(
		'Installed ooxml-core does not match VISIO_CORE_DIR. Run npm run link:core with the same override after npm ci.',
	);
// Verify the selected published baseline or intentionally modified local format code,
// rather than treating a successful viewer bundle as parser conformance evidence.
for (const script of ['typecheck:strict', 'typecheck:pptx']) {
	runNpm(['run', script], { cwd: core, stdio: 'inherit' });
}
execFileSync(
	process.execPath,
	[resolve(core, 'node_modules/vitest/vitest.mjs'), 'run', 'src/visio'],
	{
		cwd: core,
		stdio: 'inherit',
	},
);
// The viewer imports the built package entry. Rebuild after source checks so
// worker, renderer and packed-consumer tests cannot accidentally use stale code.
runNpm(['run', 'build:visio'], { cwd: core, stdio: 'inherit' });
