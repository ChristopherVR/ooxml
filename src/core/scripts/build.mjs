// Builds every area. The tsc declarations for the strict areas, the tsup bundles (strict areas and
// pptx) and the pptx declaration bundle are independent, so they run concurrently; only the
// declaration merge needs all of them.
import { rmSync } from 'node:fs';
import { rewriteEsmDeclarationImports } from '../../../scripts/esm-declarations.mjs';
import { runBuild as run } from './run-build.mjs';

rmSync('dist', { recursive: true, force: true });
rmSync('.types-pptx', { recursive: true, force: true });
await Promise.all([
	run('tsc', ['-p', 'tsconfig.build.json']),
	run('tsup', ['--config', 'tsup.config.ts']),
	run('tsup', ['--config', 'tsup.pptx.config.ts']),
	run('tsup', ['--config', 'tsup.pptx-editor.config.ts']),
	run('tsdown', ['--config', 'tsdown.pptx.config.ts']),
]);
await run('node', ['scripts/pptx/merge-declarations.mjs']);
rewriteEsmDeclarationImports('dist');
