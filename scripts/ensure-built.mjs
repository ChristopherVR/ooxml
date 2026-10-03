// `prepack` hook: build only when dist is missing, so `npm pack` and `npm publish` after a
// build do not build a second time. Run `bun run build` first for a fresh build.
import { existsSync } from 'node:fs';
import { spawnSync } from 'node:child_process';

const needed = [
	'dist/index.mjs',
	'dist/docx/index.cjs',
	'dist/pptx/index.mjs',
	'dist/pptx/index.d.ts',
	'dist/automation/index.mjs',
	'dist/automation/node.mjs',
	'dist/pptx/automation/index.mjs',
];
if (needed.every((f) => existsSync(f))) process.exit(0);
process.exit(spawnSync('node', ['scripts/build.mjs'], { stdio: 'inherit' }).status ?? 1);
