// `npm pack` / `npm publish` run prepack: build only when dist is missing.
import { existsSync } from 'node:fs';
import { spawnSync } from 'node:child_process';

if (!existsSync('dist/index.js')) {
	const result = spawnSync('node', ['scripts/build.mjs'], { stdio: 'inherit' });
	process.exit(result.status ?? 1);
}
