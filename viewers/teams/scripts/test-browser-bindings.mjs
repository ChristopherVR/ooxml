import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

// Sequential servers keep each framework isolated without multiplying browser or Vite workers.
const root = fileURLToPath(new URL('../', import.meta.url));
const cli = createRequire(import.meta.url).resolve('@playwright/test/cli');
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	console.log(`Testing OpenTeams ${framework} workflows`);
	const result = spawnSync(process.execPath, [cli, 'test'], {
		cwd: root,
		stdio: 'inherit',
		env: { ...process.env, TEAMS_TEST_FRAMEWORK: framework },
	});
	if (result.error) throw result.error;
	if (result.status !== 0) process.exit(result.status ?? 1);
}
