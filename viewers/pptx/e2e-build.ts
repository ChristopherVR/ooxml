/** Build the demo artifacts once for local verification and all CI browser shards. */
import { spawnSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname } from 'node:path';

import { generateAllFixtures } from '../../e2e/pptx/fixtures/generate-all';
import { demoDir, demoInputs, demoIsFresh, demoStamp } from './e2e-build-inputs';
import { neededBindings } from './e2e-servers';

await generateAllFixtures();
const hash = demoInputs();
for (const binding of neededBindings()) {
	if (demoIsFresh(binding, hash)) {
		console.log(`${binding}: browser build inputs unchanged`);
		continue;
	}
	rmSync(demoStamp(binding), { force: true });
	const result = spawnSync('bun', ['run', 'build'], {
		cwd: demoDir(binding),
		env: { ...process.env, DEMO_BASE: '/', NODE_ENV: 'production', VITE_PPTX_E2E_HANDLE: '1' },
		stdio: 'inherit',
	});
	if (result.error) throw result.error;
	if (result.status !== 0) process.exit(result.status ?? 1);
	mkdirSync(dirname(demoStamp(binding)), { recursive: true });
	writeFileSync(demoStamp(binding), hash);
}
