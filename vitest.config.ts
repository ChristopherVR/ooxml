import { readFileSync, readdirSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { syncE2eFixtures } from './scripts/pptx/sync-e2e-fixtures.mjs';

/** Test files under `dir` that read the pptx-viewer e2e decks (not stored in this repository). */
function testsNeedingE2eFixtures(dir: string): string[] {
	const found: string[] = [];
	for (const name of readdirSync(dir)) {
		const path = join(dir, name);
		if (statSync(path).isDirectory()) found.push(...testsNeedingE2eFixtures(path));
		else if (name.endsWith('.test.ts') && readFileSync(path, 'utf8').includes('fixtures/e2e'))
			found.push(path.replace(/\\/g, '/'));
	}
	return found;
}

// Fill the git-ignored e2e fixture folder from a pptx-viewer checkout when one is available;
// otherwise skip exactly the tests that need it.
const exclude = syncE2eFixtures() ? [] : testsNeedingE2eFixtures('src/pptx');
if (exclude.length)
	console.warn(`[ooxml-core] skipping ${exclude.length} test files that need the e2e decks`);

export default {
	test: {
		globals: true,
		environment: 'node',
		include: ['src/**/*.test.ts'],
		exclude: ['node_modules', 'dist', ...exclude],
		// The pptx area's integration tests load real multi-megabyte decks through the full
		// parse pipeline (and several round-trip them through save); the 5s default timed out on CI.
		testTimeout: 30_000,
		hookTimeout: 30_000,
		maxWorkers: 4,
	},
};
