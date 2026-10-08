import { assertDistFreshness } from './dist-freshness';
import { generateAllFixtures } from './fixtures/generate-all';

export default async function globalSetup() {
	// Before anything else: a stale dist means the run tests code that is not on
	// disk, and can report a spurious PASS. Fail with the build command instead.
	await assertDistFreshness();
	await generateAllFixtures();
}
