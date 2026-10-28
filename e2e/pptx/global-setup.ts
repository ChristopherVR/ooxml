import { demoInputs, demoIsFresh } from '../../viewers/pptx/e2e-build-inputs';
import { neededBindings } from '../../viewers/pptx/e2e-servers';
import { assertDistFreshness } from './dist-freshness';
import { generateAllFixtures } from './fixtures/generate-all';

export default async function globalSetup() {
	// Before anything else: a stale dist means the run tests code that is not on
	// disk, and can report a spurious PASS. Fail with the build command instead.
	await assertDistFreshness();
	if (process.env.PPTX_E2E_SERVE === 'preview' && process.env.GITHUB_ACTIONS !== 'true') {
		const hash = demoInputs();
		const stale = neededBindings().filter((binding) => !demoIsFresh(binding, hash));
		if (stale.length)
			throw new Error(
				`Stale browser builds (${stale.join(', ')}): run bun e2e-build.ts in viewers/pptx`,
			);
	}
	// The build job already checks every generated deck against the committed bytes.
	// Preview preparation regenerates locally; development runs regenerate here.
	if (
		process.env.PPTX_E2E_SERVE !== 'preview' &&
		!(process.env.GITHUB_ACTIONS === 'true' && process.env.PPTX_E2E_FIXTURES_VERIFIED === '1')
	) {
		await generateAllFixtures();
	}
}
