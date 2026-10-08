/**
 * CI guard: regenerate every generated fixture and fail if any committed deck
 * changed.
 *
 * The generators save through the core's own pipeline, so a core change that
 * alters what a save writes (a reordered part, a new app.xml statistic)
 * changes the decks too. Unless the refreshed decks are committed with that
 * change, every later Playwright run rewrites them and the working tree never
 * comes back clean. Generation itself is deterministic
 * (`writeFixtureDeterministic`), so any difference here is real drift.
 *
 * Needs the core and pptx packages built. Run: `bun run fixtures:check` in
 * `viewers/pptx`; commit what it rewrites to fix a failure.
 */
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

import { generateAllFixtures } from './generate-all';

const fixturesDir = fileURLToPath(new URL('.', import.meta.url));

await generateAllFixtures();

// Only generated output counts: a generator's own source is reviewed as code.
const drift = execFileSync('git', ['status', '--porcelain', '--', '.', ':!*.ts'], {
	cwd: fixturesDir,
	encoding: 'utf8',
}).trimEnd();

if (drift) {
	console.error(
		'Generated e2e fixtures differ from the committed ones. Commit the regenerated decks:\n' +
			drift,
	);
	process.exit(1);
}
console.log('Generated e2e fixtures match the committed decks.');
