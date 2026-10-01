// The pptx tests read a set of decks that belong to the pptx-viewer end-to-end suite
// (`e2e/fixtures`). They are not stored in this repository (about 100 MB of binaries); this script
// copies them from a pptx-viewer checkout into the git-ignored `src/pptx/__tests__/fixtures/e2e`.
//
// Source lookup: $PPTX_VIEWER_E2E_FIXTURES, else `../pptx-viewer/e2e/fixtures` next to this repo.
// Returns true when the fixtures are present afterwards.
import { cpSync, existsSync, readdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
export const target = resolve(root, 'src/pptx/__tests__/fixtures/e2e');

export function syncE2eFixtures({ quiet = false } = {}) {
	const present = () => existsSync(target) && readdirSync(target).some((f) => f.endsWith('.pptx'));
	if (present()) return true;
	const source = resolve(
		process.env.PPTX_VIEWER_E2E_FIXTURES ?? resolve(root, '../pptx-viewer/e2e/fixtures'),
	);
	if (!existsSync(source)) {
		if (!quiet)
			console.warn(
				`[ooxml-core] pptx e2e fixtures not found (looked in ${source}); tests that read them are skipped. ` +
					'Set PPTX_VIEWER_E2E_FIXTURES or check out pptx-viewer next to this repository.',
			);
		return false;
	}
	cpSync(source, target, { recursive: true });
	if (!quiet) console.log(`[ooxml-core] copied pptx e2e fixtures from ${source}`);
	return present();
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url))
	process.exitCode = syncE2eFixtures() ? 0 : 1;
