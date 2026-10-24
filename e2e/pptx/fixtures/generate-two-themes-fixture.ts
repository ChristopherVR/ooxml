/**
 * Generates `two-themes.pptx`: two slide masters on two theme parts, for the
 * inspector's theme selector (Presentation properties > Theme), which points
 * masters at another theme part already in the package.
 *
 * - `ppt/theme/theme1.xml` ("Tide": accent 1 `1F6FB2`, headings in Georgia),
 *   used by master 1, whose layout slides 1 and 2 use;
 * - `ppt/theme/theme2.xml` ("Ember": accent 1 `C0392B`, headings in Courier
 *   New), used by master 2, whose layout slide 3 uses;
 * - every slide has an `accent1` rectangle named `AccentShape`.
 *
 * The deck is defined once, next to the core tests
 * (`src/core/pptx/__tests__/helpers/two-themes-deck.ts`), which the binding
 * unit tests build it from too.
 *
 * Run with: bun run e2e/pptx/fixtures/generate-two-themes-fixture.ts
 */
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { buildTwoThemesDeck } from '../../../src/core/pptx/__tests__/helpers/two-themes-deck';
import { writeFixtureDeterministic } from './write-fixture';

export { TWO_THEMES } from '../../../src/core/pptx/__tests__/helpers/two-themes-deck';

const __dirname = dirname(fileURLToPath(import.meta.url));

export async function generateTwoThemesFixture(): Promise<string> {
	const outPath = resolve(__dirname, 'two-themes.pptx');
	await writeFixtureDeterministic(outPath, await buildTwoThemesDeck());
	return outPath;
}

if (process.argv[1]?.endsWith('generate-two-themes-fixture.ts')) {
	generateTwoThemesFixture()
		.then((path) => console.log(`Wrote ${path}`))
		.catch((error) => {
			console.error(error);
			process.exit(1);
		});
}
