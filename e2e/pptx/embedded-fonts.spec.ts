/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * A deck's embedded fonts, in every binding.
 *
 * The fixture is a Google Slides export that embeds Playfair Display and Lato
 * (regular, bold, italic, bold italic) as MicroType Express compressed EOT
 * `.fntdata` parts. The core returns those variants without decoding them and
 * decodes each one when a renderer first reads its data, so this checks what
 * a reader sees, whatever the decoding order: the embedded faces that match
 * their typeface are registered and load, slide text measures in them rather
 * than in a fallback, and saving writes each font part back byte for byte.
 *
 * The deck's Playfair Display files are not the faces their slots claim (the
 * regular slot holds a SemiBold file, the bold slot a Black one), so that
 * typeface is left to Google Fonts and never registered from the deck (see
 * `embedded-font-trust.ts`). Lato's four files match and are registered.
 * Google Fonts serves both families, so its endpoints are stubbed out: a face
 * found here can only have come from the deck.
 *
 * Face registration is compared across the bindings, with React as the
 * reference; see {@link EXCLUSIONS} for the known gap.
 *
 * Run: bunx playwright test embedded-fonts
 */
import { readFileSync } from 'node:fs';

import { expect, test } from '@playwright/test';
import type { Page } from '@playwright/test';

import { savePptxViaBackstage } from './save-pptx';
import { fixture, loadDeck, loadDeckAt } from './support/deck';
import { downloadBytes } from './support/exports';
import { acrossFrameworks } from './support/parity';
import { applyExclusions } from './support/parity-exclusions';
import type { ParityExclusion } from './support/parity-exclusions';
import { readZipPartBytes } from './support/pptx-xml';

const DECK = fixture('Slide_Animations_Speaker_comments_8_Slides_2_7_MB_c8f64d1a03.pptx');
/** Typefaces whose embedded files match them, and one whose files do not. */
const TRUSTED = 'Lato';
const REJECTED = 'Playfair Display';
const FONT_PARTS = [
	'ppt/fonts/PlayfairDisplay-regular.fntdata',
	'ppt/fonts/PlayfairDisplay-bold.fntdata',
	'ppt/fonts/PlayfairDisplay-italic.fntdata',
	'ppt/fonts/PlayfairDisplay-boldItalic.fntdata',
	'ppt/fonts/Lato-regular.fntdata',
	'ppt/fonts/Lato-bold.fntdata',
	'ppt/fonts/Lato-italic.fntdata',
	'ppt/fonts/Lato-boldItalic.fntdata',
];

/**
 * Known gaps, each to be deleted when its binding is fixed.
 */
const EXCLUSIONS: readonly ParityExclusion[] = [
	{
		binding: 'vanilla',
		pattern: /^no embedded @font-face registered$/u,
		why: 'The vanilla viewer never injects buildEmbeddedFontStyles output: embedded fonts reach only the glyph-outline cache and exports, so slide text uses fallback fonts.',
	},
	{
		binding: 'svelte',
		pattern: /^no embedded @font-face registered$/u,
		why: 'The svelte viewer never injects buildEmbeddedFontStyles output: embedded fonts reach only the glyph-outline cache and exports, so slide text uses fallback fonts.',
	},
];

/** Answer Google Fonts with nothing, so only the deck can supply a face. */
async function blockGoogleFonts(page: Page): Promise<void> {
	await page.route('**/fonts.googleapis.com/**', (route) =>
		route.fulfill({ status: 200, contentType: 'text/css', body: '' }),
	);
	await page.route('**/fonts.gstatic.com/**', (route) => route.abort());
}

interface FaceReport {
	family: string;
	weight: string;
	style: string;
	status: string;
}

/** Every registered face of `family`, once each has finished loading. */
async function facesOf(page: Page, family: string): Promise<FaceReport[]> {
	return page.evaluate(async (name) => {
		const faces = [...document.fonts].filter((face) => face.family.replace(/"/gu, '') === name);
		await Promise.allSettled(faces.map((face) => face.load()));
		return faces
			.map((face) => ({
				family: name,
				weight: face.weight,
				style: face.style,
				status: face.status,
			}))
			.sort((a, b) => `${a.weight}|${a.style}`.localeCompare(`${b.weight}|${b.style}`));
	}, family);
}

test.describe('embedded fonts', () => {
	test(
		'registers and loads the embedded faces that match their typeface',
		{ tag: '@cross-binding' },
		async ({ browser }, testInfo) => {
			const expected = [
				['400', 'italic'],
				['400', 'normal'],
				['700', 'italic'],
				['700', 'normal'],
			].map(([weight, style]) => ({ family: TRUSTED, weight, style, status: 'loaded' }));
			const results = await acrossFrameworks(
				browser,
				testInfo,
				async (page, origin) => {
					await blockGoogleFonts(page);
					await loadDeckAt(page, origin, DECK);
					// Faces are registered by an effect after the first paint.
					await expect
						.poll(async () => (await facesOf(page, TRUSTED)).length, { timeout: 5_000 })
						.toBe(4)
						.catch(() => undefined);
					// A loaded face only proves the CSS parsed; measuring proves the
					// decoded font data is a real font with its own metrics.
					const widths = await page.evaluate((name) => {
						const context = document.createElement('canvas').getContext('2d')!;
						context.font = `40px "${name}", monospace`;
						const embedded = context.measureText('Hamburgefonstiv 0123456789').width;
						context.font = '40px monospace';
						return { embedded, fallback: context.measureText('Hamburgefonstiv 0123456789').width };
					}, TRUSTED);
					return {
						trusted: await facesOf(page, TRUSTED),
						rejected: await facesOf(page, REJECTED),
						measuresInFace: Math.abs(widths.embedded - widths.fallback) > 1,
					};
				},
				{ concurrency: 'sequential' },
			);

			const problems: string[] = [];
			for (const { framework, value } of results) {
				const found: string[] = [];
				if (value.trusted.length === 0) {
					found.push('no embedded @font-face registered');
				} else {
					if (JSON.stringify(value.trusted) !== JSON.stringify(expected)) {
						found.push(`registered ${JSON.stringify(value.trusted)}`);
					}
					if (!value.measuresInFace) {
						found.push(`text does not measure in the embedded ${TRUSTED}`);
					}
				}
				if (value.rejected.length > 0) {
					found.push(`registered the mismatched ${REJECTED} files`);
				}
				problems.push(
					...applyExclusions(found, { binding: framework.name }, EXCLUSIONS).map(
						(problem) => `${framework.name}: ${problem}`,
					),
				);
			}
			expect(problems).toStrictEqual([]);
		},
	);

	test('saves every embedded font part byte for byte', async ({ page }) => {
		await blockGoogleFonts(page);
		await loadDeck(page, DECK);
		const saved = await downloadBytes(await savePptxViaBackstage(page));
		const source = new Uint8Array(readFileSync(DECK));
		for (const part of FONT_PARTS) {
			const [actual, original] = await Promise.all([
				readZipPartBytes(saved, part),
				readZipPartBytes(source, part),
			]);
			expect(Buffer.from(actual).equals(Buffer.from(original)), part).toBe(true);
		}
	});
});
