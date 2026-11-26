/**
 * The corpus-wide load -> save -> reload -> validate harness.
 *
 * ## What it is for
 *
 * Eight defects that made PowerPoint refuse to open a saved deck shipped under
 * a green suite because no test ever opened a saved file. Every one of them is
 * visible from a plain no-edit round trip of a real deck, and every one of them
 * is invisible to a model-level assertion, because re-parsing our own output
 * forgives the damage. This harness runs that round trip over EVERY deck in
 * `fixture-corpus-manifest.ts` and asserts, per fixture:
 *
 *   1. the saved package still opens through our own loader;
 *   2. the part list is preserved (allowed losses are declared per fixture);
 *   3. per-slide element counts are stable;
 *   4. every XML part in the output is well-formed;
 *   5. `validatePptx` reports no new errors, and none at all on a deck that
 *      was clean going in.
 *
 * ## Running it
 *
 *     bun run --filter pptx-viewer-core test -- fixture-corpus-roundtrip
 *
 * The whole binary corpus is 130 MB.
 *
 * ## PowerPoint acceptance
 *
 * Passing here is necessary, not sufficient: our loader is more forgiving than
 * PowerPoint's. `scripts/com-acceptance.mjs` drives the real application over
 * the same manifest and is the ground truth. It is deliberately NOT part of
 * `bun run test`, because it needs a local PowerPoint install. Run it after any
 * change to the save pipeline:
 *
 *     bun run scripts/com-acceptance.mjs
 *
 * @see fixture-corpus-manifest.ts for provenance and the accepted deviations.
 * @see save-invariants.test.ts for the specific corruption classes this repo
 *      has already shipped once.
 */
import { describe, it, expect } from 'vitest';

import { readFixture, validationErrors } from './fixture-corpus-harness';
import { FIXTURE_MANIFEST, listFixturesOnDisk, manifestFor } from './fixture-corpus-manifest';
import { describeFixtureRoundTrips } from './fixture-corpus-roundtrip';

const TIMEOUT = 120_000;

/**
 * The presence guard, and the reason `it.skipIf(!existsSync(fixture))` is
 * banned in this directory: a skipped test is a green test, so deleting a
 * fixture used to reduce coverage silently. Here the manifest and the
 * directory must agree exactly, in both directions.
 */
describe('fixture corpus: the manifest and the tree agree', () => {
	for (const dir of ['e2e', 'corpus'] as const) {
		it(`every .pptx / .ppt in ${dir} is declared, and every declaration exists`, () => {
			expect(listFixturesOnDisk(dir)).toStrictEqual(manifestFor(dir).map((e) => e.file));
		});
	}

	it('keeps genuine, non-synthetic decks in the corpus', () => {
		// Synthetic decks are our serializer feeding our parser, so they cannot
		// witness a disagreement between us and PowerPoint. If this ever drops
		// to zero the harness has stopped proving anything about real files.
		const genuine = FIXTURE_MANIFEST.filter((e) => e.provenance !== 'synthetic');
		expect(genuine.length).toBeGreaterThanOrEqual(15);
		expect(genuine.filter((e) => e.provenance === 'powerpoint').length).toBeGreaterThanOrEqual(8);
	});

	it('documents a cause for every known defect', () => {
		for (const entry of FIXTURE_MANIFEST) {
			for (const [check, reason] of Object.entries(entry.knownDefects ?? {})) {
				if (reason.length <= 40) {
					throw new Error(
						`${entry.file} knownDefects.${check} needs a written cause, got: ${reason}`,
					);
				}
			}
		}
	});
});

describe('fixture corpus: an encrypted package is rejected, not mis-parsed', () => {
	for (const entry of FIXTURE_MANIFEST.filter((e) => e.status === 'encrypted')) {
		it(
			`${entry.file} fails to open with a diagnosable error`,
			async () => {
				const { PptxHandler } = await import('../../core/PptxHandler');
				await expect(new PptxHandler().load(readFixture(entry))).rejects.toThrow();

				// The validator must say WHY rather than crash or pass it.
				const errors = await validationErrors(readFixture(entry));
				expect(errors.map((e) => e.code)).toContain('INVALID_ZIP');
			},
			TIMEOUT,
		);
	}
});

/**
 * Legacy binary `.ppt` cannot go through the load -> save -> reload comparison
 * (it loads as a compound file and saves as an OpenXML package, deliberately),
 * so what this asserts is the claim the product now makes on its behalf: the
 * same `PptxHandler.load()` opens it, and it saves as a real `.pptx`. The
 * fidelity comparison against the `.pptx` twin lives in `ppt-import.test.ts`.
 */
describe('fixture corpus: a legacy .ppt imports and saves as .pptx', () => {
	for (const entry of FIXTURE_MANIFEST.filter((e) => e.status === 'legacy-ppt')) {
		it(
			`${entry.file} opens through the ordinary load path`,
			async () => {
				const { PptxHandler } = await import('../../core/PptxHandler');
				const handler = new PptxHandler();
				const data = await handler.load(readFixture(entry));
				expect(data.slides.length).toBeGreaterThan(0);

				// We never write the binary format back. The saved bytes must be a
				// ZIP (PK\x03\x04), which is what makes re-extensioning the save
				// name to `.pptx` the honest thing to do rather than a cosmetic one.
				const saved = await handler.save(data.slides);
				expect([saved[0], saved[1], saved[2], saved[3]]).toStrictEqual([0x50, 0x4b, 0x03, 0x04]);
			},
			TIMEOUT,
		);
	}
});

describeFixtureRoundTrips('corpus');
