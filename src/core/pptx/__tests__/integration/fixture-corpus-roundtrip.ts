import { describe, it, expect, beforeAll } from 'vitest';

import {
	chartFlavourMismatches,
	describeIssues,
	malformedParts,
	partNames,
	roundTrip,
	validationErrors,
} from './fixture-corpus-harness';
import type { RoundTrip } from './fixture-corpus-harness';
import { FIXTURE_MANIFEST } from './fixture-corpus-manifest';
import type { FixtureEntry } from './fixture-corpus-manifest';

import type { ValidationIssue } from '../../core/utils/pptx-validator';

const TIMEOUT = 120_000;

/** Element counts per slide, the cheapest signal that a shape was dropped or duplicated. */
function elementCounts(slides: RoundTrip['original']): number[] {
	return slides.map((slide) => slide.elements.length);
}

/**
 * The `type` discriminant of every element, per slide, in document order.
 *
 * Counts alone cannot see a RECLASSIFICATION: a shape that goes in as one kind
 * and comes back as another keeps the total the same while changing what the
 * markup means and what every renderer and editor does with it. Two such
 * defects are known (see the manifest ledger), and both are invisible to a
 * count-only assertion.
 */
function elementTypeShape(slides: RoundTrip['original']): string[] {
	return slides.map((slide) => slide.elements.map((el) => el.type).join(','));
}

/**
 * Assert a check, honouring a declared known defect. A known defect must
 * actually still be broken: if it starts passing the entry is stale and the
 * test says so, which is what keeps the ledger honest.
 */
function expectOrKnownDefect(entry: FixtureEntry, check: string, assertion: () => void): void {
	const known = entry.knownDefects?.[check];
	if (!known) {
		assertion();
		return;
	}
	let threw = false;
	try {
		assertion();
	} catch {
		threw = true;
	}
	if (!threw) {
		throw new Error(
			`${entry.file}: knownDefects.${check} says this is broken, but it passed. ` +
				`If it has been fixed, delete the entry from fixture-corpus-manifest.ts. ` +
				`Recorded cause: ${known}`,
		);
	}
}

export function describeFixtureRoundTrips(group: 'corpus' | 0 | 1 | 2 | 3): void {
	const all = FIXTURE_MANIFEST.filter((entry) => entry.status === 'roundtrip');
	const e2e = all.filter((entry) => entry.dir === 'e2e');
	const entries =
		group === 'corpus'
			? all.filter((entry) => entry.dir === 'corpus')
			: e2e.filter((_, index) => index % 4 === group);
	for (const entry of entries) {
		describe(`fixture corpus: ${entry.file}`, () => {
			let rt: RoundTrip;
			let before: Promise<ValidationIssue[]> | undefined;
			let after: Promise<ValidationIssue[]> | undefined;
			const originalErrors = () => (before ??= validationErrors(rt.originalBytes));
			const savedErrors = () => (after ??= validationErrors(rt.savedBytes));

			beforeAll(async () => {
				rt = await roundTrip(entry);
			}, TIMEOUT);

			it('reopens through our own loader with the same slide count', () => {
				expect(rt.reloaded).toHaveLength(rt.original.length);
				expect(rt.reloaded.length).toBeGreaterThan(0);
			});

			it('preserves the part list', () => {
				const before = partNames(rt.before);
				const after = new Set(partNames(rt.after));
				const allowed = new Set(entry.allowedPartLoss ?? []);
				const lost = before.filter((name) => !after.has(name) && !allowed.has(name));
				expect(lost).toStrictEqual([]);

				// An allowance that no longer applies is dead weight: drop it.
				const stillPresent = [...allowed].filter((name) => after.has(name));
				expect(stillPresent, 'allowedPartLoss lists parts that are not lost').toStrictEqual([]);
			});

			it('keeps a consistent package: no dangling content types or relationships', async () => {
				const errors = await savedErrors();
				const opc = errors.filter(
					(e) =>
						e.code === 'MISSING_CONTENT_TYPE' ||
						e.code === 'UNRESOLVED_RELATIONSHIP' ||
						e.code === 'MISSING_REQUIRED_FILE',
				);
				expect(describeIssues(opc)).toStrictEqual([]);
			});

			it('keeps per-slide element counts stable', () => {
				expectOrKnownDefect(entry, 'elementCountStable', () => {
					expect(elementCounts(rt.reloaded)).toStrictEqual(elementCounts(rt.original));
				});
			});

			it('keeps per-slide element types stable', () => {
				expectOrKnownDefect(entry, 'elementTypeStable', () => {
					expect(elementTypeShape(rt.reloaded)).toStrictEqual(elementTypeShape(rt.original));
				});
			});

			it(
				'writes well-formed XML in every part',
				async () => {
					await expect(malformedParts(rt.after)).resolves.toStrictEqual([]);
				},
				TIMEOUT,
			);

			it('binds every chart part to the content type its root element implies', async () => {
				await expect(chartFlavourMismatches(rt.after)).resolves.toStrictEqual([]);
			});

			/**
			 * The validator only earns the right to gate saves if it is honest
			 * about genuine input, so this asserts BOTH directions: a deck that was
			 * clean going in must be clean coming out, and a deck that was not must
			 * not get worse. See `pptx-validator-facet-constants.ts` for the three
			 * rule families that had to be corrected before this could be asserted
			 * at all.
			 */
			it(
				'introduces no new validation errors',
				async () => {
					const before = await originalErrors();
					const after = await savedErrors();
					const beforeCodes = new Set(before.map((e) => `${e.code}|${e.message}`));
					const introduced = after.filter((e) => !beforeCodes.has(`${e.code}|${e.message}`));
					expectOrKnownDefect(entry, 'validationErrorsNotIntroduced', () => {
						expect(describeIssues(introduced)).toStrictEqual([]);
					});
				},
				TIMEOUT,
			);

			it(
				'is a valid package before it is saved',
				async () => {
					expect(describeIssues(await originalErrors())).toStrictEqual([]);
				},
				TIMEOUT,
			);

			it(
				'is a valid package after it is saved',
				async () => {
					const after = describeIssues(await savedErrors());
					expectOrKnownDefect(entry, 'validationErrorsNotIntroduced', () => {
						expect(after).toStrictEqual([]);
					});
				},
				TIMEOUT,
			);
		});
	}
}
