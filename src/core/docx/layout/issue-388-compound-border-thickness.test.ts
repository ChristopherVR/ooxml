import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';

async function layout(name: string) {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	return layoutDocumentModel(model, createFakeMeasurer());
}

/** Bordered paragraphs of a layout, in document order. */
async function borders(name: string) {
	const result = await layout(name);
	return result.pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph' && block.frame?.borders?.top && block.frame.borders.bottom
					? [block.frame.borders]
					: [],
			),
		),
	);
}

const pt = (px: number) => (px * 72) / 96;

/**
 * Word's bands (pt), read from the vector paths of its own PDF export of the fixtures, band then
 * gap, from the outer edge in (`issue-388-word-compound-bands.json`). Key: style|w:sz|w:space.
 */
const WORD = JSON.parse(
	await readFile(
		new URL('./fixtures/word-layout-reports/issue-388-word-compound-bands.json', import.meta.url),
		'utf8',
	),
) as Record<string, { seq: number[]; box: number }>;

const STYLES = [
	'double',
	'thinThickSmallGap',
	'thickThinSmallGap',
	'thinThickThinSmallGap',
	'thinThickMediumGap',
	'thickThinMediumGap',
	'thinThickThinMediumGap',
	'thinThickLargeGap',
	'thickThinLargeGap',
	'thinThickThinLargeGap',
	'triple',
];

// w:sz 3 is the one size Word narrows below the rule (0.24 pt, not 0.36 pt); it is left out.
const FILES = [
	['fit-space0', 0, [2, 4, 6, 8, 12, 18, 24]],
	['fit-space4', 4, [2, 4, 6, 8, 12, 18, 24]],
	['held-space0', 0, [10, 16, 20]],
	['held-space4', 4, [10, 16, 20]],
] as const;

describe('compound paragraph borders against Word (#388)', () => {
	for (const [file, space, sizes] of FILES) {
		it(`draws every band and gap at Word's width and order (${file})`, async () => {
			const all = await borders(`issue-388-word-bands-${file}`);
			const fixtureSizes = file.startsWith('fit') ? [2, 4, 6, 8, 12, 18, 24] : [3, 10, 16, 20];
			expect(all).toHaveLength(STYLES.length * fixtureSizes.length);
			let checked = 0;
			STYLES.forEach((style, styleIndex) => {
				fixtureSizes.forEach((sz, sizeIndex) => {
					if (!(sizes as readonly number[]).includes(sz)) return;
					const word = WORD[`${style}|${sz}|${space}`];
					const sides = all[styleIndex * fixtureSizes.length + sizeIndex];
					expect(word, `${style} ${sz}`).toBeDefined();
					for (const side of ['top', 'bottom'] as const) {
						const compound = (sides?.[side]?.compound ?? []).map(pt);
						expect(compound, `${style} ${sz} ${side}`).toHaveLength(word?.seq.length ?? -1);
						compound.forEach((width, index) =>
							expect(
								Math.abs(width - (word?.seq[index] ?? NaN)),
								`${style} ${sz} ${side}[${index}]`,
							).toBeLessThan(0.035),
						);
						expect(pt(sides?.[side]?.spacePx ?? NaN)).toBeCloseTo(space, 6);
					}
					checked += 1;
				});
			});
			expect(checked).toBe(STYLES.length * sizes.length);
		});
	}
});

it('keeps single borders at their nominal width (#388 control)', async () => {
	const result = await layout('issue-388-single-top-borders');
	const totals = result.pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph' && block.frame?.borders?.top
					? [block.frame.borders.top.widthPx * 0.75]
					: [],
			),
		),
	);
	expect(totals.slice(0, 7).every((total) => Math.abs(total - 1.5) < 0.01)).toBe(true);
	expect(totals.slice(7).every((total) => Math.abs(total - 3) < 0.01)).toBe(true);
});
