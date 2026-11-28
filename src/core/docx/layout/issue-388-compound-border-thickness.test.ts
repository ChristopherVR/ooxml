import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { createFakeMeasurer, layoutDocumentModel } from './index';

async function layout(name: string) {
	const { model } = await loadDocx(
		await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)),
	);
	return layoutDocumentModel(model, createFakeMeasurer());
}

/** Total thickness (pt) of the top border of each bordered paragraph, in document order. */
async function topThickness(name: string) {
	const result = await layout(name);
	return result.pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph' && block.frame?.borders?.top
					? [block.frame.borders.top.widthPx * 0.75]
					: [],
			),
		),
	);
}

// Word's totals (pt) for `w:sz` 12 then 24, from the report's PDF: single, double, thinThick and
// thickThin small gap, thinThick medium and large gap, and thinThickThin small gap.
const WORD_TOTALS = [
	[1.44, 3.0],
	[4.32, 9.0],
	[2.88, 4.44],
	[2.88, 4.44],
	[2.88, 5.88],
	[3.6, 5.16],
	[4.32, 5.88],
];

it('gives compound paragraph borders the total thickness Word draws (#388)', async () => {
	const totals = await topThickness('issue-388-compound-top-borders');
	expect(totals).toHaveLength(14);
	WORD_TOTALS.forEach(([thin, thick], index) => {
		// Word rounds the thick band of `w:sz` 12 (1.5 pt) to 1.44 pt.
		expect(Math.abs((totals[index] ?? NaN) - (thin ?? 0))).toBeLessThan(0.2);
		expect(Math.abs((totals[index + 7] ?? NaN) - (thick ?? 0))).toBeLessThan(0.2);
	});
});

it('keeps single borders at their nominal width (#388 control)', async () => {
	const totals = await topThickness('issue-388-single-top-borders');
	expect(totals.slice(0, 7).every((total) => Math.abs(total - 1.5) < 0.2)).toBe(true);
	expect(totals.slice(7).every((total) => Math.abs(total - 3) < 0.2)).toBe(true);
});
