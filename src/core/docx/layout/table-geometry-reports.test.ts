import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import type { LayoutResult, LayoutTableBox } from './result';

/** Fixtures from generate-table-geometry.mjs, following the parameters of the reports. */
async function layout(name: string): Promise<LayoutResult> {
	const bytes = await readFile(
		new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url),
	);
	const loaded = await loadDocx(new Uint8Array(bytes));
	return layoutDocumentModel(loaded.model, {
		widthOf: (text) => text.length * 5,
		lineHeightOf: () => 16,
	});
}
const tables = (result: LayoutResult, page = 0): LayoutTableBox[] =>
	result.pages[page]!.columns[0]!.blocks.filter(
		(block): block is LayoutTableBox => block.kind === 'table',
	);
const pt = (px: number) => px * 0.75;
const rowPitchPt = (table: LayoutTableBox) => pt(table.rows[1]!.yPx - table.rows[0]!.yPx);

// Text area: A4 with 1417 twip margins is 9072 twips (453.6 pt) wide.
it('#387: a centred or right-aligned table wider than the text area overhangs the margins', async () => {
	const [centred, centredMore, right] = tables(await layout('wide-tables-centered-right'));
	// Word: edge 10.6 pt left of the margin, then 35.7 pt, then 21.2 pt (right edge on the margin).
	expect(pt(centred!.xPx!)).toBeCloseTo(-10.575, 1);
	expect(pt(centredMore!.xPx!)).toBeCloseTo(-35.7, 1);
	expect(pt(right!.xPx!)).toBeCloseTo(-21.15, 1);
});

it('#387 control: tables narrower than the text area keep their positions', async () => {
	const [centred, right] = tables(await layout('narrow-tables-control'));
	expect(pt(centred!.xPx!)).toBeCloseTo(1.8, 1);
	expect(pt(right!.xPx!)).toBeCloseTo(53.6, 1);
});

it.each([
	['row-height-at-least-with-borders-sz6', 50.78],
	['row-height-at-least-with-borders-sz24', 53.04],
	['row-height-at-least-no-borders', 50],
])('#405: %s steps %s pt per row', async (name, pitch) => {
	const [table] = tables(await layout(name));
	for (let index = 1; index < table!.rows.length; index++)
		expect(
			pt(table!.rows[index]!.yPx - table!.rows[index - 1]!.yPx),
			name + ' row ' + index,
		).toBeCloseTo(pitch, 0);
});
