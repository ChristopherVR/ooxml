import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
interface Evidence {
	name: string;
	pages: number;
	measurements: {
		borders: { role: string; rectPt: number[] }[];
		lines: { text: string; leftPt: number; baselinePt: number }[];
	}[];
}
const native = JSON.parse(
	readFileSync(
		new URL('../__fixtures__/ordinary-table-borders/native-measurements.json', import.meta.url),
		'utf8',
	),
) as { cases: Evidence[] };
it.each(native.cases.filter((c) => c.name !== 'enclosing-border-split'))(
	'matches native ordinary border row bands and source preservation: $name',
	async (sample) => {
		const bytes = new Uint8Array(
			readFileSync(
				new URL('../__fixtures__/ordinary-table-borders/' + sample.name + '.docx', import.meta.url),
			),
		);
		const loaded = await loadDocx(bytes),
			layout = layoutDocumentModel(loaded.model, {
				widthOf: (text) => text.length * 5,
				lineHeightOf: () => 16,
			});
		expect(layout.pages).toHaveLength(sample.pages);
		for (const [index, page] of layout.pages.entries()) {
			const table = page.columns[0]!.blocks.find((block) => block.kind === 'table');
			if (!table || table.kind !== 'table') throw new Error('Expected table');
			const borders = sample.measurements[index]!.borders.filter(
				(border) => border.role === 'outer',
			).map((border) => border.rectPt);
			const top = Math.min(...borders.map((rect) => rect[1]!)),
				bottom = Math.max(...borders.map((rect) => rect[3]!));
			expect(
				Math.abs(table.heightPx * 0.75 - (bottom - top)),
				sample.name + ' page ' + index,
			).toBeLessThan(0.2);
			expect(table.rows.at(-1)!.yPx + table.rows.at(-1)!.heightPx).toBeCloseTo(table.heightPx);
			for (const row of table.rows)
				for (const [column, boxes] of row.cells.entries()) {
					for (const box of boxes) {
						const text = box.lines[0]?.fragments.map((fragment) => fragment.text).join('') ?? '';
						const match = sample.measurements[index]!.lines.find((line) =>
							text.includes(line.text.trim().split(' ')[0]!),
						);
						if (
							match &&
							(text.startsWith('T0') || text.startsWith('T1') || text.startsWith('Outer'))
						) {
							const actualX =
								(table.xPx ?? 0) +
								row.geometry![column]!.xPx +
								row.geometry![column]!.paddingLeftPx +
								(box.xPx ?? 0);
							expect(Math.abs(actualX * 0.75 - (match.leftPt - 40)), text).toBeLessThan(0.2);
							expect(
								Math.abs((row.yPx + box.yPx) * 0.75 - (match.baselinePt - 9.6 - 40)),
								text,
							).toBeLessThan(0.2);
						}
					}
				}
		}
		const saved = await loaded.save(loaded.model),
			before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(saved);
		expect(await after.file('word/document.xml')!.async('string')).toBe(
			await before.file('word/document.xml')!.async('string'),
		);
		expect((await loadDocx(saved)).model).toEqual(loaded.model);
	},
);

it('reports independently cut nested note cells without changing source', async () => {
	const bytes = new Uint8Array(
		readFileSync(
			new URL(
				'../__fixtures__/ordinary-table-borders/enclosing-border-split.docx',
				import.meta.url,
			),
		),
	);
	const loaded = await loadDocx(bytes);
	const layout = layoutDocumentModel(loaded.model, {
		widthOf: (text) => text.length * 5,
		lineHeightOf: () => 16,
	});
	expect(layout.approximations).toContain(
		'Independent inner-cell continuation around nested note references remains approximate when an enclosing bordered row splits.',
	);
	const saved = await loaded.save(loaded.model);
	expect(await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')).toBe(
		await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string'),
	);
	expect((await loadDocx(saved)).model).toEqual(loaded.model);
});
