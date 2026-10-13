import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
it.each(['short', 'two-rows', 'centered', 'split'])(
	'preserves and projects native nested frames: %s',
	async (name) => {
		const bytes = new Uint8Array(
			readFileSync(
				new URL(
					'../__fixtures__/nested-table-frames/nested-frame-' + name + '.docx',
					import.meta.url,
				),
			),
		);
		const loaded = await loadDocx(bytes);
		const layout = layoutDocumentModel(loaded.model, {
			widthOf: (text) => text.length * 5,
			lineHeightOf: () => 16,
		});
		const frames = layout.pages.flatMap((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) =>
					block.kind === 'table' ? block.rows.flatMap((row) => row.cellTables?.flat() ?? []) : [],
				),
			),
		);
		expect(frames.length).toBeGreaterThan(0);
		for (const frame of frames) {
			expect(frame.widthPx).toBeGreaterThan(0);
			expect(frame.heightPx).toBeGreaterThan(0);
			for (const row of frame.rows) {
				expect(row.yPx + row.heightPx).toBeLessThanOrEqual(frame.heightPx + 0.001);
				expect(['#FFFF00', '#00FF00']).toContain(row.geometry?.[0]?.shading);
				expect(row.geometry?.[0]?.borders?.left?.color).toBe('#FF0000');
			}
		}
		if (name === 'centered') expect(frames[0]!.xPx).toBeCloseTo((20 * 4) / 3);
		if (name === 'split')
			expect(
				layout.pages.filter((page) =>
					page.columns.some((column) =>
						column.blocks.some(
							(block) =>
								block.kind === 'table' &&
								block.rows.some((row) => row.cellTables?.some((tables) => tables.length)),
						),
					),
				).length,
			).toBe(2);
		const saved = await loaded.save(loaded.model),
			original = await JSZip.loadAsync(bytes),
			reopened = await JSZip.loadAsync(saved);
		expect(await reopened.file('word/document.xml')!.async('string')).toBe(
			await original.file('word/document.xml')!.async('string'),
		);
		expect((await loadDocx(saved)).model).toEqual(loaded.model);
	},
);

it('aligns projected nested ink with the shared fixed-height cell geometry', async () => {
	const positions: number[] = [];
	for (const alignment of ['top', 'center', 'bottom']) {
		const loaded = await loadDocx(
			new Uint8Array(
				readFileSync(
					new URL(
						'../__fixtures__/nested-table-frames/nested-frame-align-' + alignment + '.docx',
						import.meta.url,
					),
				),
			),
		);
		const layout = layoutDocumentModel(loaded.model, {
			widthOf: (text) => text.length * 5,
			lineHeightOf: () => 16,
		});
		const table = layout.pages[0]!.columns[0]!.blocks.find((block) => block.kind === 'table');
		if (!table || table.kind !== 'table') throw new Error('Expected table');
		const row = table.rows[0]!,
			frame = row.cellTables![0]![0]!;
		const paragraph = row.cells[0]!.find((box) => box.readOnly)!;
		expect(frame.rows[0]!.heightPx).toBeCloseTo(40);
		expect(frame.rows[0]!.geometry![0]!.verticalAlign ?? 'top').toBe(alignment);
		positions.push(paragraph.yPx - frame.yPx);
	}
	expect(positions[1]! - positions[0]!).toBeCloseTo(12);
	expect(positions[2]! - positions[0]!).toBeCloseTo(24);
	const native = JSON.parse(
		readFileSync(
			new URL('../__fixtures__/nested-table-frames/native-measurements.json', import.meta.url),
			'utf8',
		),
	) as {
		cases: {
			name: string;
			measurements: {
				borders: { rectPt: number[] }[];
				lines: { text: string; baselinePt: number }[];
			}[];
		}[];
	};
	const baselines = ['top', 'center', 'bottom'].map(
		(alignment) =>
			native.cases
				.find((c) => c.name === 'nested-frame-align-' + alignment)!
				.measurements[0]!.lines.find((line) => line.text.startsWith('T1R0C0L00'))!.baselinePt,
	);
	expect((positions[1]! - positions[0]!) * 0.75).toBeCloseTo(9);
	expect(baselines[1]! - baselines[0]!).toBeCloseTo(8.52);
	expect(baselines[2]! - baselines[0]!).toBeCloseTo(16.92);
	const split = native.cases.find((c) => c.name === 'nested-frame-split')!;
	for (const page of split.measurements) {
		const horizontal = page.borders.filter((border) => border.rectPt[2]! - border.rectPt[0]! > 10);
		expect(horizontal).toHaveLength(4);
		expect(horizontal.filter((border) => border.rectPt[1]! < 41)).toHaveLength(2);
		expect(horizontal.filter((border) => border.rectPt[1]! > 41)).toHaveLength(2);
	}
});
