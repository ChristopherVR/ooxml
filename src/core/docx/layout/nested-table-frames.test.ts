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
