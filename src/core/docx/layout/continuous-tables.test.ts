import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index.js';
import { createFakeMeasurer } from './measure.js';
import { layoutDocumentModel } from './layout.js';
import { layoutSections } from './page-flow.js';

const fixture = (name: string) => new URL(`./fixtures/continuous-tables/${name}`, import.meta.url);
describe('native Word tables before continuous section breaks', () => {
	it('rejects trial capacities that overflow a kept row in an earlier column', () => {
		const page = {
			widthPx: 200,
			heightPx: 100,
			marginTopPx: 0,
			marginBottomPx: 0,
			marginLeftPx: 0,
			marginRightPx: 0,
		};
		const result = layoutSections(
			{
				sections: [
					{
						page,
						columns: { count: 2, gapPx: 0 },
						blocks: [
							{
								kind: 'table',
								id: 'table',
								rows: [60, 20].map((heightPx) => ({
									heightPx,
									heightRule: 'exact',
									cantSplit: true,
									cells: [{ paragraphs: [] }],
								})),
							},
						],
					},
					{
						page,
						break: 'continuous',
						blocks: [{ kind: 'paragraph', id: 'after', runs: [{ text: 'After' }] }],
					},
				],
			},
			createFakeMeasurer(),
		);
		expect(result.pages).toHaveLength(1);
		expect(result.pages[0]!.columns[2]!.blocks[0]!.yPx).toBe(60);
		expect(result.approximations).toEqual([]);
	});
	it.each(['table-even', 'table-odd', 'table-overflow'])(
		'matches each row and following paragraph in %s',
		async (name) => {
			const evidence = JSON.parse(await readFile(fixture('evidence.json'), 'utf8')) as {
				cases: {
					name: string;
					pages: number;
					positions: { text: string; page: number; xPt: number; yPt: number }[];
				}[];
			};
			const reference = evidence.cases.find((entry) => entry.name === name)!;
			const { model } = await loadDocx(new Uint8Array(await readFile(fixture(`${name}.docx`))));
			const result = layoutDocumentModel(model, createFakeMeasurer());
			expect(result.pages).toHaveLength(reference.pages);
			const positions = result.pages.flatMap((sheet) =>
				sheet.columns.flatMap((column) =>
					column.blocks.flatMap((block) => {
						const lines =
							block.kind === 'paragraph'
								? block.lines.map((line) => ({ line, x: 0, y: block.yPx }))
								: block.rows.flatMap((row) =>
										row.cells.flatMap((cell, index) =>
											cell.flatMap((paragraph) =>
												paragraph.lines.map((line) => ({
													line,
													x:
														(block.xPx ?? 0) +
														(row.geometry?.[index]?.xPx ?? 0) +
														(row.geometry?.[index]?.paddingLeftPx ?? 0),
													y: block.yPx + row.yPx,
												})),
											),
										),
									);
						return lines.map(({ line, x, y }) => ({
							text: line.fragments.map((fragment) => fragment.text).join(''),
							page: sheet.index + 1,
							xPt: (sheet.marginLeftPx + column.xPx + x) * 0.75,
							yPt: (sheet.marginTopPx + y + line.yPx) * 0.75,
						}));
					}),
				),
			);
			for (const position of reference.positions.filter((entry) => entry.text.trim())) {
				const actual = positions.filter((entry) => entry.text === position.text);
				expect(actual).toEqual([
					{ text: position.text, page: position.page, xPt: position.xPt, yPt: position.yPt },
				]);
			}
			expect(result.approximations.some((note) => note.includes('balancing those layouts'))).toBe(
				false,
			);
		},
	);
});
