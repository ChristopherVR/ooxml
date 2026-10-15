import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
interface Evidence {
	name: string;
	measurements: {
		fills: { color: string; rectPt: number[] }[];
		borders: { rectPt: number[] }[];
		lines: { text: string; leftPt: number; baselinePt: number }[];
	}[];
}
const native = JSON.parse(
	readFileSync(
		new URL('../__fixtures__/nested-table-frames/native-measurements.json', import.meta.url),
		'utf8',
	),
) as { cases: Evidence[] };
it.each(['mixed', 'mixed-margin', 'mixed-centered', 'mixed-two-rows', 'mixed-split'])(
	'matches native asymmetric cell border centers and fill widths: %s',
	async (name) => {
		const loaded = await loadDocx(
			new Uint8Array(
				readFileSync(
					new URL(
						'../__fixtures__/nested-table-frames/nested-frame-' + name + '.docx',
						import.meta.url,
					),
				),
			),
		);
		const layout = layoutDocumentModel(loaded.model, {
			widthOf: (text) => text.length * 5,
			lineHeightOf: () => 16,
		});
		expect(layout.approximations.join(' ')).not.toContain('Nested table border occupancy');
		const sample = native.cases.find((c) => c.name === 'nested-frame-' + name)!;
		expect(layout.pages).toHaveLength(sample.measurements.length);
		for (const [index, page] of layout.pages.entries()) {
			const table = page.columns[0]!.blocks.find((block) => block.kind === 'table');
			if (!table || table.kind !== 'table') throw new Error('Expected table');
			const frame = table.rows[0]!.cellTables![0]![0]!,
				row = frame.rows[0]!,
				evidence = sample.measurements[index]!;
			const nativeLeft = Math.min(...evidence.borders.map((b) => b.rectPt[0]!));
			for (const [cellIndex, color] of ['#FFFF00', '#00FFFF'].entries()) {
				const fill = evidence.fills
					.filter((f) => f.color === color)
					.sort(
						(a, b) =>
							(b.rectPt[2]! - b.rectPt[0]!) * (b.rectPt[3]! - b.rectPt[1]!) -
							(a.rectPt[2]! - a.rectPt[0]!) * (a.rectPt[3]! - a.rectPt[1]!),
					)[0]!;
				const geometry = row.geometry![cellIndex]!;
				const left = geometry.xPx + (geometry.borders?.left?.widthPx ?? 0);
				const right =
					geometry.xPx +
					geometry.widthPx -
					(cellIndex === 1 ? (geometry.borders?.right?.widthPx ?? 0) : 0);
				expect(Math.abs(left * 0.75 - (fill.rectPt[0]! - nativeLeft))).toBeLessThan(0.2);
				expect(Math.abs(right * 0.75 - (fill.rectPt[2]! - nativeLeft))).toBeLessThan(0.2);
			}
			const ink = table.rows[0]!.cells[0]!.filter((box) => box.nested);
			expect(ink[1]!.xPx! - ink[0]!.xPx!).toBeCloseTo(
				((name === 'mixed-centered' ? 60 : 80) * 4) / 3 + (name === 'mixed-margin' ? 0 : 2 / 3),
				1,
			);
			if (name === 'mixed-split')
				expect(ink.map((box) => box.lines.length)).toEqual(index === 0 ? [16, 16] : [4, 4]);
		}
	},
);

it.each(['mixed-outer-border', 'mixed-deep'])(
	'keeps native inner-frame relative ink coordinates inside bordered enclosing cells: %s',
	async (name) => {
		const loaded = await loadDocx(
			new Uint8Array(
				readFileSync(
					new URL(
						'../__fixtures__/nested-table-frames/nested-frame-' + name + '.docx',
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
			frame = row.cellTables![0]![0]!,
			evidence = native.cases.find((c) => c.name === 'nested-frame-' + name)!.measurements[0]!;
		const left = Math.min(...evidence.borders.map((border) => border.rectPt[0]!)),
			top = Math.min(...evidence.borders.map((border) => border.rectPt[1]!));
		for (const text of name === 'mixed-deep' ? ['T1R0C0L00', 'D0'] : ['T1R0C0L00']) {
			const ink = row.cells[0]!.find((box) =>
				box.lines.some((line) => line.fragments.some((fragment) => fragment.text.includes(text))),
			)!;
			const source = evidence.lines.find((line) => line.text.includes(text))!;
			expect(Math.abs((ink.xPx! - frame.xPx!) * 0.75 - (source.leftPt - left))).toBeLessThan(0.15);
			expect(Math.abs((ink.yPx - frame.yPx) * 0.75 - (source.baselinePt - 9.6 - top))).toBeLessThan(
				0.15,
			);
		}
	},
);
