import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
it.each([
	'short',
	'two-rows',
	'centered',
	'split',
	'align-top',
	'align-center',
	'align-bottom',
	'border-half',
	'border-two',
	'border-three-margin',
	'margin',
	'unequal-split',
	'mixed',
	'mixed-margin',
	'mixed-centered',
	'mixed-two-rows',
	'mixed-split',
	'mixed-outer-border',
	'mixed-deep',
])('preserves and projects native nested frames: %s', async (name) => {
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
	if (name === 'centered') expect(frames[0]!.xPx).toBeCloseTo((19.5 * 4) / 3);
	if (name.endsWith('split'))
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
});

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
		expect(frame.rows[0]!.heightPx).toBeCloseTo(124 / 3);
		expect(frame.rows[0]!.geometry![0]!.verticalAlign ?? 'top').toBe(alignment);
		positions.push(paragraph.yPx - frame.yPx);
	}
	expect(positions[1]! - positions[0]!).toBeCloseTo(34 / 3);
	expect(positions[2]! - positions[0]!).toBeCloseTo(68 / 3);
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
	expect((positions[1]! - positions[0]!) * 0.75).toBeCloseTo(8.5);
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

it.each([
	'short',
	'two-rows',
	'centered',
	'border-half',
	'border-two',
	'border-three-margin',
	'margin',
	'unequal-split',
	'mixed',
	'mixed-margin',
	'mixed-centered',
	'mixed-two-rows',
	'mixed-split',
	'align-top',
	'align-center',
	'align-bottom',
])('reserves native border and margin geometry: %s', async (name) => {
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
		ink = row.cells[0]!.find((box) => box.readOnly)!;
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
				lines: { text: string; leftPt: number; baselinePt: number }[];
			}[];
		}[];
	};
	const sample = native.cases.find((c) => c.name === 'nested-frame-' + name)!.measurements[0]!;
	const nativeInk = sample.lines.find((line) => line.text.startsWith('T1R0C0L00'))!;
	const top = Math.min(...sample.borders.map((b) => b.rectPt[1]!));
	const left = Math.min(...sample.borders.map((b) => b.rectPt[0]!));
	const right = Math.max(...sample.borders.map((b) => b.rectPt[2]!));
	const bottom = Math.max(...sample.borders.map((b) => b.rectPt[3]!));
	expect(Math.abs(frame.heightPx * 0.75 - (bottom - top))).toBeLessThan(0.2);
	expect(Math.abs(frame.widthPx! * 0.75 - (right - left))).toBeLessThan(0.2);
	expect(Math.abs((ink.yPx - frame.yPx) * 0.75 - (nativeInk.baselinePt - 9.6 - top))).toBeLessThan(
		0.15,
	);
	expect(Math.abs((ink.xPx! - frame.xPx!) * 0.75 - (nativeInk.leftPt - left))).toBeLessThan(0.15);
});

it('matches native bordered continuation bands and reference-line pagination', async () => {
	const loaded = await loadDocx(
		new Uint8Array(
			readFileSync(
				new URL('../__fixtures__/nested-table-frames/nested-frame-split.docx', import.meta.url),
			),
		),
	);
	const layout = layoutDocumentModel(loaded.model, {
		widthOf: (text) => text.length * 5,
		lineHeightOf: () => 16,
	});
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
	const pages = native.cases.find((c) => c.name === 'nested-frame-split')!.measurements;
	expect(layout.pages).toHaveLength(pages.length);
	for (const [index, page] of layout.pages.entries()) {
		const table = page.columns[0]!.blocks.find((block) => block.kind === 'table');
		if (!table || table.kind !== 'table') throw new Error('Expected table');
		const row = table.rows[0]!,
			frame = row.cellTables![0]![0]!,
			ink = row.cells[0]!.find((box) => box.readOnly)!;
		const evidence = pages[index]!;
		const top = Math.min(...evidence.borders.map((b) => b.rectPt[1]!)),
			bottom = Math.max(...evidence.borders.map((b) => b.rectPt[3]!));
		expect(Math.abs(frame.heightPx * 0.75 - (bottom - top))).toBeLessThan(0.2);
		expect(ink.lines).toHaveLength(index === 0 ? 16 : 4);
		expect((ink.yPx - frame.yPx) * 0.75).toBeCloseTo(1);
		expect(ink.yPx + ink.heightPx + (ink.tableFrameInsets?.bottom ?? 0)).toBeLessThanOrEqual(
			row.heightPx + 0.001,
		);
	}
});

it('keeps unequal native nested continuations and their note inside the row band', async () => {
	const loaded = await loadDocx(
		new Uint8Array(
			readFileSync(
				new URL(
					'../__fixtures__/nested-table-frames/nested-frame-unequal-split.docx',
					import.meta.url,
				),
			),
		),
	);
	const layout = layoutDocumentModel(loaded.model, {
		widthOf: (text) => text.length * 5,
		lineHeightOf: () => 16,
	});
	expect(layout.pages).toHaveLength(2);
	const table = layout.pages[1]!.columns[0]!.blocks.find((block) => block.kind === 'table');
	if (!table || table.kind !== 'table') throw new Error('Expected table');
	const row = table.rows[0]!;
	const paragraphs = row.cells[0]!.filter((box) => box.readOnly);
	expect(paragraphs.map((box) => box.lines.length)).toEqual([4, 2]);
	const bottom = Math.max(
		...paragraphs.map((box) => box.yPx + box.heightPx + (box.tableFrameInsets?.bottom ?? 0)),
	);
	expect(row.heightPx).toBeGreaterThanOrEqual(bottom);
	expect(row.geometry![0]!.contentHeightPx).toBeGreaterThanOrEqual(bottom);
	expect(layout.pages[0]!.footnotes).toHaveLength(1);
	expect(layout.pages[1]!.footnotes).toHaveLength(1);
});
