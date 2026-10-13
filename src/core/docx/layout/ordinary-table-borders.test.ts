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
const innerNative = JSON.parse(
	readFileSync(
		new URL('../__fixtures__/independent-inner-notes/native-measurements.json', import.meta.url),
		'utf8',
	),
) as { cases: Evidence[] };
const multipleNative = JSON.parse(
	readFileSync(
		new URL('../__fixtures__/multiple-inner-notes/native-measurements.json', import.meta.url),
		'utf8',
	),
) as { cases: Evidence[] };
it.each([...native.cases, ...innerNative.cases, ...multipleNative.cases])(
	'matches native ordinary border row bands and source preservation: $name',
	async (sample) => {
		const bytes = new Uint8Array(
			readFileSync(
				new URL(
					'../__fixtures__/' +
						(sample.name.startsWith('multi-note-')
							? 'multiple-inner-notes/'
							: sample.name.startsWith('inner-note-')
								? 'independent-inner-notes/'
								: 'ordinary-table-borders/') +
						sample.name +
						'.docx',
					import.meta.url,
				),
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
			const sourceMarkers = table.rows
				.flatMap((row) => row.cells.flat())
				.flatMap((box) => box.lines)
				.flatMap(
					(line) =>
						line.fragments
							.map((fragment) => fragment.text)
							.join('')
							.match(/T\dR\d+C\d+L\d+/g) ?? [],
				)
				.sort();
			const nativeMarkers = sample.measurements[index]!.lines.flatMap(
				(line) => line.text.match(/T\dR\d+C\d+L\d+/g) ?? [],
			).sort();
			expect(sourceMarkers).toEqual(nativeMarkers);
			const noteMarkers = (page.footnotes ?? [])
				.flatMap((note) => note.paragraphs)
				.flatMap((box) => box.lines)
				.flatMap(
					(line) =>
						line.fragments
							.map((fragment) => fragment.text)
							.join('')
							.match(/N\d+L\d+/g) ?? [],
				)
				.sort();
			expect(noteMarkers).toEqual(
				sample.measurements[index]!.lines.flatMap(
					(line) => line.text.match(/N\d+L\d+/g) ?? [],
				).sort(),
			);

			if (sample.name.startsWith('multi-note-')) {
				const notes = page.footnotes ?? [];
				const areaTop =
					page.heightPx -
					page.marginBottomPx -
					(page.footnoteHeightPx ?? notes.reduce((height, note) => height + note.heightPx, 0));
				for (const note of notes)
					for (const box of note.paragraphs)
						for (const line of box.lines) {
							const marker = line.fragments
								.map((fragment) => fragment.text)
								.join('')
								.match(/N\d+L\d+/)?.[0];
							if (!marker) continue;
							const nativeLine = sample.measurements[index]!.lines.find((line) =>
								line.text.includes(marker),
							)!;
							expect(
								Math.abs(
									(areaTop + note.yPx + box.yPx + line.yPx) * 0.75 + 9.6 - nativeLine.baselinePt,
								),
								marker,
							).toBeLessThan(0.2);
						}
			}

			for (const row of table.rows)
				for (const [column, boxes] of row.cells.entries()) {
					for (const box of boxes) {
						for (const line of box.lines) {
							const marker = line.fragments
								.map((fragment) => fragment.text)
								.join('')
								.match(/T\dR\d+C\d+L\d+/)?.[0];
							if (!marker) continue;
							const nativeLine = sample.measurements[index]!.lines.find((line) =>
								line.text.includes(marker),
							)!;
							expect(
								Math.abs(
									(row.yPx + box.yPx + line.yPx) * 0.75 - (nativeLine.baselinePt - 9.6 - 40),
								),
								marker,
							).toBeLessThan(0.2);
						}

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
		for (const part of ['word/document.xml', 'word/footnotes.xml']) {
			const original = before.file(part);
			if (original)
				expect(await after.file(part)!.async('string')).toBe(await original.async('string'));
		}

		expect((await loadDocx(saved)).model).toEqual(loaded.model);
	},
);

it('cuts native parallel inner note cells independently without changing source', async () => {
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
	expect(layout.approximations).not.toContain(
		'Independent inner-cell continuation around nested note references remains approximate when an enclosing bordered row splits.',
	);
	const counts = layout.pages.map((page) => {
		const table = page.columns[0]!.blocks.find((block) => block.kind === 'table');
		if (!table || table.kind !== 'table') throw new Error('Expected table');
		return [0, 1].map(
			(index) =>
				table.rows
					.flatMap((row) => row.cells.flat())
					.flatMap((box) => box.lines)
					.filter((line) =>
						line.fragments
							.map((fragment) => fragment.text)
							.join('')
							.includes('T1R0C' + index + 'L'),
					).length,
		);
	});
	expect(counts).toEqual([
		[15, 17],
		[5, 3],
	]);
	const saved = await loaded.save(loaded.model);
	expect(await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string')).toBe(
		await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string'),
	);
	expect((await loadDocx(saved)).model).toEqual(loaded.model);
});
