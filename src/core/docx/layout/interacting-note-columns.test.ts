import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
const evidence = JSON.parse(
	readFileSync(
		new URL('../__fixtures__/interacting-note-columns/native-measurements.json', import.meta.url),
		'utf8',
	),
) as {
	cases: {
		name: string;
		pages: number;
		measurements: {
			lines: { text: string; baselinePt: number; leftPt: number }[];
			borders: { role: string; rectPt: number[] }[];
		}[];
	}[];
};
it.each(evidence.cases)(
	'matches native interacting note column geometry and source: $name',
	async (sample) => {
		const bytes = new Uint8Array(
			readFileSync(
				new URL(
					'../__fixtures__/interacting-note-columns/' + sample.name + '.docx',
					import.meta.url,
				),
			),
		);
		const loaded = await loadDocx(bytes),
			layout = layoutDocumentModel(loaded.model, {
				widthOf: (text) => text.length * 5,
				lineHeightOf: () => 16,
			});
		const summaries = layout.pages.map((page) => ({
			body: page.columns
				.flatMap((column) => column.blocks)
				.filter((block) => block.kind === 'table')
				.flatMap((table) => table.rows)
				.flatMap((row) => row.cells.flat())
				.flatMap((box) => box.lines)
				.flatMap(
					(line) =>
						line.fragments
							.map((fragment) => fragment.text)
							.join('')
							.match(/T\dR\d+C\d+L\d{2}/g) ?? [],
				),
			notes: (page.footnotes ?? [])
				.flatMap((note) => note.paragraphs)
				.flatMap((box) => box.lines)
				.flatMap(
					(line) =>
						line.fragments
							.map((fragment) => fragment.text)
							.join('')
							.match(/N\d+L\d{2}/g) ?? [],
				),
		}));
		expect(layout.pages).toHaveLength(sample.pages);
		for (const [index, page] of layout.pages.entries()) {
			expect(summaries[index]!.body.sort()).toEqual(
				sample.measurements[index]!.lines.flatMap(
					(line) => line.text.match(/T\dR\d+C\d+L\d{2}/g) ?? [],
				).sort(),
			);
			expect(summaries[index]!.notes.sort()).toEqual(
				sample.measurements[index]!.lines.flatMap(
					(line) => line.text.match(/N\d+L\d{2}/g) ?? [],
				).sort(),
			);
			const tables = page.columns
				.flatMap((column) => column.blocks)
				.filter((block) => block.kind === 'table');
			const borders = sample.measurements[index]!.borders.filter(
				(border) => border.role === 'outer',
			);
			if (!borders.length) expect(tables).toHaveLength(0);
			else {
				expect(tables).toHaveLength(1);
				const table = tables[0]!;
				const top = Math.min(...borders.map((border) => border.rectPt[1]!)),
					bottom = Math.max(...borders.map((border) => border.rectPt[3]!));
				expect(Math.abs(table.heightPx * 0.75 - (bottom - top))).toBeLessThan(0.2);
				for (const row of table.rows)
					for (const [column, boxes] of row.cells.entries())
						for (const box of boxes)
							for (const line of box.lines) {
								const marker = line.fragments
									.map((fragment) => fragment.text)
									.join('')
									.match(/T\dR\d+C\d+L\d{2}|Outer cell/)?.[0];
								if (!marker) continue;
								const native = sample.measurements[index]!.lines.find((line) =>
									line.text.includes(marker),
								);
								if (!native) throw new Error('Missing native body ' + marker);
								const geometry = row.geometry![column]!;
								const x = (table.xPx ?? 0) + geometry.xPx + geometry.paddingLeftPx + (box.xPx ?? 0);
								expect(Math.abs(x * 0.75 + 40 - native.leftPt), marker).toBeLessThan(0.2);
								expect(
									Math.abs(
										(table.yPx + row.yPx + box.yPx + line.yPx) * 0.75 +
											40 +
											9.6 -
											native.baselinePt,
									),
									marker,
								).toBeLessThan(0.2);
							}
			}
			const notes = page.footnotes ?? [];
			const areaTop =
				page.heightPx -
				page.marginBottomPx -
				(page.footnoteHeightPx ?? notes.reduce((height, note) => height + note.heightPx, 0));
			for (const note of notes) {
				const first = note.paragraphs
					.flatMap((box) => box.lines)
					.map((line) => line.fragments.map((fragment) => fragment.text).join(''))
					.join('')
					.match(/N(\d+)L(\d{2})/);
				const sourceFirst = sample.measurements[index]!.lines.map((line) =>
					line.text.match(/N(\d+)L(\d{2})/),
				)
					.filter((match) => match?.[1] === first?.[1])
					.map((match) => Number(match![2]));
				expect(!!note.continued).toBe(Math.min(...sourceFirst) > 0);
				for (const box of note.paragraphs)
					for (const line of box.lines) {
						const marker = line.fragments
							.map((fragment) => fragment.text)
							.join('')
							.match(/N\d+L\d{2}/)?.[0];
						if (!marker) continue;
						const native = sample.measurements[index]!.lines.find((line) =>
							line.text.includes(marker),
						)!;
						expect(
							Math.abs(
								(page.marginLeftPx + (note.xPx ?? 0) + (box.xPx ?? 0)) * 0.75 - native.leftPt,
							),
							marker,
						).toBeLessThan(0.2);
						expect(
							Math.abs((areaTop + note.yPx + box.yPx + line.yPx) * 0.75 + 9.6 - native.baselinePt),
							marker,
						).toBeLessThan(0.2);
					}
			}
		}
		const saved = await loaded.save(loaded.model),
			before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(saved);
		for (const part of ['word/document.xml', 'word/footnotes.xml'])
			expect(await after.file(part)!.async('string')).toBe(
				await before.file(part)!.async('string'),
			);
		expect((await loadDocx(saved)).model).toEqual(loaded.model);
	},
);
