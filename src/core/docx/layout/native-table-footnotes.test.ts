import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { adaptDocumentModel } from './adapter';
import { layoutDocumentModel } from './layout';
import type { LayoutParagraph } from './input';
interface NativeCase {
	name: string;
	references: { id: string; page: number }[];
	pages: {
		columns: { leftPt: number; lines: { id: string }[] }[];
		separatorCaptions: string[];
		cellLines: { id: string; leftPt: number }[];
	}[];
}
const fixture = (name: string) =>
	new Uint8Array(
		readFileSync(new URL('../__fixtures__/table-footnotes/' + name + '.docx', import.meta.url)),
	);
const corpus = JSON.parse(
	readFileSync(
		new URL('../__fixtures__/table-footnotes/native-measurements.json', import.meta.url),
		'utf8',
	),
) as { cases: NativeCase[] };
const measurer = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 16 };
describe('native Word table footnote pagination', () => {
	it.each(corpus.cases)(
		'$name preserves native cell boundaries, reference pages and every note line',
		async (native) => {
			const loaded = await loadDocx(fixture(native.name));
			const before = JSON.stringify(loaded.model);
			const input = adaptDocumentModel(loaded.model),
				paragraphs = new Map<string, LayoutParagraph>();
			for (const section of input.sections)
				for (const block of section.blocks)
					if (block.kind === 'table')
						for (const row of block.rows)
							for (const cell of row.cells)
								for (const paragraph of cell.paragraphs) paragraphs.set(paragraph.id, paragraph);
			const ids = new Map(
				(loaded.model.footnotes ?? []).map((note) => [
					note.id,
					note.blocks
						.flatMap((block) =>
							block.type === 'paragraph' ? block.runs.map((run) => run.text) : [],
						)
						.join('')
						.match(/N\d+/)?.[0],
				]),
			);
			const references = new Map<string, number>();
			const result = layoutDocumentModel(loaded.model, measurer);
			expect(result.pages).toHaveLength(native.pages.length);
			for (const [index, page] of result.pages.entries()) {
				const expected = native.pages[index]!;
				const noteIds = (page.footnotes ?? []).flatMap((note) =>
					note.paragraphs.flatMap((p) =>
						p.lines.flatMap(
							(line) =>
								line.fragments
									.map((fragment) => fragment.text)
									.join('')
									.match(/N\dL\d+/g) ?? [],
						),
					),
				);
				expect(noteIds).toEqual(
					expected.columns.flatMap((column) => column.lines.map((line) => line.id)),
				);
				const captions =
					page.footnoteColumnAreas?.flatMap((area) => (area.separator ? [area.separator] : [])) ??
					(page.footnoteSeparator ? [page.footnoteSeparator] : []);
				expect(
					captions.map((story) =>
						story.paragraphs
							.flatMap((p) =>
								p.lines.flatMap((line) => line.fragments.map((fragment) => fragment.text)),
							)
							.join('')
							.trim(),
					),
				).toEqual(expected.separatorCaptions);
				const cells: { id: string; leftPt: number }[] = [];
				for (const column of page.columns)
					for (const block of column.blocks)
						if (block.kind === 'table')
							for (const row of block.rows)
								for (const [cellIndex, cell] of row.cells.entries())
									for (const paragraph of cell) {
										const source = paragraphs.get(paragraph.blockId)!;
										for (const line of paragraph.lines) {
											for (const fragment of line.fragments) {
												const noteId = source.runs[fragment.runIndex]?.footnoteReferenceId;
												if (noteId && !row.repeated && !references.has(noteId))
													references.set(noteId, index + 1);
											}
											if (paragraph.yPx + line.yPx + line.heightPx > row.heightPx + 1e-6) continue;
											const text = line.fragments.map((fragment) => fragment.text).join('');
											for (const id of text.match(/T\dR\d+C\d+L\d+/g) ?? [])
												cells.push({
													id,
													leftPt:
														(page.marginLeftPx +
															column.xPx +
															(block.xPx ?? 0) +
															row.geometry![cellIndex]!.xPx +
															row.geometry![cellIndex]!.paddingLeftPx) *
														0.75,
												});
										}
									}
				expect(cells.map((cell) => cell.id)).toEqual(expected.cellLines.map((cell) => cell.id));
				for (const [cellIndex, cell] of cells.entries())
					expect(Math.abs(cell.leftPt - expected.cellLines[cellIndex]!.leftPt)).toBeLessThan(0.2);
			}
			expect([...references].map(([id, page]) => ({ id: ids.get(id), page }))).toEqual(
				native.references,
			);
			expect(JSON.stringify(loaded.model)).toBe(before);
		},
	);
});
