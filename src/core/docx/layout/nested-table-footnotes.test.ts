import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadDocx, numberNotesInOrder } from '../index';
import JSZip from 'jszip';
import { layoutDocumentModel } from './layout';
const fixture = (name: string) =>
	new Uint8Array(
		readFileSync(new URL('../__fixtures__/table-footnotes/' + name + '.docx', import.meta.url)),
	);
describe('nested table source stories', () => {
	it('projects SDK-inserted direct paragraphs without dropping their notes and safely refuses unsupported source restructuring', async () => {
		const bytes = fixture('nested-table-reference'),
			loaded = await loadDocx(bytes),
			original = structuredClone(loaded.model);
		const table = loaded.model.blocks.find((block) => block.type === 'table')!;
		if (table.type !== 'table') throw new Error('Expected table');
		table.rows[0]![0]!.paragraphs.push({
			type: 'paragraph',
			id: 'added',
			runs: [{ text: 'Added tail' }, { text: '', noteReference: { kind: 'footnote', id: '2' } }],
		});
		loaded.model.footnotes!.push({
			id: '2',
			blocks: [{ type: 'paragraph', id: 'new-note', runs: [{ text: 'Added source' }] }],
		});
		expect([...numberNotesInOrder(loaded.model.blocks, 'footnote')]).toEqual([
			['1', 1],
			['2', 2],
		]);
		const result = layoutDocumentModel(loaded.model, {
			widthOf: (t) => t.length * 5,
			lineHeightOf: () => 16,
		});
		const paragraphs = result.pages.flatMap((p) =>
			p.columns.flatMap((c) =>
				c.blocks.flatMap((b) => (b.kind === 'table' ? b.rows.flatMap((r) => r.cells.flat()) : [])),
			),
		);
		expect(paragraphs.some((p) => p.blockId === 'added')).toBe(true);
		expect(result.pages.flatMap((p) => p.footnotes ?? []).map((note) => note.id)).toEqual([
			'1',
			'2',
		]);
		await expect(loaded.save(loaded.model)).rejects.toThrow('Cannot change the structure');
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(await loaded.save(original));
		for (const path of Object.keys(before.files).filter((path) => !before.files[path]!.dir))
			expect(await after.file(path)!.async('uint8array')).toEqual(
				await before.file(path)!.async('uint8array'),
			);
	});
	it.each(['nested-table-reference', 'nested-split-early', 'nested-split-late'])(
		'%s retains source ordering, marker labels and read-only local bounds through an outer edit',
		async (name) => {
			const loaded = await loadDocx(fixture(name)),
				table = loaded.model.blocks.find((b) => b.type === 'table')!;
			if (table.type !== 'table') throw new Error('Expected table');
			const cell = table.rows[0]![0]!;
			expect(cell.sourceBlocks?.map((b) => b.type)).toEqual(['table', 'paragraph']);
			const nested = cell.sourceBlocks![0]!;
			if (nested.type !== 'table') throw new Error('Expected nested table');
			expect(nested.rows[0]![0]!.paragraphs[0]!.runs.some((r) => r.noteReference?.id === '1')).toBe(
				true,
			);
			expect([...numberNotesInOrder(loaded.model.blocks, 'footnote')]).toEqual([['1', 1]]);
			cell.paragraphs[0] = { ...cell.paragraphs[0]!, runs: [{ text: 'Edited outer' }] };
			const immediate = layoutDocumentModel(loaded.model, {
				widthOf: (t) => t.length * 5,
				lineHeightOf: () => 16,
			});
			expect(
				immediate.pages
					.flatMap((p) =>
						p.columns.flatMap((c) =>
							c.blocks.flatMap((b) =>
								b.kind === 'table'
									? b.rows.flatMap((r) =>
											r.cells.flatMap((cell) =>
												cell.flatMap((p) => p.lines.flatMap((l) => l.fragments.map((f) => f.text))),
											),
										)
									: [],
							),
						),
					)
					.join(''),
			).toContain('Edited outer');
			const saved = await loaded.save(loaded.model);
			const originalZip = await JSZip.loadAsync(fixture(name)),
				savedZip = await JSZip.loadAsync(saved);
			for (const path of Object.keys(originalZip.files).filter(
				(path) => !originalZip.files[path]!.dir && path !== 'word/document.xml',
			))
				expect(await savedZip.file(path)!.async('uint8array')).toEqual(
					await originalZip.file(path)!.async('uint8array'),
				);
			const reopened = await loadDocx(saved);
			const reopenedTable = reopened.model.blocks.find((block) => block.type === 'table')!;
			if (reopenedTable.type !== 'table') throw new Error('Expected reopened table');
			expect(reopenedTable.rows[0]![0]!.sourceBlocks![0]).toEqual(nested);
			const result = layoutDocumentModel(reopened.model, {
				widthOf: (t) => t.length * 5,
				lineHeightOf: () => 16,
			});
			const projected = result.pages.flatMap((p) =>
				p.columns.flatMap((c) =>
					c.blocks.flatMap((b) =>
						b.kind === 'table'
							? b.rows.flatMap((r) => r.cells.flatMap((cell) => cell.filter((p) => p.readOnly)))
							: [],
					),
				),
			);
			expect(projected.some((p) => p.xPx === 0)).toBe(true);
			expect(projected.some((p) => p.xPx === 106.66666666666667)).toBe(true);
			expect(projected.every((p) => p.widthPx === 106.66666666666667)).toBe(true);
			expect(
				result.pages
					.flatMap((p) => p.footnotes ?? [])
					.flatMap((n) =>
						n.paragraphs.flatMap((p) =>
							p.lines.flatMap(
								(l) =>
									l.fragments
										.map((f) => f.text)
										.join('')
										.match(/N0L\d+/g) ?? [],
							),
						),
					),
			).toEqual(['N0L00', 'N0L01', 'N0L02']);
		},
	);
});
