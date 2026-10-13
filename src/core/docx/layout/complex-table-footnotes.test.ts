import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { loadDocx } from '../index';
import JSZip from 'jszip';
import { layoutDocumentModel } from './layout';
it.each(['nested-table-reference', 'vertical-merge-reference'])(
	'%s preserves the native source and reports its measured unsupported layout',
	async (name) => {
		const loaded = await loadDocx(
			new Uint8Array(
				readFileSync(new URL('../__fixtures__/table-footnotes/' + name + '.docx', import.meta.url)),
			),
		);
		const result = layoutDocumentModel(loaded.model, {
			widthOf: (text) => text.length * 5,
			lineHeightOf: () => 16,
		});
		expect(result.approximations.join(' ')).toContain(
			name.startsWith('nested') ? 'Nested tables are projected' : 'Vertically merged',
		);
		const reopened = await loadDocx(await loaded.save(loaded.model));
		expect(reopened.model).toEqual(loaded.model);
		expect(reopened.model.footnotes).toHaveLength(1);
		const original = new Uint8Array(
			readFileSync(new URL('../__fixtures__/table-footnotes/' + name + '.docx', import.meta.url)),
		);
		const before = await JSZip.loadAsync(original),
			unchanged = await JSZip.loadAsync(await loaded.save(loaded.model));
		for (const path of Object.keys(before.files).filter((path) => !before.files[path]!.dir))
			expect(await unchanged.file(path)!.async('uint8array')).toEqual(
				await before.file(path)!.async('uint8array'),
			);
		const table = loaded.model.blocks.find((block) => block.type === 'table')!;
		if (table.type !== 'table') throw new Error('Expected table');
		table.rows[0]![0]!.paragraphs[0]!.runs[0]!.text += ' edit';
		const edited = await loadDocx(await loaded.save(loaded.model));
		const reopenedTable = edited.model.blocks.find((block) => block.type === 'table')!;
		if (reopenedTable.type !== 'table') throw new Error('Expected reopened table');
		expect(reopenedTable.rows[0]![0]!.nestedTables).toEqual(table.rows[0]![0]!.nestedTables);
		expect(reopenedTable.rows[0]![0]!.verticalMerge).toEqual(table.rows[0]![0]!.verticalMerge);
		expect(edited.model.footnotes).toEqual(loaded.model.footnotes);
	},
);
