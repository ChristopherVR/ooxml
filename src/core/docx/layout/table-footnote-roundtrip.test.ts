import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
const measurer = { widthOf: (text: string) => text.length * 5, lineHeightOf: () => 16 };
it.each([
	'split-late-reference',
	'repeated-header',
	'exact-clipped-reference',
	'mode14-split-long-note-columns',
])('%s preserves note stories and table properties through a cell edit', async (name) => {
	const bytes = new Uint8Array(
			readFileSync(new URL('../__fixtures__/table-footnotes/' + name + '.docx', import.meta.url)),
		),
		loaded = await loadDocx(bytes),
		source = await JSZip.loadAsync(bytes);
	const unchanged = await JSZip.loadAsync(await loaded.save(loaded.model));
	for (const path of Object.keys(source.files).filter((path) => !source.files[path]!.dir))
		expect(await unchanged.file(path)!.async('uint8array')).toEqual(
			await source.file(path)!.async('uint8array'),
		);
	const table = loaded.model.blocks.find((block) => block.type === 'table')!;
	if (table.type !== 'table') throw new Error('Expected table');
	const properties = structuredClone(table.rowProperties),
		notes = structuredClone(loaded.model.footnotes);
	const run = table.rows[0]![0]!.paragraphs[0]!.runs.find((run) => run.text.includes('T0R'))!;
	run.text = run.text.replace('T0R', 'E0R');
	const expected = layoutDocumentModel(loaded.model, measurer),
		saved = await loaded.save(loaded.model),
		zip = await JSZip.loadAsync(saved);
	for (const path of Object.keys(source.files).filter(
		(path) => !source.files[path]!.dir && path !== 'word/document.xml',
	))
		expect(await zip.file(path)!.async('uint8array')).toEqual(
			await source.file(path)!.async('uint8array'),
		);
	const reopened = await loadDocx(saved),
		newTable = reopened.model.blocks.find((block) => block.type === 'table')!;
	if (newTable.type !== 'table') throw new Error('Expected reopened table');
	expect(newTable.rowProperties).toEqual(properties);
	expect(reopened.model.footnotes).toEqual(notes);
	expect(layoutDocumentModel(reopened.model, measurer).pages.map((page) => page.footnotes)).toEqual(
		expected.pages.map((page) => page.footnotes),
	);
	expect(newTable.rows[0]![0]!.paragraphs[0]!.runs.some((run) => run.text.includes('E0R'))).toBe(
		true,
	);
});
