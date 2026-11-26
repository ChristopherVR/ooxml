import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createDocument, loadDocx, type ContentControl, type Paragraph } from '../index';
import { adaptDocumentModel } from './adapter';
import { layoutDocumentModel } from './layout';
import { createFakeMeasurer } from './measure';
import { layoutParagraph } from './paragraph-layout';
import { at } from './test-helpers';

const measurer = createFakeMeasurer();
const fixtures = './fixtures/word-layout-reports/';

it('lays out one pair of tags for each paragraph, row and cell group', async () => {
	const { model } = await loadDocx(
		await readFile(new URL(fixtures + 'block-content-control-tags.docx', import.meta.url)),
	);
	const { pages } = layoutDocumentModel(model, measurer);
	const paragraphs = pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph' ? [block] : block.rows.flatMap((row) => row.cells.flat()),
			),
		),
	);
	expect(
		paragraphs.map((paragraph) =>
			paragraph.lines.flatMap((line) =>
				line.fragments
					.filter((fragment) => fragment.contentControlTag)
					.map((tag) => [
						tag.contentControlTag,
						tag.text,
						tag.sourceStart,
						tag.sourceEnd,
						tag.widthPx > 0,
					]),
			),
		),
	).toEqual([
		[['start', 'Block', 0, 0, true]],
		[],
		[['end', '', 4, 4, true]],
		[
			['start', 'Rows', 0, 0, true],
			['start', 'Cells', 0, 0, true],
		],
		[['end', '', 10, 10, true]],
		[],
		[['end', '', 7, 7, true]],
		[],
	]);
});

it('wraps the block title with its first word and keeps the closing tag with the last word', () => {
	const contentControl: ContentControl = {
		id: '8',
		level: 'block',
		kind: 'richText',
		alias: 'A label',
		appearance: 'tags',
	};
	const document = {
		...createDocument(),
		blocks: [{ type: 'paragraph' as const, id: 'p', runs: [{ text: 'Lorem ipsum' }] }],
	};
	const linesFor = (paragraph: Paragraph) => {
		const section = at(adaptDocumentModel({ ...document, blocks: [paragraph] }).sections, 0);
		const block = at(section.blocks, 0);
		if (block.kind !== 'paragraph') throw new Error('Expected a paragraph');
		return layoutParagraph(block, 115, measurer, () => {}).lines.map((line) =>
			line.fragments.map((fragment) => fragment.contentControlTag ?? fragment.text).join(''),
		);
	};
	const paragraph = at(document.blocks, 0);
	expect(linesFor(paragraph)).toEqual(['Lorem ipsum']);
	expect(linesFor({ ...paragraph, contentControl })).toEqual(['startLorem', 'ipsumend']);
});

it('clears a tagged appearance on save and preserves the remaining properties', async () => {
	const { model, save } = await loadDocx(
		await readFile(new URL(fixtures + 'issue-349-content-control-tags.docx', import.meta.url)),
	);
	const first = at(model.blocks, 0);
	if (first.type !== 'paragraph') throw new Error('Expected a paragraph');
	const run = first.runs.find((item) => item.text === 'alpha');
	expect(run?.contentControl?.appearance).toBe('tags');
	if (!run?.contentControl) throw new Error('Expected an inline control');
	run.text = 'updated';
	delete run.contentControl.appearance;
	const reopened = await loadDocx(await save(model));
	const paragraph = at(reopened.model.blocks, 0);
	if (paragraph.type !== 'paragraph') throw new Error('Expected a paragraph');
	const control = paragraph.runs.find((item) => item.text === 'updated')?.contentControl;
	expect(control).toMatchObject({ alias: 'Condition', tag: 'Condition' });
	expect(control?.appearance).toBeUndefined();
	expect(control?.sourceXml).not.toContain('appearance');
	expect(paragraph.runs.find((item) => item.text === 'beta')?.contentControl?.appearance).toBe(
		'tags',
	);
});

it('saves appearance edits on a paragraph group, a row group and a cell group', async () => {
	const { model, save } = await loadDocx(
		await readFile(new URL(fixtures + 'block-content-control-tags.docx', import.meta.url)),
	);
	for (const block of model.blocks) {
		if (block.contentControl) block.contentControl.appearance = 'hidden';
		if (block.type !== 'table') continue;
		for (const row of block.rowProperties ?? []) {
			if (row.contentControl) delete row.contentControl.appearance;
		}
		for (const cell of block.rows.flat()) {
			if (cell.contentControl) delete cell.contentControl.appearance;
		}
	}
	const { model: reopened } = await loadDocx(await save(model));
	const paragraph = at(reopened.blocks, 0);
	expect(paragraph.contentControl?.appearance).toBe('hidden');
	expect(paragraph.contentControl?.sourceXml).toMatch(/appearance[^>]*\/><w:id/);
	const table = reopened.blocks.find((block) => block.type === 'table');
	if (!table) throw new Error('Expected a table');
	const rows = table.rowProperties?.map((row) => row.contentControl);
	const cells = at(table.rows, 0).map((cell) => cell.contentControl);
	expect(rows?.map((control) => [control?.alias, control?.appearance])).toEqual([
		['Rows', undefined],
		['Rows', undefined],
	]);
	expect(cells.map((control) => [control?.alias, control?.appearance])).toEqual([
		['Cells', undefined],
		['Cells', undefined],
	]);
});

it('reserves line space for each tagged control in the Word fixture (#349)', async () => {
	const layout = async (name: string) => {
		const { model } = await loadDocx(
			await readFile(
				new URL(`./fixtures/word-layout-reports/issue-349-${name}.docx`, import.meta.url),
			),
		);
		const { pages } = layoutDocumentModel(model, measurer);
		const fragments = pages.flatMap((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) =>
					block.kind === 'paragraph' ? block.lines.flatMap((line) => line.fragments) : [],
				),
			),
		);
		return { pages, tags: fragments.filter((fragment) => fragment.contentControlTag) };
	};
	const control = await layout('content-control-tags-control');
	expect(control.tags).toEqual([]);
	expect(control.pages).toHaveLength(2);
	const tagged = await layout('content-control-tags');
	expect(tagged.tags).toHaveLength(240);
	expect(
		tagged.tags.filter((tag) => tag.contentControlTag === 'start').map((tag) => tag.text),
	).toEqual(Array.from({ length: 40 }, () => ['Condition', 'Then', 'Else']).flat());
	expect(tagged.tags.filter((tag) => tag.contentControlTag === 'end')).toHaveLength(120);
	expect(tagged.tags.every((tag) => tag.widthPx > 0 && tag.sourceStart === tag.sourceEnd)).toBe(
		true,
	);
	expect(tagged.pages).toHaveLength(3);
});

it('preserves the fixture appearance when edited content is saved', async () => {
	const { model, save } = await loadDocx(
		await readFile(
			new URL(
				'./fixtures/word-layout-reports/issue-349-content-control-tags.docx',
				import.meta.url,
			),
		),
	);
	const first = at(model.blocks, 0);
	if (first.type !== 'paragraph') throw new Error('Expected a paragraph');
	const edit = {
		...first,
		runs: first.runs.map((run) => (run.text === 'alpha' ? { ...run, text: 'modified' } : run)),
	};
	const reopened = await loadDocx(
		await save({ ...model, blocks: [edit, ...model.blocks.slice(1)] }),
	);
	const paragraph = at(reopened.model.blocks, 0);
	if (paragraph.type !== 'paragraph') throw new Error('Expected a paragraph');
	expect(paragraph.runs.find((run) => run.text === 'modified')?.contentControl).toMatchObject({
		alias: 'Condition',
		appearance: 'tags',
	});
});

it('keeps the two tags next to a multi-run control and maps them to its source boundaries', () => {
	const control: ContentControl = {
		id: '100',
		level: 'inline',
		kind: 'richText',
		alias: 'A label',
		appearance: 'tags',
	};
	const paragraph: Paragraph = {
		type: 'paragraph',
		id: 'p',
		runs: [
			{ text: 'First ' },
			{ text: 'wo', contentControl: control },
			{ text: 'rd', bold: true, contentControl: { ...control } },
			{ text: ' tail' },
		],
	};
	const block = at(
		at(adaptDocumentModel({ ...createDocument(), blocks: [paragraph] }).sections, 0).blocks,
		0,
	);
	if (block.kind !== 'paragraph') throw new Error('Expected a paragraph');
	const { lines } = layoutParagraph(block, 90, measurer, () => {});
	expect(lines.map((line) => line.fragments.map((fragment) => fragment.text).join(''))).toEqual([
		'First',
		'A labelword',
		'tail',
	]);
	const tagged = at(lines, 1).fragments;
	expect(at(tagged, 0)).toMatchObject({ contentControlTag: 'start', sourceStart: 6, sourceEnd: 6 });
	expect(tagged.at(-1)).toMatchObject({ contentControlTag: 'end', sourceStart: 10, sourceEnd: 10 });
});
