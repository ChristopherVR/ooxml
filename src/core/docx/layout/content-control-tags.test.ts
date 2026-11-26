import { readFile } from 'node:fs/promises';
import { expect, it } from 'vitest';
import { createDocument, loadDocx, type ContentControl, type Paragraph } from '../index';
import { adaptDocumentModel } from './adapter';
import { layoutDocumentModel } from './layout';
import { createFakeMeasurer } from './measure';
import { layoutParagraph } from './paragraph-layout';
import { at } from './test-helpers';

const measurer = createFakeMeasurer();

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
