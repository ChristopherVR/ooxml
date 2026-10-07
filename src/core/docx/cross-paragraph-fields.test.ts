import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx, type DocumentModel, type Paragraph } from './index';
import { parseBlocksFromContainer, parseParagraph } from './block-parser';
import { first, parseXml } from './xml';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const marker = (kind: string) => `<w:r><w:fldChar w:fldCharType="${kind}"/></w:r>`;
const code = (instr: string) =>
	`<w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r>`;
const text = (value: string) => `<w:r><w:t>${value}</w:t></w:r>`;
const p = (content: string) => `<w:p>${content}</w:p>`;
const table = (content: string) => `<w:tbl><w:tr><w:tc>${content}</w:tc></w:tr></w:tbl>`;
const start = (instr: string) => `${marker('begin')}${code(instr)}${marker('separate')}`;

async function loadBody(content: string) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}"><w:body>${content}<w:sectPr/></w:body></w:document>`,
	);
	return loadDocx(await zip.generateAsync({ type: 'uint8array' }));
}

const paragraphs = (model: DocumentModel): Paragraph[] =>
	model.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [block]
			: block.rows.flatMap((row) => row.flatMap((cell) => cell.paragraphs)),
	);
const results = (model: DocumentModel) =>
	paragraphs(model).flatMap((item) => item.runs.filter((run) => run.text));

describe('complex fields across paragraphs', () => {
	it('retains field metadata across paragraphs and after edited save/reparse', async () => {
		const loaded = await loadBody(
			p(start(' QUOTE ') + text('first')) + p(text('second') + marker('end') + text('outside')),
		);
		expect(results(loaded.model)).toEqual([
			{ text: 'first', field: { instr: 'QUOTE' } },
			{ text: 'second', field: { instr: 'QUOTE' } },
			{ text: 'outside' },
		]);
		const edited = structuredClone(loaded.model);
		paragraphs(edited)[1]!.runs[0]!.text = 'changed';
		const reloaded = await loadDocx(await loaded.save(edited));
		expect(results(reloaded.model)).toEqual([
			{ text: 'first', field: { instr: 'QUOTE' } },
			{ text: 'changed', field: { instr: 'QUOTE' } },
			{ text: 'outside' },
		]);
	});

	it('follows split instructions and nested deleted results across paragraphs', async () => {
		const deleted = '<w:del w:id="7" w:author="Ann"><w:r><w:delText>1</w:delText></w:r></w:del>';
		const loaded = await loadBody(
			p(marker('begin') + code(' QUO')) +
				p(code('TE ') + marker('separate') + text('before') + marker('begin') + code(' PA')) +
				p(code('GE ') + marker('separate') + deleted) +
				p(marker('end') + text('after') + marker('end')),
		);
		expect(results(loaded.model)).toEqual([
			{ text: 'before', field: { instr: 'QUOTE' } },
			{ text: '1', field: { instr: 'PAGE' }, revision: { kind: 'delete', id: '7', author: 'Ann' } },
			{ text: 'after', field: { instr: 'QUOTE' } },
		]);
		const edited = structuredClone(loaded.model);
		paragraphs(edited)[2]!.runs.find((run) => run.text === '1')!.text = '2';
		const reloaded = await loadDocx(await loaded.save(edited));
		expect(results(reloaded.model)[1]).toMatchObject({
			text: '2',
			field: { instr: 'PAGE' },
			revision: { kind: 'delete', id: '7', author: 'Ann' },
		});
	});

	it('traverses nested table XML in story order without making nested tables editable', async () => {
		const loaded = await loadBody(
			p(start(' QUOTE ') + text('before')) +
				table(
					p(text('cell-before')) +
						table(p(start(' PAGE ') + text('1') + marker('end'))) +
						p(text('cell-after')),
				) +
				p(text('after') + marker('end') + text('outside')),
		);
		expect(results(loaded.model)).toEqual([
			{ text: 'before', field: { instr: 'QUOTE' } },
			{ text: 'cell-before', field: { instr: 'QUOTE' } },
			{ text: 'cell-after', field: { instr: 'QUOTE' } },
			{ text: 'after', field: { instr: 'QUOTE' } },
			{ text: 'outside' },
		]);
		const original = loaded.model.blocks[1]!;
		expect(original.type).toBe('table');
		if (original.type !== 'table') throw new Error('Expected table');
		expect(original.structureEditable).toBe(false);
		expect(original.rows[0]![0]!.nestedTables).toEqual([{ rows: [[{ text: '1' }]] }]);
		const edited = structuredClone(loaded.model);
		paragraphs(edited)[2]!.runs[0]!.text = 'changed-cell';
		const reloaded = await loadDocx(await loaded.save(edited));
		expect(results(reloaded.model)[2]).toEqual({ text: 'changed-cell', field: { instr: 'QUOTE' } });
		const savedTable = reloaded.model.blocks[1]!;
		if (savedTable.type !== 'table') throw new Error('Expected table');
		expect(savedTable.rows[0]![0]!.nestedTables).toEqual(original.rows[0]![0]!.nestedTables);
	});

	it('observes end markers inside read-only nested tables before later modeled paragraphs', async () => {
		const loaded = await loadBody(
			p(start(' QUOTE ') + text('before')) +
				table(
					p(text('cell-before')) + table(p(text('nested') + marker('end'))) + p(text('cell-after')),
				) +
				p(text('outside')),
		);
		expect(results(loaded.model)).toEqual([
			{ text: 'before', field: { instr: 'QUOTE' } },
			{ text: 'cell-before', field: { instr: 'QUOTE' } },
			{ text: 'cell-after' },
			{ text: 'outside' },
		]);
		const edited = structuredClone(loaded.model);
		paragraphs(edited)[2]!.runs[0]!.text = 'changed';
		const reloaded = await loadDocx(await loaded.save(edited));
		expect(results(reloaded.model)[2]).toEqual({ text: 'changed' });
	});

	it('keeps simple caches and textbox fields from changing the surrounding story tracker', async () => {
		const simple = `<w:fldSimple w:instr=" AUTHOR ">${text('Ann')}</w:fldSimple>`;
		const textbox = `<w:r><w:pict><w:txbxContent>${p(start(' PAGE ') + text('1'))}</w:txbxContent></w:pict></w:r>`;
		const loaded = await loadBody(
			p(start(' QUOTE ') + simple + textbox) + p(text('after') + marker('end')),
		);
		expect(results(loaded.model)).toContainEqual({
			text: 'Ann',
			field: { instr: 'AUTHOR', simple: true },
			fieldInstanceId: 'p0:simple-field-0',
		});
		expect(results(loaded.model)).toContainEqual({ text: 'after', field: { instr: 'QUOTE' } });
	});

	it('isolates consecutive story containers and keeps standalone paragraph parsing local', () => {
		const body = parseXml(
			`<w:body xmlns:w="${W}">${p(start(' QUOTE ') + text('open'))}</w:body>`,
		).documentElement!;
		const header = parseXml(
			`<w:hdr xmlns:w="${W}">${p(text('header') + marker('end'))}</w:hdr>`,
		).documentElement!;
		parseBlocksFromContainer(body);
		const headerBlocks = parseBlocksFromContainer(header);
		expect((headerBlocks[0] as Paragraph).runs[0]).toEqual({ text: 'header' });
		const next = parseXml(
			`<w:p xmlns:w="${W}">${text('next')}${marker('end')}</w:p>`,
		).documentElement!;
		parseParagraph(first(body, 'p')!, 'p0');
		expect(parseParagraph(next, 'p1').runs[0]).toEqual({ text: 'next' });
	});
});
