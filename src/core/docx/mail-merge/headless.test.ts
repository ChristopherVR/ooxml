import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS } from '../../xml/index';
import { documentParagraphs } from '../cell-blocks';
import type { Block } from '../model';
import { loadDocx } from '../parse';
import { listMergeFields, mergeFields } from './index';
import { code, complex, marker, paragraph, run, simple, template } from './test-fixture';

const text = (blocks: readonly Block[]) =>
	documentParagraphs(blocks).map((item) => item.runs.map((r) => r.text).join(''));
const field = (instruction: string, cached = 'x') => complex(code(instruction), cached);
const documentXml = async (bytes: Uint8Array) =>
	(await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
const cell = (inner: string) => `<w:tc><w:tcPr><w:shd w:fill="FFFF00"/></w:tcPr>${inner}</w:tc>`;
const table = (...rows: string[]) =>
	`<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${rows.join('')}</w:tbl>`;
const row = (...cells: string[]) => `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cells.join('')}</w:tr>`;

describe('listMergeFields', () => {
	it('lists fields in every story with names as written, switches and regions', async () => {
		const loaded = await template(
			paragraph(simple('MERGEFIELD client.2.name \\* Upper')) +
				paragraph(field(' MERGEFIELD "First name" \\b "Dear " ')) +
				paragraph(field(' MERGEFIELD TableStart:Items ') + field(' MERGEFIELD amount__words ')) +
				paragraph(field(' MERGEFIELD TableEnd:Items ')) +
				paragraph(
					marker('begin') +
						code(' IF ') +
						field(' MERGEFIELD Kind ') +
						code(' = "a" "A" "B" ') +
						marker('separate') +
						run('A') +
						marker('end'),
				) +
				`<w:p><w:r><w:fldChar w:fldCharType="begin" w:fldLock="1"/></w:r>${code(' MERGEFIELD Locked ')}${marker('separate')}${run('kept')}${marker('end')}</w:p>` +
				'<w:sectPr><w:footerReference w:type="default" r:id="rId1"/></w:sectPr>',
			{
				'word/footer1.xml': `<w:ftr xmlns:w="${NS.w}">${paragraph(simple('MERGEFIELD Page-ref'))}</w:ftr>`,
				'word/_rels/document.xml.rels': `<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/footer" Target="footer1.xml"/></Relationships>`,
			},
		);
		const { fields, regions } = await listMergeFields(loaded.model);
		expect(fields.map((item) => [item.name, item.story, item.kind])).toEqual([
			['client.2.name', 'body', 'simple'],
			['First name', 'body', 'complex'],
			['amount__words', 'body', 'complex'],
			['Kind', 'body', 'complex'],
			['Locked', 'body', 'complex'],
			['Page-ref', 'footer', 'simple'],
		]);
		expect(fields[0]!.switches).toEqual(['\\* Upper']);
		expect(fields[1]!.switches).toEqual(['\\b "Dear "']);
		expect(fields[2]!.region).toBe('Items');
		expect(fields[3]!.nested).toBe(true);
		expect(fields[4]!.locked).toBe(true);
		expect(regions).toEqual([{ name: 'Items', kind: 'paragraphs', story: 'body' }]);
	});
});

describe('mergeFields', () => {
	it('fills simple and complex fields with split codes, line breaks and tabs, and unlinks', async () => {
		const loaded = await template(
			paragraph(run('To ') + simple('MERGEFIELD Name') + run('!')) +
				paragraph(
					marker('begin') +
						code(' MERGE') +
						code('FIELD Addr') +
						code('ess ') +
						marker('separate') +
						run('a') +
						run('b') +
						marker('end'),
				) +
				paragraph(field(' MERGEFIELD Missing ', 'old')),
		);
		const { model, results } = await mergeFields(
			loaded.model,
			{ values: { Name: 'Ada', Address: 'Line 1\nLine 2\tTabbed' } },
			{ unlink: true },
		);
		expect(text(model.blocks)).toEqual(['To Ada!', 'Line 1\nLine 2\tTabbed', '']);
		expect(results.map((item) => [item.name, item.status])).toEqual([
			['Name', 'filled'],
			['Address', 'filled'],
			['Missing', 'missing'],
		]);
		const xml = await documentXml(await loaded.save(model));
		expect(xml).toContain('<w:t>Line 1</w:t><w:br/><w:t>Line 2</w:t><w:tab/>');
		expect(xml).not.toContain('fldChar');
		expect(xml).not.toContain('fldSimple');
	});

	it('keeps fields with updated results unless unlinked, and keeps missing results on request', async () => {
		const loaded = await template(
			paragraph(
				field(' MERGEFIELD Name ', '«Name»') + run(' ') + field(' MERGEFIELD Other ', 'old'),
			),
		);
		const { model } = await mergeFields(
			loaded.model,
			{ values: { Name: 'Ada' } },
			{ missing: 'keep' },
		);
		const reopened = await loadDocx(await loaded.save(model));
		expect(text(reopened.model.blocks)).toEqual(['Ada old']);
		const xml = await documentXml(await loaded.save(model));
		expect(xml.match(/fldCharType="begin"/g)).toHaveLength(2);
		const unlinked = await mergeFields(
			loaded.model,
			{ values: {} },
			{ missing: 'keep', unlink: true },
		);
		expect(text(unlinked.model.blocks)).toEqual(['«Name» old']);
	});

	it('leaves locked fields as they are and reports them', async () => {
		const loaded = await template(
			`<w:p><w:r><w:fldChar w:fldCharType="begin" w:fldLock="1"/></w:r>${code(' MERGEFIELD Name ')}${marker('separate')}${run('kept')}${marker('end')}</w:p>`,
		);
		const { model, results } = await mergeFields(loaded.model, { values: { Name: 'Ada' } });
		expect(text(model.blocks)).toEqual(['kept']);
		expect(results).toEqual([{ name: 'Name', story: 'body', status: 'locked' }]);
	});

	it('fills headers, footnotes and fields nested in IF', async () => {
		const loaded = await template(
			paragraph(
				marker('begin') +
					code(' IF ') +
					field(' MERGEFIELD Kind ') +
					code(' = "a" "Alpha" "Beta" ') +
					marker('separate') +
					run('?') +
					marker('end'),
			) + '<w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr>',
			{
				'word/header1.xml': `<w:hdr xmlns:w="${NS.w}">${paragraph(run('Dear ') + simple('MERGEFIELD Name'))}</w:hdr>`,
				'word/_rels/document.xml.rels': `<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/header" Target="header1.xml"/></Relationships>`,
			},
		);
		const { model } = await mergeFields(
			loaded.model,
			{ values: { Kind: 'a', Name: 'Ada' } },
			{ unlink: true },
		);
		expect(text(model.blocks)).toEqual(['Alpha']);
		const reopened = await loadDocx(await loaded.save(model));
		expect(text(reopened.model.sections?.[0]?.headers?.default?.blocks ?? [])).toEqual([
			'Dear Ada',
		]);
	});

	it('repeats table rows with their properties and leaves a valid table when empty', async () => {
		const body =
			table(
				row(cell(paragraph(run('Item'))), cell(paragraph(run('Price')))),
				row(
					cell(paragraph(field(' MERGEFIELD TableStart:Items ', '') + field(' MERGEFIELD Name '))),
					cell(paragraph(field(' MERGEFIELD Price ') + field(' MERGEFIELD TableEnd:Items ', ''))),
				),
				row(cell(paragraph(run('Total'))), cell(paragraph(field(' MERGEFIELD Total ')))),
			) + paragraph(run('after'));
		const loaded = await template(body);
		const data = {
			values: { Total: '30' },
			regions: {
				Items: [
					{ Name: 'Pen', Price: '10' },
					{ Name: 'Ink', Price: '20' },
				],
			},
		};
		const merged = await mergeFields(loaded.model, data, { unlink: true });
		expect(merged.regions).toEqual([{ name: 'Items', kind: 'rows', story: 'body', rows: 2 }]);
		expect(text(merged.model.blocks)).toEqual([
			'Item',
			'Price',
			'Pen',
			'10',
			'Ink',
			'20',
			'Total',
			'30',
			'after',
		]);
		const bytes = await loaded.save(merged.model);
		const xml = await documentXml(bytes);
		expect(xml.match(/<w:tr>/g)).toHaveLength(4);
		expect(xml.match(/<w:cantSplit\/>/g)).toHaveLength(4);
		expect(xml.match(/w:fill="FFFF00"/g)).toHaveLength(8);
		expect(xml).not.toContain('TableStart');
		expect(text((await loadDocx(bytes)).model.blocks)).toEqual(text(merged.model.blocks));

		const empty = await mergeFields(loaded.model, { values: { Total: '0' } }, { unlink: true });
		const emptyXml = await documentXml(await loaded.save(empty.model));
		expect(emptyXml.match(/<w:tr>/g)).toHaveLength(2);
		expect(text(empty.model.blocks)).toEqual(['Item', 'Price', 'Total', '0', 'after']);
	});

	it('repeats paragraphs, falls back to values, and removes markers in one paragraph', async () => {
		const loaded = await template(
			paragraph(run('Head')) +
				paragraph(
					field(' MERGEFIELD TableStart:Lines ', '') +
						field(' MERGEFIELD Text ') +
						run(' / ') +
						field(' MERGEFIELD Shared ') +
						field(' MERGEFIELD TableEnd:Lines ', ''),
				) +
				paragraph(run('Tail')),
		);
		const merged = await mergeFields(
			loaded.model,
			{
				values: { Shared: 'S' },
				regions: { Lines: [{ Text: 'one' }, { Text: 'two', Shared: 'T' }] },
			},
			{ unlink: true },
		);
		expect(text(merged.model.blocks)).toEqual(['Head', 'one / S', 'two / T', 'Tail']);
		expect(merged.results.filter((item) => item.region).map((item) => item.row)).toEqual([
			0, 0, 1, 1,
		]);
		const reopened = await loadDocx(await loaded.save(merged.model));
		expect(text(reopened.model.blocks)).toEqual(['Head', 'one / S', 'two / T', 'Tail']);
		const none = await mergeFields(loaded.model, { values: {} }, { unlink: true });
		expect(text(none.model.blocks)).toEqual(['Head', 'Tail']);
	});

	it('removes paragraphs left empty by empty fields on request', async () => {
		const loaded = await template(
			paragraph(run('A')) + paragraph(field(' MERGEFIELD Line2 ')) + paragraph(run('B')),
		);
		const merged = await mergeFields(
			loaded.model,
			{ values: { Line2: '' } },
			{ unlink: true, removeEmptyParagraphs: true },
		);
		expect(text(merged.model.blocks)).toEqual(['A', 'B']);
	});
});

describe('mergeFields with HTML values', () => {
	it('replaces a field inline, inheriting its run formatting', async () => {
		const loaded = await template(
			paragraph(
				run('Hi ') +
					`<w:r><w:rPr><w:i/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:i/></w:rPr><w:instrText> MERGEFIELD Body </w:instrText></w:r><w:r><w:rPr><w:i/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>x</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>` +
					run('!'),
			),
		);
		const { model } = await mergeFields(
			loaded.model,
			{ values: { Body: 'a <b>bold</b> <a href="https://example.com/">link</a>' }, html: ['Body'] },
			{ unlink: true },
		);
		const runs = (model.blocks[0] as Extract<Block, { type: 'paragraph' }>).runs;
		expect(runs.map((item) => item.text).join('')).toBe('Hi a bold link!');
		const bold = runs.find((item) => item.text === 'bold')!;
		expect(bold).toMatchObject({ bold: true, italic: true });
		expect(runs.find((item) => item.text === 'link')!.link).toEqual({
			href: 'https://example.com/',
		});
		await loaded.save(model);
	});

	it('only treats names listed in html as HTML', async () => {
		const loaded = await template(paragraph(field(' MERGEFIELD Html ')));
		const { model } = await mergeFields(
			loaded.model,
			{ values: { Html: '<b>x</b>' } },
			{ unlink: true },
		);
		expect(text(model.blocks)).toEqual(['<b>x</b>']);
	});

	it('trims only the trailing spaces of a block', async () => {
		const loaded = await template(paragraph(field(' MERGEFIELD Body ')));
		const html = `<p>one \t </p><pre>a${' '.repeat(5)}b\t   </pre><pre>c${' '.repeat(50_000)}d </pre>`;
		const started = performance.now();
		const { model } = await mergeFields(loaded.model, { values: { Body: html }, html: ['Body'] });
		expect(performance.now() - started).toBeLessThan(1_000);
		expect(text(model.blocks)).toEqual(['one', 'a     b\t', `c${' '.repeat(50_000)}d`]);
	});

	it('replaces the host paragraph with blocks, lists and a real table', async () => {
		const loaded = await template(paragraph(field(' MERGEFIELD Body ')) + paragraph(run('end')));
		const html =
			'<p>First</p><ul><li>one</li><li>two</li></ul><table border="1"><tr><th>A</th><td>B</td></tr></table><p>Last</p>';
		const { model } = await mergeFields(loaded.model, { values: { Body: html }, html: ['Body'] });
		expect(model.blocks.map((block) => block.type)).toEqual([
			'paragraph',
			'paragraph',
			'paragraph',
			'table',
			'paragraph',
			'paragraph',
		]);
		const list = model.blocks[1] as Extract<Block, { type: 'paragraph' }>;
		expect(list.numbering?.level).toBe(0);
		expect(model.numberingCatalog?.nums[String(list.numbering!.numId)]).toBeDefined();
		const bytes = await loaded.save(model);
		const xml = await documentXml(bytes);
		expect(xml).toContain('<w:tbl>');
		expect(xml).toContain('<w:numId w:val=');
		expect(text((await loadDocx(bytes)).model.blocks)).toEqual([
			'First',
			'one',
			'two',
			'A',
			'B',
			'Last',
			'end',
		]);
	});

	it('inserts block HTML after the paragraph when the field shares it', async () => {
		const loaded = await template(paragraph(run('Notes: ') + field(' MERGEFIELD Body ')));
		const { model } = await mergeFields(
			loaded.model,
			{ values: { Body: '<p>One</p><p>Two</p>' }, html: ['Body'] },
			{ unlink: true },
		);
		expect(text(model.blocks)).toEqual(['Notes: ', 'One', 'Two']);
	});

	it('drops scripts, pictures and unsafe links and never fails on invalid HTML', async () => {
		const loaded = await template(paragraph(field(' MERGEFIELD Body ')));
		const { model, htmlIssues } = await mergeFields(
			loaded.model,
			{
				values: {
					Body: '<script>alert(1)</script><img src="https://evil.example/x.png"><a href="javascript:alert(1)">t</a><b>unclosed <i>',
				},
				html: ['Body'],
			},
			{ unlink: true },
		);
		expect(text(model.blocks)).toEqual(['tunclosed ']);
		expect(htmlIssues.sort()).toEqual(['image-refused', 'link-refused']);
		expect(JSON.stringify(model)).not.toContain('evil.example');
		expect(JSON.stringify(model)).not.toContain('javascript');
	});
});

describe('result formatting, as Word gives it', () => {
	// Measured with Word 16 (Update Field on the same runs, then Range.Font of the result):
	// no switch takes the last run of the field code, MERGEFORMAT keeps the old result's
	// formatting, CHARFORMAT takes the first letter of the field name.
	const r = (props: string, inner: string) => `<w:r><w:rPr>${props}</w:rPr>${inner}</w:r>`;
	const B = '<w:b/><w:color w:val="FF0000"/>';
	const I = '<w:i/>';
	const U = '<w:u w:val="single"/>';
	const instr = (props: string, text: string) =>
		r(props, `<w:instrText xml:space="preserve">${text}</w:instrText>`);
	const char = (kind: string) => r(U, `<w:fldChar w:fldCharType="${kind}"/>`);
	const formatted = (...code: string[]) =>
		paragraph(
			char('begin') + code.join('') + char('separate') + r(I, '<w:t>old</w:t>') + char('end'),
		);
	const cases = async () =>
		template(
			formatted(instr(B, ' MERGEFIELD Name ')) +
				formatted(instr(B, ' MERGEFIELD Name \\* MERGEFORMAT ')) +
				formatted(instr(U, ' M'), instr(B, 'ERGEFIELD Name \\* CHARFORMAT ')) +
				formatted(instr(B, ' MERGE'), instr(I, 'FIELD Na'), instr(U, 'me ')),
		);
	const shown = (blocks: readonly Block[]) =>
		documentParagraphs(blocks).map((item) => {
			const run = item.runs.find((candidate) => candidate.text === 'Ada')!;
			return [Boolean(run.bold), Boolean(run.italic), Boolean(run.underline)];
		});
	const expected = [
		[true, false, false],
		[false, true, false],
		[false, false, true],
		[false, false, true],
	];

	it('formats unlinked results', async () => {
		const loaded = await cases();
		const { model } = await mergeFields(
			loaded.model,
			{ values: { Name: 'Ada' } },
			{ unlink: true },
		);
		expect(shown(model.blocks)).toEqual(expected);
		expect(shown((await loadDocx(await loaded.save(model))).model.blocks)).toEqual(expected);
	});

	it('formats kept results', async () => {
		const loaded = await cases();
		const { model } = await mergeFields(loaded.model, { values: { Name: 'Ada' } });
		expect(shown((await loadDocx(await loaded.save(model))).model.blocks)).toEqual(expected);
	});
});
