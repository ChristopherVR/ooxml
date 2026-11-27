// The headless merge through the `ooxml-core/docx` entry, saved with the template's own writer
// and read back: every story and slot, kept and unlinked fields, the options and what is kept.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS } from '../../xml/index';
import { documentParagraphs } from '../cell-blocks';
import { listMergeFields, mergeFields } from '../index';
import type { Block } from '../model';
import { loadDocx } from '../parse';
import { code, complex, paragraph, run, simple, template } from './test-fixture';

const text = (blocks: readonly Block[]) =>
	documentParagraphs(blocks).map((item) => item.runs.map((r) => r.text).join(''));
const field = (instruction: string, cached = 'x') => complex(code(instruction), cached);
const part = async (bytes: Uint8Array, path: string) =>
	(await JSZip.loadAsync(bytes)).file(path)?.async('string');
const rels = (...items: [id: string, type: string, target: string][]) =>
	`<Relationships xmlns="${NS.rels}">${items
		.map(
			([id, type, target]) =>
				`<Relationship Id="${id}" Type="${NS.r}/${type}" Target="${target}"/>`,
		)
		.join('')}</Relationships>`;
const story = (root: 'hdr' | 'ftr', label: string) =>
	`<w:${root} xmlns:w="${NS.w}">${paragraph(run(`${label} `) + simple(`MERGEFIELD ${label}`))}</w:${root}>`;

async function everyStory() {
	const slots = ['default', 'first', 'even'] as const;
	return template(
		paragraph(run('Body ') + field(' MERGEFIELD Body ')) +
			paragraph('<w:r><w:endnoteReference w:id="2"/></w:r>') +
			`<w:sectPr>${slots
				.map(
					(slot, index) =>
						`<w:headerReference w:type="${slot}" r:id="rH${index}"/><w:footerReference w:type="${slot}" r:id="rF${index}"/>`,
				)
				.join('')}<w:titlePg/></w:sectPr>`,
		{
			...Object.fromEntries(
				slots.flatMap((slot, index) => [
					[`word/header${index}.xml`, story('hdr', `H${slot}`)],
					[`word/footer${index}.xml`, story('ftr', `F${slot}`)],
				]),
			),
			'word/endnotes.xml': `<w:endnotes xmlns:w="${NS.w}"><w:endnote w:id="2">${paragraph(run('End ') + field(' MERGEFIELD Note '))}</w:endnote></w:endnotes>`,
			'word/settings.xml': `<w:settings xmlns:w="${NS.w}"><w:evenAndOddHeaders/></w:settings>`,
			'word/_rels/document.xml.rels': rels(
				...slots.flatMap((_, index): [string, string, string][] => [
					[`rH${index}`, 'header', `header${index}.xml`],
					[`rF${index}`, 'footer', `footer${index}.xml`],
				]),
				['rE', 'endnotes', 'endnotes.xml'],
				['rS', 'settings', 'settings.xml'],
			),
		},
	);
}

const VALUES = {
	Body: 'B',
	Note: 'N',
	Hdefault: 'hd',
	Hfirst: 'hf',
	Heven: 'he',
	Fdefault: 'fd',
	Ffirst: 'ff',
	Feven: 'fe',
};

describe('mergeFields from ooxml-core/docx, saved and reloaded', () => {
	it('lists and fills every header and footer slot and the endnotes', async () => {
		const loaded = await everyStory();
		const { fields } = await listMergeFields(loaded.model);
		expect(fields.map((item) => [item.name, item.story, item.slot ?? null])).toEqual([
			['Body', 'body', null],
			['Hdefault', 'header', 'default'],
			['Hfirst', 'header', 'first'],
			['Heven', 'header', 'even'],
			['Fdefault', 'footer', 'default'],
			['Ffirst', 'footer', 'first'],
			['Feven', 'footer', 'even'],
			['Note', 'endnote', null],
		]);
		for (const unlink of [true, false]) {
			const { model, results } = await mergeFields(loaded.model, { values: VALUES }, { unlink });
			expect(results.every((item) => item.status === 'filled')).toBe(true);
			expect(results.find((item) => item.name === 'Ffirst')).toMatchObject({
				story: 'footer',
				slot: 'first',
			});
			const bytes = await loaded.save(model);
			for (const [index, slot] of ['default', 'first', 'even'].entries()) {
				const header = (await part(bytes, `word/header${index}.xml`))!;
				expect(header).toContain(`>${VALUES[`H${slot}` as keyof typeof VALUES]}<`);
				expect(header.includes('MERGEFIELD')).toBe(!unlink);
			}
			const reopened = await loadDocx(bytes);
			const section = reopened.model.sections?.[0];
			expect(text(section?.headers?.first?.blocks ?? [])).toEqual(['Hfirst hf']);
			expect(text(section?.headers?.even?.blocks ?? [])).toEqual(['Heven he']);
			expect(text(section?.footers?.default?.blocks ?? [])).toEqual(['Fdefault fd']);
			expect(text(section?.footers?.first?.blocks ?? [])).toEqual(['Ffirst ff']);
			expect(reopened.model.endnotes?.map((note) => text(note.blocks))).toContainEqual(['End N']);
			expect(text(reopened.model.blocks)[0]).toBe('Body B');
		}
	});

	it('keeps fields as fields with their new cached results, so the output can merge again', async () => {
		const loaded = await template(
			paragraph(simple('MERGEFIELD Name')) + paragraph(field(' MERGEFIELD City \\* Upper ')),
		);
		const { model } = await mergeFields(loaded.model, { values: { Name: 'Ada', City: 'paris' } });
		const bytes = await loaded.save(model);
		const xml = (await part(bytes, 'word/document.xml'))!;
		expect(xml).toMatch(/<w:fldSimple w:instr="MERGEFIELD Name"[^>]*>.*Ada.*<\/w:fldSimple>/s);
		expect(xml).toContain('MERGEFIELD City \\* Upper');
		expect(xml).toContain('>PARIS<');
		const reopened = await loadDocx(bytes);
		expect(text(reopened.model.blocks)).toEqual(['Ada', 'PARIS']);
		const again = await mergeFields(
			reopened.model,
			{ values: { Name: 'Bo', City: 'rome' } },
			{ unlink: true },
		);
		expect(text(again.model.blocks)).toEqual(['Bo', 'ROME']);
	});

	it('keeps comments, tracked changes, bookmarks and custom XML parts', async () => {
		const loaded = await template(
			`<w:p><w:commentRangeStart w:id="0"/>${run('Dear ')}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>${field(' MERGEFIELD Name ')}<w:ins w:id="5" w:author="Rev" w:date="2026-01-01T00:00:00Z">${run(' added')}</w:ins><w:bookmarkStart w:id="1" w:name="anchor"/>${run('.')}<w:bookmarkEnd w:id="1"/></w:p>`,
			{
				'word/comments.xml': `<w:comments xmlns:w="${NS.w}"><w:comment w:id="0" w:author="Rev">${paragraph(run('check'))}</w:comment></w:comments>`,
				'word/_rels/document.xml.rels': rels(['rC', 'comments', 'comments.xml']),
				'customXml/item1.xml': '<data xmlns="urn:example"><id>42</id></data>',
			},
		);
		const { model } = await mergeFields(
			loaded.model,
			{ values: { Name: 'Ada' } },
			{ unlink: true },
		);
		const bytes = await loaded.save(model);
		const xml = (await part(bytes, 'word/document.xml'))!;
		expect(xml).toContain('<w:commentRangeStart w:id="0"/>');
		expect(xml).toContain('<w:commentReference w:id="0"/>');
		expect(xml).toMatch(/<w:ins w:id="5" w:author="Rev"[^>]*>.*added/s);
		expect(xml).toContain('w:name="anchor"');
		expect(await part(bytes, 'customXml/item1.xml')).toBe(
			'<data xmlns="urn:example"><id>42</id></data>',
		);
		expect(await part(bytes, 'word/comments.xml')).toContain('check');
		expect(text((await loadDocx(bytes)).model.blocks)).toEqual(['Dear Ada added.']);
	});

	it('refreshes computed fields afterwards unless updateFields is false', async () => {
		const loaded = await template(
			paragraph(field(' MERGEFIELD Name ') + run(' ') + field(' = 2 + 3 ', 'old')),
		);
		const values = { values: { Name: 'Ada' } };
		const updated = await mergeFields(loaded.model, values, { unlink: true });
		expect(text((await loadDocx(await loaded.save(updated.model))).model.blocks)).toEqual([
			'Ada 5',
		]);
		const kept = await mergeFields(loaded.model, values, { unlink: true, updateFields: false });
		expect(text(kept.model.blocks)).toEqual(['Ada old']);
	});

	it('says why a field was not filled, or was filled without its picture', async () => {
		const loaded = await template(
			paragraph(field(' MERGEFIELD Absent ')) +
				paragraph(field(' MERGEFIELD Amount \\# "0.00" ')) +
				paragraph(field(' MERGEFIELD When \\@ "d MMMM yyyy" ')) +
				paragraph(field(' MERGEFIELD Amount \\# "0.00" ')) +
				paragraph(field(' MERGEFIELD TableStart:Rows ', '') + field(' MERGEFIELD Cell ')) +
				paragraph(field(' MERGEFIELD TableEnd:Rows ', '')),
		);
		const first = await mergeFields(
			loaded.model,
			{ values: { Amount: 'abc', When: 'soon' }, regions: { Rows: [{}] } },
			{ unlink: true },
		);
		expect(first.results.map(({ name, status, reason }) => [name, status, reason])).toEqual([
			['Absent', 'missing', 'no value named "Absent"'],
			['Amount', 'filled', 'the value is not a number, so the \\# picture was not applied'],
			['When', 'filled', 'the value is not a date, so the \\@ picture was not applied'],
			['Amount', 'filled', 'the value is not a number, so the \\# picture was not applied'],
			['Cell', 'missing', 'no value named "Cell" in row 0 of region Rows or in values'],
		]);
		// The region covers two paragraphs, so its one row writes two.
		expect(text(first.model.blocks)).toEqual(['', 'abc', 'soon', 'abc', '', '']);
		const second = await mergeFields(
			loaded.model,
			{ values: { Amount: '3.5', When: '2026-03-04', Absent: '' } },
			{ unlink: true },
		);
		expect(second.results.every((item) => item.status === 'filled' && !item.reason)).toBe(true);
		expect(text(second.model.blocks)).toEqual(['', '3.50', '4 March 2026', '3.50']);
	});

	it('reports a merge field in the code of a field other than IF as unsupported', async () => {
		const loaded = await template(
			paragraph(
				complex(code(' QUOTE ') + field(' MERGEFIELD Name ') + code(' '), 'outer') + run('!'),
			),
		);
		const { model, results } = await mergeFields(
			loaded.model,
			{ values: { Name: 'Ada' } },
			{ unlink: true },
		);
		expect(results).toHaveLength(1);
		expect(results[0]).toMatchObject({ status: 'unsupported' });
		expect(results[0]!.reason).toContain('other than IF');
		// The QUOTE field stays a field, with its code and its saved result.
		const xml = (await part(await loaded.save(model), 'word/document.xml'))!;
		expect(xml).toContain(' QUOTE ');
		expect(xml).toContain('>outer<');
	});

	it('reports an unreadable merge field as invalid', async () => {
		const loaded = await template(paragraph(field(' MERGEFIELD one two ', 'kept')));
		const { results } = await mergeFields(loaded.model, { values: { one: '1' } });
		expect(results.map(({ status, reason }) => [status, reason])).toEqual([
			['invalid', 'the field instruction could not be read'],
		]);
	});

	it('repeats a region spanning table rows and keeps one paragraph in emptied cells', async () => {
		const cell = (inner: string) =>
			`<w:tc><w:tcPr><w:tcBorders><w:top w:val="double"/></w:tcBorders></w:tcPr>${inner}</w:tc>`;
		const tr = (inner: string) => `<w:tr>${cell(inner)}</w:tr>`;
		const loaded = await template(
			`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${
				tr(paragraph(field(' MERGEFIELD TableStart:People ', '') + field(' MERGEFIELD Name '))) +
				tr(paragraph(field(' MERGEFIELD Role ') + field(' MERGEFIELD TableEnd:People ', ''))) +
				tr(paragraph(field(' MERGEFIELD Empty ')))
			}</w:tbl>${paragraph(run('end'))}`,
		);
		const { model, regions } = await mergeFields(
			loaded.model,
			{
				values: { Empty: '' },
				regions: {
					People: [
						{ Name: 'Ada', Role: 'Lead' },
						{ Name: 'Bo', Role: 'Dev' },
					],
				},
			},
			{ unlink: true, removeEmptyParagraphs: true },
		);
		expect(regions).toEqual([{ name: 'People', kind: 'rows', story: 'body', rows: 2 }]);
		const bytes = await loaded.save(model);
		const xml = (await part(bytes, 'word/document.xml'))!;
		expect(xml.match(/<w:tr>/g)).toHaveLength(5);
		expect(xml.match(/<w:top w:val="double"\/>/g)).toHaveLength(5);
		expect(xml.match(/<w:tc>.*?<\/w:tc>/gs)!.every((tc) => tc.includes('<w:p'))).toBe(true);
		expect(text((await loadDocx(bytes)).model.blocks)).toEqual([
			'Ada',
			'Lead',
			'Bo',
			'Dev',
			'',
			'end',
		]);
	});

	it('writes an HTML table value as a real table and inline HTML into a kept field', async () => {
		const loaded = await template(
			paragraph(field(' MERGEFIELD Grid ')) + paragraph(run('Hi ') + field(' MERGEFIELD Who ')),
		);
		const { model } = await mergeFields(loaded.model, {
			values: {
				Grid: '<table border="1"><tr><td>a</td><td>b</td></tr><tr><td>c</td><td>d</td></tr></table>',
				Who: '<b>Ada</b>',
			},
			html: ['Grid', 'Who'],
		});
		const bytes = await loaded.save(model);
		const xml = (await part(bytes, 'word/document.xml'))!;
		expect(xml.match(/<w:tbl>/g)).toHaveLength(1);
		expect(xml.match(/<w:tr>/g)).toHaveLength(2);
		expect(xml).not.toContain('MERGEFIELD Grid');
		expect(xml).toContain('MERGEFIELD Who');
		const reopened = await loadDocx(bytes);
		expect(reopened.model.blocks[0]?.type).toBe('table');
		expect(text(reopened.model.blocks)).toEqual(['a', 'b', 'c', 'd', 'Hi Ada']);
		const who = documentParagraphs(reopened.model.blocks).at(-1)!;
		expect(who.runs.find((item) => item.text === 'Ada')).toMatchObject({ bold: true });
	});
});
