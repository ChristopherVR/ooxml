import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS } from '../../xml/index';
import { documentParagraphs } from '../cell-blocks';
import type { Block, Paragraph } from '../model';
import { loadDocx } from '../parse';
import { listMergeFields, mergeFields } from './index';
import { code, complex, paragraph, run, simple, template } from './test-fixture';

const text = (blocks: readonly Block[]) =>
	documentParagraphs(blocks).map((item) => item.runs.map((r) => r.text).join(''));
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const textBox = (inner: string) =>
	`<w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps"><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" behindDoc="0"><wp:extent cx="1905000" cy="952500"/><wp:docPr id="7" name="Text Box 7"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="${WPS}"><wps:wsp xmlns:wps="${WPS}"><wps:spPr/><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><w:t>stale</w:t></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
const nestedTable = (inner: string) =>
	`<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>${inner}</w:tc></w:tr></w:tbl>`;

describe('mergeFields in every story', () => {
	it('fills text boxes, footnotes, nested tables and content controls, keeping the controls', async () => {
		const loaded = await template(
			paragraph(textBox(paragraph(run('Box ') + simple('MERGEFIELD Name')))) +
				nestedTable(
					paragraph(run('outer')) +
						nestedTable(paragraph(simple('MERGEFIELD Name'))) +
						paragraph(run('x')),
				) +
				`<w:p><w:sdt><w:sdtPr><w:tag w:val="keep-me"/></w:sdtPr><w:sdtContent>${run('In ')}${complex(code(' MERGEFIELD Name '), 'x')}</w:sdtContent></w:sdt></w:p>` +
				`<w:p><w:bookmarkStart w:id="0" w:name="mark"/>${simple('MERGEFIELD Name')}<w:bookmarkEnd w:id="0"/></w:p>` +
				paragraph('<w:r><w:footnoteReference w:id="1"/></w:r>'),
			{
				'word/footnotes.xml': `<w:footnotes xmlns:w="${NS.w}"><w:footnote w:id="1">${paragraph(run('Note ') + simple('MERGEFIELD Name'))}</w:footnote></w:footnotes>`,
			},
		);
		expect((await listMergeFields(loaded.model)).fields.map((field) => field.story)).toEqual([
			'textBox',
			'body',
			'body',
			'body',
			'footnote',
		]);
		const { model, results } = await mergeFields(
			loaded.model,
			{ values: { Name: 'Ada' } },
			{ unlink: true },
		);
		expect(results.map((result) => result.status)).toEqual(Array(5).fill('filled'));
		const bytes = await loaded.save(model);
		const xml = await (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
		expect(xml).toContain('<w:tag w:val="keep-me"/>');
		expect(xml).toContain('w:name="mark"');
		expect(xml).not.toContain('MERGEFIELD');
		const reopened = await loadDocx(bytes);
		expect(text(reopened.model.blocks)).toEqual(['', 'outer', 'Ada', 'x', 'In Ada', 'Ada', '']);
		const box = (reopened.model.blocks[0] as Paragraph).runs.find((r) => r.image)!.image!;
		expect(text(box.textBoxContent!.blocks)).toEqual(['Box Ada']);
		expect(reopened.model.footnotes?.map((note) => text(note.blocks))).toContainEqual(['Note Ada']);
	});

	it('repeats a paragraph region in a header', async () => {
		const loaded = await template(
			paragraph(run('body')) +
				'<w:sectPr><w:headerReference w:type="default" r:id="rId1"/></w:sectPr>',
			{
				'word/header1.xml': `<w:hdr xmlns:w="${NS.w}">${paragraph(
					complex(code(' MERGEFIELD TableStart:Who '), '') +
						simple('MERGEFIELD Name') +
						complex(code(' MERGEFIELD TableEnd:Who '), ''),
				)}</w:hdr>`,
				'word/_rels/document.xml.rels': `<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/header" Target="header1.xml"/></Relationships>`,
			},
		);
		const merged = await mergeFields(
			loaded.model,
			{ values: {}, regions: { Who: [{ Name: 'Ada' }, { Name: 'Bea' }] } },
			{ unlink: true },
		);
		expect(merged.regions).toEqual([{ name: 'Who', kind: 'paragraphs', story: 'header', rows: 2 }]);
		const reopened = await loadDocx(await loaded.save(merged.model));
		expect(text(reopened.model.sections?.[0]?.headers?.default?.blocks ?? [])).toEqual([
			'Ada',
			'Bea',
		]);
	});

	it('keeps the input model unchanged', async () => {
		const loaded = await template(paragraph(simple('MERGEFIELD Name')));
		const before = JSON.stringify(loaded.model);
		await mergeFields(loaded.model, { values: { Name: 'Ada' } }, { unlink: true });
		expect(JSON.stringify(loaded.model)).toBe(before);
	});
});
