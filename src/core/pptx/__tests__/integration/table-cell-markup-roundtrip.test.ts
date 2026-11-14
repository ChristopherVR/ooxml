import JSZip from 'jszip';
import { PresentationBuilder } from 'ooxml-core/pptx';
import type { PptxSlide, TablePptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import { editSegmentsText, segmentsPlainText } from '../../editor/render/segments-text-edit';
import {
	setTableElementCellSegments,
	tableCellSegments,
} from '../../editor/render/table-cell-rich-text';
import { tableCellSourceXml } from '../../editor/render/table-cell-source';

/**
 * One cell holding everything an edit used to drop: a hyperlink run (with its
 * relationship id), a slide-number field, a date field and a scheme colour
 * with a luminance transform, plus a plain run the edit actually changes.
 */
const cell =
	'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/>' +
	'<a:r><a:rPr lang="en-GB" dirty="0"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:rPr><a:t>Themed </a:t></a:r>' +
	'<a:r><a:rPr lang="en-US" u="sng"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>link</a:t></a:r>' +
	'<a:r><a:rPr lang="en-US"/><a:t> page </a:t></a:r>' +
	'<a:fld id="{11111111-2222-3333-4444-555555555555}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld>' +
	'<a:r><a:rPr lang="en-US"/><a:t> on </a:t></a:r>' +
	'<a:fld id="{66666666-7777-8888-9999-000000000000}" type="datetime1"><a:rPr lang="en-US"/><a:t>1/2/2026</a:t></a:fld>' +
	'<a:endParaRPr lang="en-US" sz="1200"/></a:p></a:txBody><a:tcPr/></a:tc>';

async function createDeck(): Promise<ArrayBuffer> {
	const { handler, data } = await PresentationBuilder.create({ initialSlideCount: 1 });
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const table =
		'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="30" name="Table 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
		'<p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
		`<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="4000000"/></a:tblGrid><a:tr h="500000">${cell}</a:tr></a:tbl>` +
		'</a:graphicData></a:graphic></p:graphicFrame>';
	const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	zip.file('ppt/slides/slide1.xml', slideXml.replace('</p:spTree>', `${table}</p:spTree>`));
	const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
	zip.file(
		'ppt/slides/_rels/slide1.xml.rels',
		rels.replace(
			'</Relationships>',
			'<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>',
		),
	);
	return zip.generateAsync({ type: 'arraybuffer' });
}

/** A cell's text in authored order (its flat `text` lists fields last). */
function runText(cell: { textRuns?: Array<{ text: string }> }): string {
	return (cell.textRuns ?? []).map((run) => run.text).join('');
}

function tableOf(slides: readonly PptxSlide[]): TablePptxElement {
	const table = slides[0]!.elements.find((element) => element.type === 'table');
	if (table?.type !== 'table') {
		throw new Error('Expected a table');
	}
	return table;
}

async function editAndSave(edit: (text: string) => { text: string; caret: number }) {
	const handler = new PptxHandler();
	const data = await handler.load(await createDeck());
	const table = tableOf(data.slides);
	const segments = tableCellSegments(
		table.tableData!.rows[0]!.cells[0]!,
		tableCellSourceXml(table, 0, 0),
	);
	const { text, caret } = edit(segmentsPlainText(segments));
	const next = setTableElementCellSegments(table, 0, 0, editSegmentsText(segments, text, caret));
	const slides = data.slides.map((slide, index) =>
		index === 0
			? { ...slide, elements: slide.elements.map((el) => (el.id === table.id ? next : el)) }
			: slide,
	);
	const saved = await handler.save(slides);
	const zip = await JSZip.loadAsync(saved);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
	const reopened = await new PptxHandler().load(
		saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
	);
	return { xml, rels, cell: tableOf(reopened.slides).tableData!.rows[0]!.cells[0]! };
}

describe('editing a table cell keeps the markup of its runs', () => {
	it('keeps a hyperlink, both fields and a scheme colour when a word is edited', async () => {
		const {
			xml,
			rels,
			cell: saved,
		} = await editAndSave((text) => {
			const at = text.indexOf(' page ') + 1;
			return { text: `${text.slice(0, at)}sheet${text.slice(at + 4)}`, caret: at + 5 };
		});
		expect(runText(saved)).toBe('Themed link sheet 1 on 1/2/2026');
		expect(xml).toContain('<a:hlinkClick r:id="rIdLink">');
		expect(rels).toContain('Id="rIdLink"');
		expect(xml).toContain('<a:schemeClr val="accent1"><a:lumMod val="75000">');
		expect(xml).toContain('lang="en-GB" dirty="0"');
		expect(xml).toMatch(/<a:fld id="\{11111111-2222-3333-4444-555555555555\}" type="slidenum">/);
		expect(xml).toMatch(/<a:fld id="\{66666666-7777-8888-9999-000000000000\}" type="datetime1">/);
		expect(xml).toContain('<a:pPr algn="ctr">');
		expect(xml).toContain('<a:endParaRPr lang="en-US" sz="1200">');
		// The field still sits between its runs, in authored order.
		expect(xml.indexOf('sheet')).toBeLessThan(xml.indexOf('type="slidenum"'));
		expect(xml.indexOf('type="slidenum"')).toBeLessThan(xml.indexOf(' on '));
		const runs = saved.textRuns!;
		expect(runs.filter((run) => run.isField)).toHaveLength(2);
		expect(runs.find((run) => run.text === 'link')?.underline).toBe(true);
	});

	it('overrides only what changed on a run: bold on a theme-coloured run keeps its colour', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await createDeck());
		const table = tableOf(data.slides);
		const segments = tableCellSegments(
			table.tableData!.rows[0]!.cells[0]!,
			tableCellSourceXml(table, 0, 0),
		).map((segment) =>
			segment.text === 'Themed '
				? { ...segment, style: { ...segment.style, bold: true } }
				: segment,
		);
		const next = setTableElementCellSegments(table, 0, 0, segments);
		const saved = await handler.save(
			data.slides.map((slide, index) =>
				index === 0
					? { ...slide, elements: slide.elements.map((el) => (el.id === table.id ? next : el)) }
					: slide,
			),
		);
		const xml = await (await JSZip.loadAsync(saved)).file('ppt/slides/slide1.xml')!.async('string');
		expect(xml).toContain(
			'<a:rPr lang="en-GB" dirty="0" b="1"><a:solidFill><a:schemeClr val="accent1">',
		);
		expect(xml).toContain('<a:hlinkClick r:id="rIdLink">');
	});

	it('turns a field into plain text only once its own text is edited', async () => {
		const { xml, cell: saved } = await editAndSave((text) => {
			const at = text.indexOf('1/2/2026');
			return { text: `${text.slice(0, at)}1/2/2027`, caret: at + 8 };
		});
		expect(runText(saved)).toBe('Themed link page 1 on 1/2/2027');
		expect(xml).toContain('type="slidenum"');
		expect(xml).not.toContain('type="datetime1"');
		expect(saved.textRuns!.filter((run) => run.isField)).toHaveLength(1);
	});

	it('gives text typed after a hyperlink run that run, as PowerPoint does', async () => {
		const { xml } = await editAndSave((text) => {
			const at = text.indexOf('link') + 4;
			return { text: `${text.slice(0, at)}s${text.slice(at)}`, caret: at + 1 };
		});
		expect(xml).toContain('<a:hlinkClick r:id="rIdLink"></a:hlinkClick></a:rPr><a:t>links</a:t>');
	});
});
