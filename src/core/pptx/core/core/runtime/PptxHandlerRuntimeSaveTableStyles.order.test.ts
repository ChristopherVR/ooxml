import { DOMParser } from '@xmldom/xmldom';
import { XMLBuilder, XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../PptxHandler';
import type { PptxTableCellStyle, XmlObject } from '../../types';
import { PptxHandlerRuntime } from './PptxHandlerRuntimeImplementation';

class TableStyleRuntime extends PptxHandlerRuntime {
	public writeText(cell: XmlObject, text: string): void {
		this.writeTableCellText(cell, text);
	}

	public writeStyle(cell: XmlObject, style: PptxTableCellStyle): void {
		this.writeTableCellStyle(cell, style);
	}
}

const runtime = new TableStyleRuntime();
const parser = new XMLParser({ ignoreAttributes: false, parseTagValue: false });
const builder = new XMLBuilder({ ignoreAttributes: false });
const alignments = ['left', 'center', 'right', 'justify'] as const;
const tokens = { left: 'l', center: 'ctr', right: 'r', justify: 'just' };

/** Check every direct paragraph child, not nested field paragraph properties. */
function expectParagraphOrder(xml: string, count: number): void {
	const doc = new DOMParser().parseFromString(
		`<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${xml}</root>`,
		'text/xml',
	);
	const paragraphs = doc.getElementsByTagName('a:p');
	expect(paragraphs.length).toBe(count);
	for (let index = 0; index < paragraphs.length; index++) {
		const names: string[] = [];
		for (let child = paragraphs[index].firstChild; child; child = child.nextSibling) {
			if (child.nodeType === 1) {
				names.push(child.nodeName);
			}
		}
		expect(names.filter((name) => name === 'a:pPr').length).toBeLessThanOrEqual(1);
		expect(names.filter((name) => name === 'a:endParaRPr').length).toBeLessThanOrEqual(1);
		const ranks = names.map((name) => {
			expect(['a:pPr', 'a:r', 'a:br', 'a:fld', 'a:endParaRPr']).toContain(name);
			return name === 'a:pPr' ? 0 : name === 'a:endParaRPr' ? 2 : 1;
		});
		expect(ranks).toStrictEqual([...ranks].sort());
	}
}

function firstParagraph(cell: XmlObject): XmlObject {
	const paragraphs = (cell['a:txBody'] as XmlObject)['a:p'];
	return (Array.isArray(paragraphs) ? paragraphs[0] : paragraphs) as XmlObject;
}

describe('table save paragraph child order', () => {
	it.each(alignments)('saves typed %s alignment before runs in every table cell', async (align) => {
		const { handler, data, createSlide } = await PptxHandler.createBlank();
		const slide = createSlide();
		slide.addElement({
			type: 'table',
			id: 'tbl1',
			x: 60,
			y: 60,
			width: 640,
			height: 160,
			tableData: {
				columnWidths: [0.5, 0.5],
				rows: [
					{
						height: 80,
						cells: ['A\nB', ''].map((text) => ({
							text,
							style: { fontSize: 14, align, vAlign: 'middle' },
						})),
					},
				],
				bandedRows: false,
				bandedColumns: false,
			},
		});
		data.slides.push(slide.build());
		const bytes = await handler.save(data.slides);
		const zip = await JSZip.loadAsync(bytes);
		const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const cells = [...xml.matchAll(/<a:tc\b[^>]*>([\s\S]*?)<\/a:tc>/g)];
		expect(cells).toHaveLength(2);
		for (const [index, cell] of cells.entries()) {
			expectParagraphOrder(cell[1], index === 0 ? 2 : 1);
			expect(cell[1]).toContain(`algn="${tokens[align]}"`);
		}
	});

	it.each(alignments)(
		'merges %s alignment with existing indentation, bullets and spacing',
		(align) => {
			const cell = parser.parse(
				'<a:tc><a:txBody><a:bodyPr/><a:p><a:pPr marL="91440" indent="-45720">' +
					'<a:spcAft><a:spcPts val="600"/></a:spcAft><a:buChar char="*"/></a:pPr>' +
					'<a:r><a:t>First</a:t></a:r><a:endParaRPr sz="1800"/></a:p>' +
					'<a:p><a:pPr algn="r"/><a:r><a:t>Second</a:t></a:r><a:endParaRPr/></a:p>' +
					'</a:txBody></a:tc>',
			)['a:tc'] as XmlObject;
			const paragraph = firstParagraph(cell);
			const original = paragraph['a:pPr'];
			runtime.writeStyle(cell, { align });
			runtime.writeStyle(cell, { align });
			expect(paragraph['a:pPr']).toBe(original);
			expect(paragraph['a:pPr']).toStrictEqual({
				'@_marL': '91440',
				'@_indent': '-45720',
				'@_algn': tokens[align],
				'a:spcAft': { 'a:spcPts': { '@_val': '600' } },
				'a:buChar': { '@_char': '*' },
			});
			const paragraphs = (cell['a:txBody'] as XmlObject)['a:p'] as XmlObject[];
			expect(paragraphs[1]['a:pPr']).toStrictEqual({ '@_algn': 'r' });
			expectParagraphOrder(builder.build({ 'a:tc': cell }), 2);
		},
	);

	it.each([
		'<a:p><a:r><a:t>Text</a:t></a:r><a:endParaRPr sz="1800"/></a:p>',
		'<a:p><a:pPr/><a:r><a:t>Text</a:t></a:r><a:endParaRPr/></a:p>',
		'<a:p><a:endParaRPr sz="1800"/></a:p>',
		'<a:p/>',
	])('aligns empty or unstyled parsed paragraphs without changing their content: %s', (xml) => {
		const cell = parser.parse(`<a:tc><a:txBody><a:bodyPr/>${xml}</a:txBody></a:tc>`)[
			'a:tc'
		] as XmlObject;
		const originalParagraph = structuredClone(firstParagraph(cell));
		runtime.writeStyle(cell, { align: 'center' });
		const paragraph = firstParagraph(cell);
		expect(paragraph['a:pPr']).toStrictEqual({ '@_algn': 'ctr' });
		expect(paragraph['a:r']).toStrictEqual(originalParagraph['a:r']);
		expect(paragraph['a:endParaRPr']).toStrictEqual(originalParagraph['a:endParaRPr']);
		expectParagraphOrder(builder.build({ 'a:tc': cell }), 1);
	});

	it('retains paragraph-end properties when rebuilding a multiline cell before alignment', () => {
		const cell: XmlObject = {
			'a:txBody': {
				'a:bodyPr': {},
				'a:p': { 'a:r': { 'a:t': 'Old' }, 'a:endParaRPr': { '@_sz': '1800' } },
			},
		};
		runtime.writeText(cell, 'First\n\nLast');
		runtime.writeStyle(cell, { align: 'justify' });
		const xml = builder.build({ 'a:tc': cell }) as string;
		expectParagraphOrder(xml, 3);
		expect([...xml.matchAll(/<a:endParaRPr sz="1800"/g)]).toHaveLength(3);
	});
});
