import { XMLParser } from 'fast-xml-parser';
import { describe, expect, it } from 'vitest';

import type { PptxTableCellParagraph, PptxTableCellStyle, XmlObject } from '../../types';
import { PptxHandlerRuntime } from './PptxHandlerRuntimeImplementation';

class TableStyleRuntime extends PptxHandlerRuntime {
	public writeStyle(
		cell: XmlObject,
		style: PptxTableCellStyle,
		paragraphs?: PptxTableCellParagraph[],
	): void {
		this.writeTableCellStyle(cell, style, paragraphs);
	}
}

const runtime = new TableStyleRuntime();
const parser = new XMLParser({ ignoreAttributes: false, parseTagValue: false });

function parseCell(paragraphs: string): XmlObject {
	return parser.parse(`<a:tc><a:txBody><a:bodyPr/>${paragraphs}</a:txBody></a:tc>`)[
		'a:tc'
	] as XmlObject;
}

function alignments(cell: XmlObject): unknown[] {
	const paragraphs = (cell['a:txBody'] as XmlObject)['a:p'] as XmlObject[];
	return paragraphs.map((paragraph) => (paragraph['a:pPr'] as XmlObject | undefined)?.['@_algn']);
}

describe('table cell paragraph alignment on save', () => {
	it("writes each paragraph's own alignment", () => {
		const cell = parseCell(
			'<a:p><a:pPr algn="r"/><a:r><a:t>One</a:t></a:r></a:p>' +
				'<a:p><a:pPr marL="0" algn="ctr"/><a:r><a:t>Two</a:t></a:r></a:p>' +
				'<a:p><a:r><a:t>Three</a:t></a:r></a:p>',
		);
		runtime.writeStyle(cell, { align: 'left' }, [
			{ align: 'left' },
			{ align: 'left', paragraphMarginLeft: 0 },
			{ align: 'left' },
		]);
		expect(alignments(cell)).toStrictEqual(['l', 'l', 'l']);
		const paragraphs = (cell['a:txBody'] as XmlObject)['a:p'] as XmlObject[];
		expect(paragraphs[1]['a:pPr']).toStrictEqual({ '@_marL': '0', '@_algn': 'l' });
		expect(Object.keys(paragraphs[2])).toStrictEqual(['a:pPr', 'a:r']);
	});

	it('leaves a cell whose paragraphs align differently as it was', () => {
		const cell = parseCell(
			'<a:p><a:pPr algn="r"/><a:r><a:t>One</a:t></a:r></a:p>' +
				'<a:p><a:pPr algn="ctr"/><a:r><a:t>Two</a:t></a:r></a:p>' +
				'<a:p><a:pPr algn="dist"/><a:r><a:t>Three</a:t></a:r></a:p>' +
				'<a:p><a:pPr algn="justify"/><a:r><a:t>Four</a:t></a:r></a:p>' +
				'<a:p><a:r><a:t>Five</a:t></a:r></a:p>' +
				'<a:p><a:endParaRPr sz="1100"/></a:p>',
		);
		const original = structuredClone(cell['a:txBody']);
		runtime.writeStyle(cell, { align: 'right' }, [
			{ align: 'right' },
			{ align: 'center' },
			{ align: 'dist' },
			{ align: 'justify' },
			{},
			{},
		]);
		expect(cell['a:txBody']).toStrictEqual(original);
	});

	it('writes no paragraph properties for an alignment it has no token for', () => {
		const cell = parseCell(
			'<a:p><a:r><a:t>One</a:t></a:r></a:p><a:p><a:r><a:t>Two</a:t></a:r></a:p>',
		);
		const original = structuredClone(cell['a:txBody']);
		runtime.writeStyle(cell, {}, [{}, { align: 'middle' as PptxTableCellParagraph['align'] }]);
		expect(cell['a:txBody']).toStrictEqual(original);
	});
});
