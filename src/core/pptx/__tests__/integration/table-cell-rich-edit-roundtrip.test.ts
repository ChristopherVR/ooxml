import JSZip from 'jszip';
import { PresentationBuilder } from 'ooxml-core/pptx';
import type { PptxSlide, TablePptxElement, TextSegment } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import { editSegmentsText } from '../../editor/render/segments-text-edit';
import {
	setTableElementCellSegments,
	tableCellSegments,
} from '../../editor/render/table-cell-rich-text';
import { applyTableCellsTextStyle } from '../../editor/render/table-cells-text-style';
import { fillTableCells } from '../../editor/render/table-grid-fill';

const richCell =
	'<a:txBody><a:bodyPr/><a:lstStyle/>' +
	'<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400"/><a:t>Revenue </a:t></a:r>' +
	'<a:r><a:rPr lang="en-US" sz="1800" b="1"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:latin typeface="Georgia"/></a:rPr><a:t>grew 42%</a:t></a:r></a:p>' +
	'<a:p><a:r><a:rPr lang="en-US" i="1"/><a:t>Second line</a:t></a:r></a:p>' +
	'</a:txBody><a:tcPr/>';
const plainCell =
	'<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Plain</a:t></a:r></a:p></a:txBody><a:tcPr/>';

async function createTableDeck(): Promise<ArrayBuffer> {
	const { handler, data } = await PresentationBuilder.create({ initialSlideCount: 1 });
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const table =
		'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="30" name="Table 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
		'<p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="1000000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
		'<a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>' +
		`<a:tr h="500000"><a:tc>${richCell}</a:tc><a:tc>${plainCell}</a:tc></a:tr>` +
		'</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
	const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	zip.file('ppt/slides/slide1.xml', slideXml.replace('</p:spTree>', `${table}</p:spTree>`));
	return zip.generateAsync({ type: 'arraybuffer' });
}

function tableOf(slides: readonly PptxSlide[]): TablePptxElement {
	const table = slides[0]!.elements.find((element) => element.type === 'table');
	if (table?.type !== 'table') {
		throw new Error('Expected a table');
	}
	return table;
}

function withTable(slides: PptxSlide[], table: TablePptxElement): PptxSlide[] {
	return slides.map((slide, index) =>
		index === 0
			? {
					...slide,
					elements: slide.elements.map((element) => (element.id === table.id ? table : element)),
				}
			: slide,
	);
}

async function saveAndReopen(handler: PptxHandler, slides: PptxSlide[]) {
	const saved = await handler.save(slides);
	const xml = await (await JSZip.loadAsync(saved)).file('ppt/slides/slide1.xml')!.async('string');
	const reopened = await new PptxHandler().load(
		saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength) as ArrayBuffer,
	);
	return { xml, table: tableOf(reopened.slides) };
}

describe('rich table cell editing round trip', () => {
	it('keeps every run format when one word of a rich cell is edited', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await createTableDeck());
		const table = tableOf(data.slides);
		const cell = table.tableData!.rows[0]!.cells[0]!;
		const segments = tableCellSegments(cell);
		expect(segments.filter((s) => !s.isParagraphBreak).map((s) => s.text)).toEqual([
			'Revenue ',
			'grew 42%',
			'Second line',
		]);
		// What the editor commits: "Revenue" retyped as "Income", caret after it.
		const edited = editSegmentsText(segments, 'Income grew 42%\nSecond line', 6);
		const next = setTableElementCellSegments(table, 0, 0, edited);
		const { xml, table: reopened } = await saveAndReopen(handler, withTable(data.slides, next));

		const runs = reopened.tableData!.rows[0]!.cells[0]!.textRuns!;
		const income = runs.find((run) => run.text.startsWith('Income'));
		const grew = runs.find((run) => run.text === 'grew 42%');
		const second = runs.find((run) => run.text === 'Second line');
		expect(income?.bold).toBeUndefined();
		expect(income?.fontSize).toBe(14);
		expect(grew).toMatchObject({ bold: true, fontSize: 18, fontFamily: 'Georgia' });
		expect(grew?.color?.toUpperCase()).toContain('C00000');
		expect(second?.italic).toBe(true);
		expect(runs.filter((run) => run.isParagraphBreak)).toHaveLength(1);
		expect(xml).toContain('algn="ctr"');
		// The untouched cell keeps its markup.
		expect(reopened.tableData!.rows[0]!.cells[1]!.text).toBe('Plain');
	});

	it('writes formatting applied inside the cell editor as its own run', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await createTableDeck());
		const table = tableOf(data.slides);
		const segments: TextSegment[] = [
			{ text: 'Pl', style: { bold: true } },
			{ text: 'ain', style: { underline: true } },
		];
		const next = setTableElementCellSegments(table, 0, 1, segments);
		const { table: reopened } = await saveAndReopen(handler, withTable(data.slides, next));
		const runs = reopened.tableData!.rows[0]!.cells[1]!.textRuns!;
		expect(runs.map((run) => [run.text, Boolean(run.bold), Boolean(run.underline)])).toEqual([
			['Pl', true, false],
			['ain', false, true],
		]);
	});

	it('leaves an unchanged cell alone', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await createTableDeck());
		const table = tableOf(data.slides);
		const cell = table.tableData!.rows[0]!.cells[0]!;
		expect(setTableElementCellSegments(table, 0, 0, tableCellSegments(cell))).toBe(table);
	});

	it('applies range formatting to every run of every selected cell', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await createTableDeck());
		const table = tableOf(data.slides);
		const next = applyTableCellsTextStyle(table, undefined, { italic: true, fontSize: 20 });
		expect(
			next.tableData!.rows[0]!.cells[0]!.textRuns!.every((r) => r.isParagraphBreak || r.italic),
		).toBe(true);
		const { table: reopened } = await saveAndReopen(handler, withTable(data.slides, next));
		const rich = reopened.tableData!.rows[0]!.cells[0]!.textRuns!.filter(
			(r) => !r.isParagraphBreak,
		);
		expect(rich.every((run) => run.italic && run.fontSize === 20)).toBe(true);
		// Per-run formatting the update did not name survives.
		expect(rich.find((run) => run.text === 'grew 42%')).toMatchObject({ bold: true });
		const plain = reopened.tableData!.rows[0]!.cells[1]!.textRuns!;
		expect(plain[0]).toMatchObject({ text: 'Plain', italic: true, fontSize: 20 });
	});

	it('fills pasted values from the anchor, growing the table', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(await createTableDeck());
		const table = tableOf(data.slides);
		const next = fillTableCells(table, {
			addRows: 1,
			addColumns: 1,
			writes: [
				{ row: 0, col: 1, text: 'b1' },
				{ row: 0, col: 2, text: 'c1' },
				{ row: 1, col: 1, text: 'b2' },
				{ row: 1, col: 2, text: 'c2' },
			],
		});
		const { table: reopened } = await saveAndReopen(handler, withTable(data.slides, next));
		const rows = reopened.tableData!.rows.map((row) => row.cells.map((cell) => cell.text));
		expect(rows).toEqual([
			['Revenue grew 42%\nSecond line', 'b1', 'c1'],
			['', 'b2', 'c2'],
		]);
	});
});
