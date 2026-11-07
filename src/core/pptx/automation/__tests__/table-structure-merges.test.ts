import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { NS } from '../../../xml/namespaces';
import { buildXml, parseXml } from '../../../xml/xml';
import { PptxHandler } from '../../index';
import { pptxTools } from '../catalogue/index';

const title = 'Merged heading';

async function inputDeck(useBooleanWords: boolean): Promise<Uint8Array> {
	const { handler, createSlide } = await PptxHandler.createBlank({ initialSlideCount: 0 });
	try {
		const slide = createSlide('Blank')
			.addTable({
				columnWidths: [0.2, 0.3, 0.5],
				rows: [
					{ cells: [{ text: title }, { text: '' }, { text: 'Top right' }] },
					{ cells: [{ text: '' }, { text: '' }, { text: 'Middle right' }] },
					{ cells: [{ text: 'Bottom left' }, { text: 'Bottom middle' }, { text: 'Independent' }] },
				],
			})
			.build();
		const table = slide.elements.find((element) => element.type === 'table');
		if (!table?.tableData) throw new Error('Missing test table');
		const [top, second] = table.tableData.rows;
		Object.assign(top.cells[0], { rowSpan: 2, gridSpan: 2 });
		Object.assign(top.cells[1], { rowSpan: 2, hMerge: true });
		Object.assign(second.cells[0], { gridSpan: 2, vMerge: true });
		Object.assign(second.cells[1], { hMerge: true, vMerge: true });

		const zip = await JSZip.loadAsync(await handler.save([slide]));
		const part = 'ppt/slides/slide1.xml';
		const doc = parseXml(await zip.file(part)!.async('string'));
		const tableCells = Array.from(doc.getElementsByTagNameNS(NS.a, 'tc'));
		const paragraph = tableCells[0].getElementsByTagNameNS(NS.a, 'p')[0];
		const rich = parseXml(
			`<a:p xmlns:a="${NS.a}"><a:r><a:rPr b="1"/><a:t>Merged </a:t></a:r>` +
				'<a:r><a:rPr i="1"/><a:t>heading</a:t></a:r></a:p>',
		);
		paragraph.parentNode!.replaceChild(doc.importNode(rich.documentElement!, true), paragraph);
		if (useBooleanWords) {
			for (const cell of tableCells) {
				for (const attr of ['hMerge', 'vMerge']) {
					if (cell.getAttribute(attr) === '1') cell.setAttribute(attr, 'true');
				}
			}
		}
		zip.file(part, buildXml(doc));
		return zip.generateAsync({ type: 'uint8array' });
	} finally {
		handler.dispose();
	}
}

async function readTable(bytes: Uint8Array) {
	const handler = new PptxHandler();
	try {
		const data = await handler.load(new Uint8Array(bytes).buffer);
		const table = data.slides[0].elements.find((element) => element.type === 'table');
		if (!table?.tableData) throw new Error('Missing saved table');
		return { id: table.id, ...table.tableData };
	} finally {
		handler.dispose();
	}
}

describe('manage_table_structure with merged cells', () => {
	it.each([
		{ action: 'deleteRow', referenceIndex: 0, merge: [1, 2], size: [2, 3] },
		{ action: 'deleteRow', referenceIndex: 1, merge: [1, 2], size: [2, 3] },
		{ action: 'deleteColumn', referenceIndex: 0, merge: [2, 1], size: [3, 2] },
		{ action: 'deleteColumn', referenceIndex: 1, merge: [2, 1], size: [3, 2] },
		{ action: 'insertRow', position: 1, merge: [3, 2], size: [4, 3] },
		{ action: 'insertColumn', position: 1, merge: [2, 3], size: [3, 4] },
	] as const)(
		'preserves the merged rectangle and rich title after $action ($referenceIndex$position)',
		async ({ merge: [height, width], size: [rowCount, columnCount], ...edit }) => {
			const input = await inputDeck(edit.action === 'insertRow');
			const original = await readTable(input);
			const tool = pptxTools.find((candidate) => candidate.name === 'manage_table_structure');
			if (tool?.kind !== 'edit') throw new Error('Missing table structure tool');
			const result = await tool.run(input, {
				slideIndex: 0,
				elementId: original.id,
				...edit,
				cellTexts: ['', '', 'Inserted'],
			});
			if (!result.bytes) throw new Error('The table edit did not save');
			expect(result).toMatchObject({ rowCount, columnCount });

			const saved = await readTable(result.bytes);
			expect(saved.rows.map((row) => row.cells.length)).toEqual(Array(rowCount).fill(columnCount));
			expect(saved.columnWidths).toHaveLength(columnCount);
			expect(saved.rows[0].cells[0].text).toBe(title);
			expect(saved.rows[0].cells[0].textRuns).toEqual(original.rows[0].cells[0].textRuns);
			expect(saved.rows[0].cells[0].textRuns).toMatchObject([
				{ text: 'Merged ', bold: true },
				{ text: 'heading', italic: true },
			]);
			for (const [r, row] of saved.rows.entries()) {
				for (const [c, cell] of row.cells.entries()) {
					const merged = r < height && c < width;
					expect({
						rowSpan: cell.rowSpan ?? 1,
						gridSpan: cell.gridSpan ?? 1,
						hMerge: cell.hMerge,
						vMerge: cell.vMerge,
					}).toEqual({
						rowSpan: merged && r === 0 ? height : 1,
						gridSpan: merged && c === 0 ? width : 1,
						hMerge: merged && c > 0,
						vMerge: merged && r > 0,
					});
				}
			}
			expect(saved.rows.at(-1)?.cells.at(-1)?.text).toBe('Independent');
			if (edit.action === 'insertRow') expect(saved.rows[1].cells[2].text).toBe('Inserted');
			if (edit.action === 'insertColumn') expect(saved.rows[2].cells[1].text).toBe('Inserted');
			if (edit.action === 'deleteColumn' && edit.referenceIndex === 0) {
				expect(saved.columnWidths[0]).toBeCloseTo(0.4);
			}
		},
	);
});
