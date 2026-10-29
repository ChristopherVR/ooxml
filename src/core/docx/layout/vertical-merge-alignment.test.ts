// Word aligns a vertically merged cell's text across every row the merge covers (a centred cell is
// centred in the whole merged area). The restart row's geometry carries the merged height.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import { cellVerticalOffset, type LayoutTableBox } from './result';
import type { TextMeasurer } from './measure';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer: TextMeasurer = { widthOf: (text) => text.length * 6, lineHeightOf: () => 20 };
const p = (text: string) =>
	`<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const tc = (body: string, props = '') => `<w:tc><w:tcPr>${props}</w:tcPr>${body}</w:tc>`;
const tr = (cells: string, height = 600) =>
	`<w:tr><w:trPr><w:trHeight w:val="${height}" w:hRule="exact"/></w:trPr>${cells}</w:tr>`;

async function firstTable(align: string, rows = 3): Promise<LayoutTableBox> {
	const body = [
		tr(tc(p('Merged'), `<w:vMerge w:val="restart"/><w:vAlign w:val="${align}"/>`) + tc(p('a'))),
		...Array.from({ length: rows - 1 }, (_, i) => tr(tc('<w:p/>', '<w:vMerge/>') + tc(p(`r${i}`)))),
	].join('');
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${body}</w:tbl><w:sectPr/></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const box = layoutDocumentModel(model, measurer)
		.pages[0]?.columns.flatMap((column) => column.blocks)
		.find((block): block is LayoutTableBox => block.kind === 'table');
	if (!box) throw new Error('No table box');
	return box;
}

describe('vertically merged cells on Print Layout pages', () => {
	it('centres the text across the whole merged area', async () => {
		const box = await firstTable('center');
		const first = box.rows[0]!;
		const geometry = first.geometry![0]!;
		const total = box.rows.reduce((sum, row) => sum + row.heightPx, 0);
		expect(geometry.mergeRows).toBe(3);
		expect(geometry.alignmentHeightPx).toBeCloseTo(total, 5);
		const offset = cellVerticalOffset(geometry, first.heightPx);
		expect(offset).toBeCloseTo((total - geometry.contentHeightPx) / 2, 5);
		expect(offset).toBeGreaterThan(first.heightPx / 2 - 10);
	});

	it('puts bottom-aligned text at the bottom of the last merged row', async () => {
		const box = await firstTable('bottom', 2);
		const first = box.rows[0]!;
		const total = box.rows.reduce((sum, row) => sum + row.heightPx, 0);
		const geometry = first.geometry![0]!;
		expect(cellVerticalOffset(geometry, first.heightPx)).toBeCloseTo(
			total - geometry.contentHeightPx,
			5,
		);
	});

	it('leaves top-aligned text in the first row, and cells outside a merge alone', async () => {
		const box = await firstTable('top');
		const geometry = box.rows[0]!.geometry![0]!;
		expect(geometry.alignmentHeightPx).toBeUndefined();
		expect(cellVerticalOffset(geometry, box.rows[0]!.heightPx)).toBe(0);
		expect(box.rows[0]!.geometry![1]!.mergeRows).toBeUndefined();
	});
});
