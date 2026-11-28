// Issue #399: when the last row of a vertical merge breaks across pages, Word keeps the first row at
// its own height and the merged content flows into the second row, as it does when the table fits.
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from '../index';
import { layoutDocumentModel } from './layout';
import type { LayoutTableBox } from './result';
import type { TextMeasurer } from './measure';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const measurer: TextMeasurer = { widthOf: (text) => text.length * 6, lineHeightOf: () => 20 };
const p = (text: string) =>
	`<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
const lines = (label: string, count: number) =>
	Array.from({ length: count }, (_, i) => p(`${label} ${i + 1}`)).join('');
const tc = (body: string, props = '') =>
	`<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${props}</w:tcPr>${body}</w:tc>`;

async function tables(fillers: number): Promise<LayoutTableBox[]> {
	const body = `<w:tr>${tc(p('Row 1'))}${tc(lines('Merged', 12), '<w:vMerge w:val="restart"/>')}${tc(p('Short'))}</w:tr>
<w:tr>${tc(p('Row 2'))}${tc('<w:p/>', '<w:vMerge/>')}${tc(lines('Tall', 30))}</w:tr>`;
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body>${lines('Filler', fillers)}<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${body}</w:tbl><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	return layoutDocumentModel(model, measurer).pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.filter((block): block is LayoutTableBox => block.kind === 'table'),
		),
	);
}

const lineCount = (row: LayoutTableBox['rows'][number], cell: number) =>
	row.cells[cell]!.reduce((sum, box) => sum + box.lines.length, 0);

describe('a vertical merge ending in a row that breaks across pages (#399)', () => {
	it('lays the table out the same when it fits (control)', async () => {
		const [table] = await tables(0);
		expect(table!.rows[0]!.heightPx).toBe(20);
		expect(table!.rows[1]!.yPx).toBe(20);
	});

	it('keeps the first row at its own height when the last row breaks', async () => {
		const [first, second] = await tables(40);
		expect(second).toBeDefined();
		expect(first!.rows[0]!.heightPx).toBe(20);
		expect(first!.rows[1]!.yPx).toBe(20);
		expect(lineCount(first!.rows[1]!, 2)).toBeGreaterThan(3);
	});
});
