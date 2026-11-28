// Issue #402: a vertically merged cell with vAlign=center is centred over the whole merge, also when
// the merge starts at the top of a page (the table's earlier rows end the previous page).
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
const lines = (label: string, count: number) =>
	Array.from({ length: count }, (_, i) => p(`${label} ${i + 1}`)).join('');
const tc = (body: string, props = '') =>
	`<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${props}</w:tcPr>${body}</w:tc>`;

async function tables(fillers: number): Promise<LayoutTableBox[]> {
	const merged = '<w:vMerge w:val="restart"/><w:vAlign w:val="center"/>';
	const body = `<w:tr>${tc(lines('Row one,', 6))}${tc(p('b'))}${tc(p('c'))}</w:tr>
<w:tr>${tc(lines('Row two,', 6))}${tc(p('Merged and centered'), merged)}${tc(p('c'))}</w:tr>
<w:tr>${tc(lines('Row three,', 6))}${tc('<w:p/>', '<w:vMerge/><w:vAlign w:val="center"/>')}${tc(p('c'))}</w:tr>`;
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body>${lines('Filler', fillers)}<w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>${body}</w:tbl><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	return layoutDocumentModel(model, measurer).pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.filter((block): block is LayoutTableBox => block.kind === 'table'),
		),
	);
}

function centreOffset(table: LayoutTableBox): {
	offset: number;
	expected: number;
	mergeRows: number | undefined;
} {
	const row = table.rows.find((item) => item.geometry?.[1]?.mergeRows);
	const geometry = row!.geometry![1]!;
	const total = table.rows.slice(table.rows.indexOf(row!)).reduce((sum, r) => sum + r.heightPx, 0);
	return {
		offset: cellVerticalOffset(geometry, row!.heightPx),
		expected: (total - geometry.contentHeightPx) / 2,
		mergeRows: geometry.mergeRows,
	};
}

describe('a centred vertical merge at the top of a page (#402)', () => {
	it('centres over both rows mid-page (control)', async () => {
		const [table] = await tables(2);
		const { offset, expected } = centreOffset(table!);
		expect(offset).toBeCloseTo(expected, 3);
	});

	it('centres over both rows when the merge starts the page', async () => {
		const found = await tables(40);
		const table = found.at(-1)!;
		expect(table.rows).toHaveLength(2);
		const { offset, expected, mergeRows } = centreOffset(table);
		expect(mergeRows).toBe(2);
		expect(offset).toBeCloseTo(expected, 3);
	});
});
