// The text of a vertically merged cell grows the LAST row it spans (measured in Word, see
// `word-table.test.ts`); a merge too tall for the room left keeps its text in the first row, where
// the row splitter can cut it across pages.
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
const tc = (body: string, props = '') => `<w:tc><w:tcPr>${props}</w:tcPr>${body}</w:tc>`;
const tr = (cells: string, trPr = '') => `<w:tr><w:trPr>${trPr}</w:trPr>${cells}</w:tr>`;

async function tables(mergedLines: number): Promise<LayoutTableBox[]> {
	const merged = Array.from({ length: mergedLines }, (_, i) => p(`m${i}`)).join('');
	const body = [
		tr(tc(merged, '<w:vMerge w:val="restart"/>') + tc(p('a'))),
		tr(tc('<w:p/>', '<w:vMerge/>') + tc(p('b'))),
		tr(tc(p('c')) + tc(p('d'))),
	].join('');
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${ns}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${body}</w:tbl><w:sectPr/></w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	return layoutDocumentModel(model, measurer).pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.filter((block): block is LayoutTableBox => block.kind === 'table'),
		),
	);
}

describe('rows spanned by a vertical merge', () => {
	it('leaves rows alone when the merged text fits the rows it covers', async () => {
		const [table] = await tables(2);
		expect(table!.rows.map((row) => row.heightPx)).toEqual([20, 20, 20]);
	});

	it('grows only the last spanned row for taller merged text', async () => {
		const [table] = await tables(5);
		// Five lines (100 px) over two rows of one line each: the first keeps its own height.
		expect(table!.rows.map((row) => row.heightPx)).toEqual([20, 80, 20]);
		expect(table!.heightPx).toBe(120);
	});

	it('keeps a merge that cannot fit the page in its first row, split across pages', async () => {
		const fragments = await tables(120);
		expect(fragments.length).toBeGreaterThan(1);
		const lines = fragments.flatMap((table) =>
			table.rows.flatMap((row) => row.cells[0]!.flatMap((box) => box.lines)),
		);
		expect(lines.length).toBeGreaterThanOrEqual(120);
	});
});
