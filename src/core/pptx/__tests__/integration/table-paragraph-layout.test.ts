import { promises as fs } from 'node:fs';
import path from 'node:path';

import { describe, it, expect } from 'vitest';

import { PptxHandler } from '../../core/PptxHandler';
import type { TablePptxElement } from '../../core/types/elements';
import { requireFixture } from '../require-fixture';

/**
 * `table-paragraph-layout.pptx` is saved by PowerPoint for Mac 16. Its table
 * has one case per row; the second column's cell sets the paragraph layout
 * the row names.
 */
const fixturePath = requireFixture(
	path.resolve(__dirname, '../fixtures/table-paragraph-layout.pptx'),
);

describe('table cell paragraph layout from a PowerPoint deck', () => {
	it("reads each paragraph's own alignment, line spacing, spacing and indent", async () => {
		const buf = await fs.readFile(fixturePath);
		const data = await new PptxHandler().load(
			buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer,
		);
		const table = data.slides[0].elements.find(
			(element) => element.type === 'table',
		) as TablePptxElement;
		const cases = Object.fromEntries(
			table.tableData!.rows.slice(1).map((row) => [row.cells[0].text, row.cells[1].paragraphs]),
		);
		const zero = { paragraphMarginLeft: 0, paragraphIndent: 0 };
		expect(cases['Alignment per paragraph']).toStrictEqual([
			{ align: 'right', ...zero },
			{ align: 'center', ...zero },
		]);
		expect(cases['Exact 10pt line spacing at 9pt']).toStrictEqual([
			{ ...zero, lineSpacingExactPt: 10 },
		]);
		expect(cases['Line spacing 0.8 (multiple)']).toStrictEqual([{ ...zero, lineSpacing: 0.8 }]);
		// 12pt is 16px.
		expect(cases['Space after 12pt between paragraphs']).toStrictEqual(
			Array.from({ length: 3 }, () => ({ ...zero, paragraphSpacingAfter: 16 })),
		);
		// 0.25in is 24px.
		expect(cases['Hanging indent 0.25in']).toStrictEqual([
			{ paragraphMarginLeft: 24, paragraphIndent: -24 },
		]);
	});
});
