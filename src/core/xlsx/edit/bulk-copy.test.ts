// Copy and paste of a large block: one copy of each format in the payload, same cells pasted.
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells';
import { internStyle } from '../styles';
import { createWorkbook, defaultCellStyle } from '../workbook';
import { toHtml } from './clipboard-html';
import { createEditSession } from './session';

describe('copy and paste of a large block', () => {
	it('share each format once in the payload and paste the same cells and formats', () => {
		const workbook = createWorkbook({ sheets: ['From', 'To'] });
		const bold = internStyle(workbook, {
			...defaultCellStyle(),
			font: { ...defaultCellStyle().font, bold: true },
		});
		const session = createEditSession(workbook, { autoRowHeight: false });
		const values = Array.from({ length: 100 }, (_, r) =>
			Array.from({ length: 50 }, (_, c) => r * c),
		);
		session.setRangeValues(0, { row: 0, col: 0 }, values);
		session.applyStyle(0, [{ start: { row: 0, col: 0 }, end: { row: 49, col: 49 } }], {
			font: { bold: true },
		});
		const block = { start: { row: 0, col: 0 }, end: { row: 99, col: 49 } };
		const payload = session.copy(0, block);
		const styles = new Set(payload.cells.data.flat().map((clip) => clip?.style));
		expect(styles.size).toBe(2);
		// The HTML a payload gives is the same whether cells share format objects or not.
		const unshared = {
			...payload.cells,
			data: payload.cells.data.map((row) =>
				row.map((clip) => (clip?.style ? { ...clip, style: structuredClone(clip.style) } : clip)),
			),
		};
		expect(toHtml(payload.cells, workbook.theme)).toBe(toHtml(unshared, workbook.theme));
		session.paste(1, { row: 0, col: 0 }, payload);
		const [from, to] = workbook.sheets;
		for (let row = 0; row < 100; row++)
			for (let col = 0; col < 50; col++) {
				const source = getCell(from!, row, col);
				const pasted = getCell(to!, row, col);
				expect(pasted?.value).toBe(source?.value);
				expect(pasted?.styleId ?? 0).toBe(source?.styleId ?? 0);
			}
		expect(getCell(to!, 0, 0)?.styleId).toBe(bold);
	});
});
