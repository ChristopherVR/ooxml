import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { putCell } from '../cells.js';
import type { ConditionalRule, DifferentialStyle } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from './index.js';

const RED: DifferentialStyle = {
	fill: { type: 'pattern', pattern: 'solid', bgColor: { rgb: 'FFFF0000' } },
};

describe('conditional-format round trip', () => {
	it('keeps stopIfTrue on every rule type and cfvo gte="0"', async () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0]!;
		for (let i = 0; i < 10; i++) putCell(sheet, i, 0, { value: i });
		const rules: ConditionalRule[] = [
			{ type: 'top10', rank: 3, style: RED, priority: 1, stopIfTrue: true },
			{ type: 'aboveAverage', style: RED, priority: 2, stopIfTrue: true },
			{ type: 'duplicateValues', style: RED, priority: 3, stopIfTrue: true },
			{ type: 'containsText', text: 'x', style: RED, priority: 4, stopIfTrue: true },
			{
				type: 'iconSet',
				iconSet: '3Arrows',
				thresholds: [
					{ type: 'percent', value: '0' },
					{ type: 'num', value: '4', gte: false },
					{ type: 'num', value: '7' },
				],
				priority: 5,
				stopIfTrue: true,
			},
		];
		sheet.conditionalFormats.push({
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 9, col: 0 } }],
			rules,
		});
		const bytes = await saveXlsx(wb);
		const xml = await (
			await JSZip.loadAsync(bytes)
		)
			.file('xl/worksheets/sheet1.xml')!
			.async('string');
		expect(xml.match(/stopIfTrue="1"/g)).toHaveLength(5);
		expect(xml).toContain('<cfvo type="num" val="4" gte="0"/>');
		expect(xml).toContain('<cfvo type="num" val="7"/>');
		const back = await loadXlsx(bytes);
		const read = back.sheets[0]!.conditionalFormats.flatMap((f) => f.rules);
		expect(read.map((r) => r.stopIfTrue)).toEqual([true, true, true, true, true]);
		const icon = read.find((r) => r.type === 'iconSet');
		expect(icon?.type === 'iconSet' && icon.thresholds.map((t) => t.gte)).toEqual([
			undefined,
			false,
			undefined,
		]);
	});
});
