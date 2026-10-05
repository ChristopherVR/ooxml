import { describe, expect, it } from 'vitest';
import { parseRange } from '../address.js';
import { putCell } from '../cells.js';
import { createConditionalFormatEvaluator } from '../layout/index.js';
import type { ConditionalFormat, Workbook } from '../model.js';
import { loadXlsx } from '../read/index.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from '../write/index.js';
import { createEditSession } from './session.js';

const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};
const rule = (ref: string, text: string): ConditionalFormat => ({
	ranges: [R(ref)],
	rules: [{ type: 'containsText', text, style: { font: { bold: true } }, priority: 1 }],
});
const formats = (wb: Workbook) => wb.sheets[0]?.conditionalFormats ?? [];
const texts = (wb: Workbook) =>
	formats(wb)
		.flatMap((f) => f.rules)
		.sort((a, b) => a.priority - b.priority)
		.map((r) => ('text' in r ? r.text : r.type));

describe('conditional format edits', () => {
	it('replaces, removes and reorders rules as one undo step each', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb, { recalc: false });
		s.addConditionalFormat(0, rule('A1:A5', 'a'));
		s.addConditionalFormat(0, rule('B1:B5', 'b'));
		s.addConditionalFormat(0, rule('C1:C5', 'c'));
		expect(texts(wb)).toEqual(['c', 'b', 'a']);
		s.moveConditionalRule(0, 0, 0, 'up');
		expect(texts(wb)).toEqual(['c', 'a', 'b']);
		expect(s.undoLabel()).toBe('Rule priority');
		s.setConditionalRulePriority(0, 0, 0, 1);
		expect(texts(wb)).toEqual(['a', 'c', 'b']);
		s.replaceConditionalFormat(0, 1, rule('B2:B9', 'bb'));
		expect(formats(wb)[1]?.ranges).toEqual([R('B2:B9')]);
		expect(s.undoLabel()).toBe('Edit rule');
		s.removeConditionalFormat(0, 2);
		expect(texts(wb)).toEqual(['a', 'bb']);
		expect(
			formats(wb)
				.flatMap((f) => f.rules.map((r) => r.priority))
				.sort(),
		).toEqual([1, 2]);
		s.undo();
		s.undo();
		expect(texts(wb)).toEqual(['a', 'c', 'b']);
		expect(() => s.removeConditionalFormat(0, 9)).toThrow(RangeError);
	});

	it('reads, writes and evaluates "A Date Occurring" rules', async () => {
		const wb = createWorkbook();
		const sheet = wb.sheets[0];
		if (!sheet) throw new Error('sheet');
		const today = 45427;
		putCell(sheet, 0, 0, { value: today });
		putCell(sheet, 1, 0, { value: today - 1 });
		putCell(sheet, 2, 0, { value: 'x' });
		sheet.conditionalFormats.push({
			ranges: [R('A1:A3')],
			rules: [
				{
					type: 'timePeriod',
					timePeriod: 'yesterday',
					style: { fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFFF0000' } } },
					priority: 1,
				},
			],
		});
		const bytes = await saveXlsx(wb);
		const back = await loadXlsx(bytes);
		expect(back.sheets[0]?.conditionalFormats[0]?.rules[0]).toMatchObject({
			type: 'timePeriod',
			timePeriod: 'yesterday',
		});
		const JSZip = (await import('jszip')).default;
		const xml = await (
			await JSZip.loadAsync(bytes)
		)
			.file('xl/worksheets/sheet1.xml')
			?.async('text');
		expect(xml).toContain('timePeriod="yesterday"');
		expect(xml).toContain('<formula>FLOOR(A1,1)=TODAY()-1</formula>');
		const cf = createConditionalFormatEvaluator(back, 0, () => null, { today: () => today });
		expect(cf.at(0, 0)).toBeUndefined();
		expect(cf.at(1, 0)?.style?.fill).toBeDefined();
		expect(cf.at(2, 0)).toBeUndefined();
		const viaToday = createConditionalFormatEvaluator(back, 0, (f) =>
			f === 'TODAY()' ? today : null,
		);
		expect(viaToday.at(1, 0)?.style?.fill).toBeDefined();
	});
});
