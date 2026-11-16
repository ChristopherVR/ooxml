import { describe, expect, it } from 'vitest';
import { chartErrorBarSourceFromText } from '../../chart/edit/error-bar-source';
import { putCell } from '../cells';
import { createWorkbook } from '../workbook';
import { workbookChartRangeNumbers } from './chart-range-numbers';

describe('workbook ranges for chart edits', () => {
	it('reads one row or column of stored numbers and refuses anything else', () => {
		const workbook = createWorkbook({ sheets: ['Data', 'My sheet'] });
		const [data, other] = workbook.sheets;
		[4, 5, 6].forEach((value, row) => putCell(data!, row + 1, 2, { value }));
		putCell(other!, 0, 0, { value: 9 });
		putCell(other!, 0, 1, { value: 'text' });
		const resolve = workbookChartRangeNumbers(workbook, data!);
		expect(resolve('Data!$C$2:$C$4')).toEqual([4, 5, 6]);
		expect(resolve('C2:C3')).toEqual([4, 5]);
		expect(resolve("'My sheet'!A1")).toEqual([9]);
		for (const bad of ["'My sheet'!A1:B1", 'Data!C1:C4', 'Data!B2:C4', 'Nope!A1'])
			expect(resolve(bad), bad).toBeUndefined();
		const source = chartErrorBarSourceFromText('=Data!$C$2:$C$4', {
			pointCount: 3,
			resolveRange: resolve,
		});
		expect(source?.cache?.points.map((point) => point.value)).toEqual(['4', '5', '6']);
	});
});
