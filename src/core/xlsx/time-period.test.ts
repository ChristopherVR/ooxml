import { describe, expect, it } from 'vitest';
import { inTimePeriod, timePeriodFormula } from './time-period.js';

// 2024-05-15 is a Wednesday, serial 45427.
const TODAY = 45427;

describe('time periods', () => {
	it('writes the formulas Excel stores', () => {
		expect(timePeriodFormula('today', 'A1')).toBe('FLOOR(A1,1)=TODAY()');
		expect(timePeriodFormula('lastMonth', 'B2')).toBe(
			'AND(MONTH(B2)=MONTH(EDATE(TODAY(),0-1)),YEAR(B2)=YEAR(EDATE(TODAY(),0-1)))',
		);
	});
	it('tests days, weeks and months', () => {
		expect(inTimePeriod('today', TODAY + 0.5, TODAY)).toBe(true);
		expect(inTimePeriod('yesterday', TODAY - 1, TODAY)).toBe(true);
		expect(inTimePeriod('tomorrow', TODAY, TODAY)).toBe(false);
		expect(inTimePeriod('last7Days', TODAY - 6, TODAY)).toBe(true);
		expect(inTimePeriod('last7Days', TODAY - 7, TODAY)).toBe(false);
		// This week: Sunday 2024-05-12 to Saturday 2024-05-18.
		expect(inTimePeriod('thisWeek', TODAY - 3, TODAY)).toBe(true);
		expect(inTimePeriod('thisWeek', TODAY + 3, TODAY)).toBe(true);
		expect(inTimePeriod('thisWeek', TODAY + 4, TODAY)).toBe(false);
		expect(inTimePeriod('lastWeek', TODAY - 4, TODAY)).toBe(true);
		expect(inTimePeriod('lastWeek', TODAY - 3, TODAY)).toBe(false);
		expect(inTimePeriod('nextWeek', TODAY + 4, TODAY)).toBe(true);
		expect(inTimePeriod('nextWeek', TODAY + 11, TODAY)).toBe(false);
		expect(inTimePeriod('thisMonth', TODAY - 14, TODAY)).toBe(true);
		expect(inTimePeriod('lastMonth', TODAY - 15, TODAY)).toBe(true);
		expect(inTimePeriod('nextMonth', TODAY + 17, TODAY)).toBe(true);
	});
});
