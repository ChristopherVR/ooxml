import { describe, expect, it } from 'vitest';
import { createEditSession, createWorkbook, getCell } from '../../index';
import { createForecastSheet, ForecastError, type ForecastSheetRequest } from './forecast-sheet';

// Header in row 1, ten monthly-style points (1..10) on a straight line 3, 5, 7, ...
function setup(extra: Partial<ForecastSheetRequest> = {}) {
	const wb = createWorkbook({ sheets: ['Data'] });
	const s = createEditSession(wb);
	s.setCellValue(0, 0, 0, 'Day');
	s.setCellValue(0, 0, 1, 'Sales');
	for (let i = 1; i <= 10; i++) {
		s.setCellValue(0, i, 0, i);
		s.setCellValue(0, i, 1, 1 + 2 * i);
	}
	const request: ForecastSheetRequest = {
		sheet: 0,
		timeline: { start: { row: 0, col: 0 }, end: { row: 10, col: 0 } },
		values: { start: { row: 0, col: 1 }, end: { row: 10, col: 1 } },
		end: 14,
		seasonality: 0,
		...extra,
	};
	return { wb, s, request };
}
const at = (wb: ReturnType<typeof setup>['wb'], sheet: number, row: number, col: number) =>
	getCell(wb.sheets[sheet]!, row, col);

describe('createForecastSheet', () => {
	it('builds the table with FORECAST.ETS formulas and confidence bounds on a new sheet', () => {
		const { wb, s, request } = setup();
		const result = createForecastSheet(s, request);
		expect(result).toEqual({ sheet: 1, historyRows: 10, forecastRows: 4 });
		expect(wb.sheets.map((x) => x.name)).toEqual(['Data', 'Sheet2']);
		expect(wb.activeSheet).toBe(1);
		expect([0, 1, 2, 3, 4].map((c) => at(wb, 1, 0, c)?.value)).toEqual([
			'Timeline',
			'Values',
			'Forecast',
			'Lower Confidence Bound',
			'Upper Confidence Bound',
		]);
		// History rows keep the data; the forecast column starts from the last actual value.
		expect([at(wb, 1, 1, 0)?.value, at(wb, 1, 1, 1)?.value]).toEqual([1, 3]);
		expect(at(wb, 1, 10, 2)?.formula).toBe('B11');
		// Forecast rows continue the timeline and use absolute ranges.
		expect([11, 12, 13, 14].map((r) => at(wb, 1, r, 0)?.value)).toEqual([11, 12, 13, 14]);
		expect(at(wb, 1, 11, 2)?.formula).toBe('FORECAST.ETS(A12,$B$2:$B$11,$A$2:$A$11,0,1,1)');
		expect(at(wb, 1, 11, 3)?.formula).toBe(
			'C12-FORECAST.ETS.CONFINT(A12,$B$2:$B$11,$A$2:$A$11,0.95,0,1,1)',
		);
		expect(at(wb, 1, 11, 4)?.formula).toContain('C12+FORECAST.ETS.CONFINT(');
		// An exact straight line is extended exactly, with a zero-width interval.
		expect(at(wb, 1, 11, 2)?.value as number).toBeCloseTo(23, 4);
		expect(at(wb, 1, 14, 2)?.value as number).toBeCloseTo(29, 4);
		expect(at(wb, 1, 14, 3)?.value as number).toBeCloseTo(29, 4);
		expect(at(wb, 1, 14, 4)?.value as number).toBeCloseTo(29, 4);
	});

	it('draws a line chart over the values, forecast and bounds with the timeline as categories', () => {
		const { wb, s, request } = setup();
		createForecastSheet(s, request);
		const chart = wb.sheets[1]!.drawings[0];
		expect(chart?.kind).toBe('chart');
		if (chart?.kind !== 'chart') return;
		expect(chart.chartType).toBe('line');
		expect(chart.series.map((x) => x.name)).toEqual([
			'Values',
			'Forecast',
			'Lower Confidence Bound',
			'Upper Confidence Bound',
		]);
		expect(chart.series[1]?.valuesRef).toBe('Sheet2!$C$2:$C$15');
		expect(chart.series[0]?.categoriesRef).toBe('Sheet2!$A$2:$A$15');
		expect(chart.title).toBe('Sales');
	});

	it('honours the options: no interval, statistics table, column chart, explicit season', () => {
		const { wb, s, request } = setup({
			confidence: null,
			statistics: true,
			chart: 'column',
			seasonality: 'auto',
			aggregate: 7,
			fill: 'zeros',
		});
		createForecastSheet(s, request);
		expect(at(wb, 1, 0, 3)?.value).toBeUndefined();
		expect(at(wb, 1, 11, 2)?.formula).toBe('FORECAST.ETS(A12,$B$2:$B$11,$A$2:$A$11,1,0,7)');
		expect(at(wb, 1, 0, 4)?.value).toBe('Statistic');
		expect(at(wb, 1, 1, 4)?.value).toBe('Alpha');
		expect(at(wb, 1, 7, 4)?.value).toBe('RMSE');
		expect(at(wb, 1, 5, 5)?.formula).toBe('FORECAST.ETS.STAT($B$2:$B$11,$A$2:$A$11,5,1,0,7)');
		const chart = wb.sheets[1]!.drawings[0];
		expect(chart?.kind === 'chart' && chart.chartType).toBe('column');
		expect(chart?.kind === 'chart' && chart.series.length).toBe(2);
	});

	it('is one undo step', () => {
		const { wb, s, request } = setup();
		createForecastSheet(s, request);
		s.undo();
		expect(wb.sheets.map((x) => x.name)).toEqual(['Data']);
		s.redo();
		expect(wb.sheets).toHaveLength(2);
		expect(wb.sheets[1]?.drawings).toHaveLength(1);
	});

	it('reads monthly dates and keeps their format on the new timeline', () => {
		const wb = createWorkbook({ sheets: ['Data'] });
		const s = createEditSession(wb);
		const serial = (m: number) => Math.round(Date.UTC(2024, m, 1) / 86400000) + 25569;
		for (let i = 0; i < 6; i++) {
			s.setCellValue(0, i, 0, serial(i));
			s.setCellValue(0, i, 1, 10 + 3 * i);
		}
		s.applyStyle(0, [{ start: { row: 0, col: 0 }, end: { row: 5, col: 0 } }], {
			numFmt: 'yyyy-mm-dd',
		});
		createForecastSheet(s, {
			sheet: 0,
			timeline: { start: { row: 0, col: 0 }, end: { row: 5, col: 0 } },
			values: { start: { row: 0, col: 1 }, end: { row: 5, col: 1 } },
			end: serial(8),
			seasonality: 0,
			chart: null,
		});
		const sheet = wb.sheets[1]!;
		expect([7, 8, 9].map((r) => getCell(sheet, r, 0)?.value)).toEqual([
			serial(6),
			serial(7),
			serial(8),
		]);
		expect(getCell(sheet, 7, 0)?.styleId).toBeDefined();
		expect(sheet.drawings).toHaveLength(0);
	});

	it('accepts a row layout', () => {
		const wb = createWorkbook({ sheets: ['Data'] });
		const s = createEditSession(wb);
		for (let i = 0; i < 6; i++) {
			s.setCellValue(0, 0, i, i + 1);
			s.setCellValue(0, 1, i, 4 * i);
		}
		const result = createForecastSheet(s, {
			sheet: 0,
			timeline: { start: { row: 0, col: 0 }, end: { row: 0, col: 5 } },
			values: { start: { row: 1, col: 0 }, end: { row: 1, col: 5 } },
			end: 8,
			seasonality: 0,
		});
		expect(result.forecastRows).toBe(2);
		expect(getCell(wb.sheets[1]!, 8, 2)?.value as number).toBeCloseTo(28, 4);
	});

	it('refuses requests it cannot honour', () => {
		const problem = (extra: Partial<ForecastSheetRequest>) => {
			const { s, request } = setup(extra);
			try {
				createForecastSheet(s, request);
			} catch (error) {
				return error instanceof ForecastError ? error.problem : String(error);
			}
			return undefined;
		};
		expect(problem({ end: 10 })).toBe('end-before-last');
		expect(problem({ end: 5 })).toBe('end-before-last');
		expect(problem({ end: 1_000_000 })).toBe('too-long');
		expect(problem({ confidence: 1 })).toBe('confidence');
		expect(problem({ seasonality: 1.5 })).toBe('seasonality');
		expect(problem({ aggregate: 9 })).toBe('aggregate');
		expect(problem({ values: { start: { row: 0, col: 1 }, end: { row: 5, col: 1 } } })).toBe(
			'shape',
		);
		expect(
			problem({
				timeline: { start: { row: 0, col: 0 }, end: { row: 10, col: 1 } },
			}),
		).toBe('shape');
	});

	it('refuses an uneven timeline and non-numeric data without changing the workbook', () => {
		const { wb, s, request } = setup();
		s.setCellValue(0, 5, 0, 5.3);
		expect(() => createForecastSheet(s, request)).toThrow(ForecastError);
		s.setCellValue(0, 5, 0, 5);
		s.setCellValue(0, 3, 1, 'x');
		expect(() => createForecastSheet(s, request)).toThrow(/numbers/);
		expect(wb.sheets).toHaveLength(1);
	});
});
