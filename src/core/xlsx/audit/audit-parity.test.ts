import { describe, expect, it } from 'vitest';
import { createCalcEngine } from '../formula/engine';
import { book } from '../formula/test-helpers';
import { findErrors, withIgnoredError } from './error-check';
import { FormulaStepper } from './evaluate-steps';
import { directDependents, directPrecedents, traceNext } from './trace';

describe('OFFSET and INDIRECT precedent/dependent tracing', () => {
	it('traces INDIRECT targets on the same sheet and across sheets', () => {
		const wb = book({
			A1: 'B2',
			B2: 42,
			C1: '=INDIRECT(A1)',
			'Sheet2!D4': 99,
			D1: '=INDIRECT("Sheet2!D4")',
		});
		const calc = createCalcEngine(wb);
		calc.recalculateAll();
		const source = (c: { sheet: number; row: number; col: number }) => calc.precedents(c);

		const precC1 = directPrecedents(wb, { sheet: 0, row: 0, col: 2 }, source);
		expect(precC1).toEqual(
			[
				{ sheet: 0, row: 0, col: 0 },
				{ sheet: 0, row: 1, col: 1 },
			].map(({ sheet, row, col }) => ({
				sheet,
				range: { start: { row, col }, end: { row, col } },
			})),
		);

		const depsB2 = directDependents(wb, { sheet: 0, row: 1, col: 1 }, source);
		expect(depsB2).toEqual([{ sheet: 0, row: 0, col: 2 }]);

		const tracedD1 = traceNext(wb, [], 'precedents', { sheet: 0, row: 0, col: 3 }, source);
		expect(tracedD1.outcome).toBe('ok');
		expect(tracedD1.arrows).toHaveLength(1);
		expect(tracedD1.arrows[0]?.offSheet).toBe(true);
		expect(tracedD1.arrows[0]?.from).toEqual({
			sheet: 1,
			range: { start: { row: 3, col: 3 }, end: { row: 3, col: 3 } },
		});
	});

	it('traces OFFSET target ranges', () => {
		const wb = book({
			A1: 1,
			B2: 10,
			B3: 20,
			D1: '=SUM(OFFSET(A1,1,1,2,1))',
		});
		const calc = createCalcEngine(wb);
		calc.recalculateAll();
		const source = (c: { sheet: number; row: number; col: number }) => calc.precedents(c);

		const prec = directPrecedents(wb, { sheet: 0, row: 0, col: 3 }, source);
		expect(prec).toEqual([
			{ sheet: 0, range: { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } } },
			{ sheet: 0, range: { start: { row: 1, col: 1 }, end: { row: 2, col: 1 } } },
		]);
		expect(directDependents(wb, { sheet: 0, row: 2, col: 1 }, source)).toEqual([
			{ sheet: 0, row: 0, col: 3 },
		]);
	});
});

describe('Error Checking: two-digit-text-year and list-data-validation', () => {
	it('flags text dates with two-digit years and honors ignoredErrors', () => {
		const wb = book({
			A1: '1/15/26',
			A2: '15-Jan-26',
			A3: '1/15/2026',
			A4: 'hello',
		});
		const issues = findErrors(wb, 0, ['two-digit-text-year']);
		expect(issues).toEqual([
			{ sheet: 0, row: 0, col: 0, rule: 'two-digit-text-year' },
			{ sheet: 0, row: 1, col: 0, rule: 'two-digit-text-year' },
		]);

		wb.sheets[0]!.ignoredErrors = withIgnoredError(wb, issues[0]!);
		expect(findErrors(wb, 0, ['two-digit-text-year'])).toEqual([issues[1]]);
	});

	it('flags values outside a list data validation and allows matching entries', () => {
		const wb = book({
			A1: 'Red',
			A2: 'Purple',
			A3: 'blue',
		});
		wb.sheets[0]!.dataValidations.push({
			ranges: [{ start: { row: 0, col: 0 }, end: { row: 2, col: 0 } }],
			type: 'list',
			formula1: '"Red,Green,Blue"',
		});
		const issues = findErrors(wb, 0, ['list-data-validation']);
		expect(issues).toEqual([{ sheet: 0, row: 1, col: 0, rule: 'list-data-validation' }]);

		wb.sheets[0]!.ignoredErrors = withIgnoredError(wb, issues[0]!);
		expect(findErrors(wb, 0, ['list-data-validation'])).toEqual([]);
	});
});

describe('Evaluate Formula: step-by-step chosen branch for lazy functions', () => {
	const traceSteps = (cells: Record<string, number | string>) => {
		const wb = book(cells);
		const calc = createCalcEngine(wb);
		calc.recalculateAll();
		const stepper = new FormulaStepper(wb, calc, { sheet: 0, row: 0, col: 0 });
		const steps: string[] = [];
		for (let i = 0; i < 20; i++) {
			const view = stepper.view();
			steps.push(view.text);
			if (view.finished) break;
			stepper.evaluate();
		}
		return { steps, result: stepper.result() };
	};

	it('steps into only the chosen branch of IF without evaluating the unchosen branch', () => {
		const { steps, result } = traceSteps({
			A1: '=IF(B1>0, C1+10, 1/0)',
			B1: 5,
			C1: 2,
		});
		expect(steps).toEqual([
			'=IF(B1>0,C1+10,1/0)',
			'=IF(5>0,C1+10,1/0)',
			'=IF(TRUE,C1+10,1/0)',
			'=IF(TRUE,2+10,1/0)',
			'=IF(TRUE,12,1/0)',
			'=12',
		]);
		expect(result).toBe(12);
	});

	it('steps into fallback of IFERROR / IFNA only on matching error', () => {
		const errCase = traceSteps({ A1: '=IFERROR(1/0, B1+3)', B1: 4 });
		expect(errCase.steps).toEqual([
			'=IFERROR(1/0,B1+3)',
			'=IFERROR(#DIV/0!,B1+3)',
			'=IFERROR(#DIV/0!,4+3)',
			'=IFERROR(#DIV/0!,7)',
			'=7',
		]);

		const okCase = traceSteps({ A1: '=IFNA(B1+1, 1/0)', B1: 4 });
		expect(okCase.steps).toEqual(['=IFNA(B1+1,1/0)', '=IFNA(4+1,1/0)', '=IFNA(5,1/0)', '=5']);
	});

	it('steps into the chosen branch of CHOOSE, IFS, and SWITCH', () => {
		const choose = traceSteps({ A1: '=CHOOSE(2, 1/0, B1*3)', B1: 5 });
		expect(choose.steps).toEqual([
			'=CHOOSE(2,1/0,B1*3)',
			'=CHOOSE(2,1/0,5*3)',
			'=CHOOSE(2,1/0,15)',
			'=15',
		]);

		const ifs = traceSteps({ A1: '=IFS(FALSE, 1/0, TRUE, B1+1)', B1: 6 });
		expect(ifs.steps).toEqual([
			'=IFS(FALSE,1/0,TRUE,B1+1)',
			'=IFS(FALSE,1/0,TRUE,6+1)',
			'=IFS(FALSE,1/0,TRUE,7)',
			'=7',
		]);

		const sw = traceSteps({ A1: '=SWITCH(B1, 1, 1/0, 2, C1+5)', B1: 2, C1: 3 });
		expect(sw.steps).toEqual([
			'=SWITCH(B1,1,1/0,2,C1+5)',
			'=SWITCH(2,1,1/0,2,C1+5)',
			'=SWITCH(2,1,1/0,2,3+5)',
			'=SWITCH(2,1,1/0,2,8)',
			'=8',
		]);
	});
});
