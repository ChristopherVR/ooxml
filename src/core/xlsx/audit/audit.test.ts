import { describe, expect, it } from 'vitest';
import { parseFormula } from '../formula/parser';
import { book, engine, locate } from '../formula/test-helpers';
import { printFormula } from './ast-text';
import { adjacentIssue, findErrors, withIgnoredError } from './error-check';
import { FormulaStepper } from './evaluate-steps';
import { directDependents, directPrecedents, removeArrows, traceNext } from './trace';
import { describeWatch } from './watch';

const at = (wb: ReturnType<typeof book>, key: string) => locate(wb, key);

describe('printFormula', () => {
	it.each([
		'1+2*3',
		'(1+2)*3',
		'-A1^2',
		'SUM(A1:B2,"a""b")',
		'A1&"x"=B1',
		'IF(A1>1,{1,2;3,4},#N/A)',
		"Sheet2!A1+'My Sheet'!$B$2%",
		'1-(3-4)',
	])('round-trips %s', (f) => {
		expect(printFormula(parseFormula(f)).text).toBe(f);
	});
});

describe('trace', () => {
	const wb = book({
		A1: 1,
		A2: 2,
		B1: '=SUM(A1:A2)',
		C1: '=B1*2+Sheet2!A1',
		D1: '=C1',
		'Sheet2!A1': 7,
	});

	it('finds direct precedents and dependents', () => {
		expect(directPrecedents(wb, at(wb, 'C1')).map((p) => p.sheet)).toEqual([0, 1]);
		expect(directDependents(wb, at(wb, 'B1'))).toEqual([at(wb, 'C1')]);
	});

	it('adds a level per repeat and marks off-sheet arrows', () => {
		let r = traceNext(wb, [], 'precedents', at(wb, 'C1'));
		expect(r.outcome).toBe('ok');
		expect(r.arrows.map((a) => [a.level, a.offSheet])).toEqual([
			[1, false],
			[1, true],
		]);
		r = traceNext(wb, r.arrows, 'precedents', at(wb, 'C1'));
		expect(r.arrows.filter((a) => a.level === 2)).toHaveLength(1);
		const again = traceNext(wb, r.arrows, 'precedents', at(wb, 'C1'));
		expect(again.outcome).toBe('no-more');
		expect(removeArrows(r.arrows, 'precedents', at(wb, 'C1')).every((a) => a.level === 1)).toBe(
			true,
		);
		expect(removeArrows(r.arrows, 'all')).toEqual([]);
	});

	it('reports a cell without a formula', () => {
		expect(traceNext(wb, [], 'precedents', at(wb, 'A1')).outcome).toBe('no-formula');
		expect(traceNext(wb, [], 'dependents', at(wb, 'D1')).outcome).toBe('no-dependents');
	});

	it('traces dependents level by level', () => {
		const r = traceNext(
			wb,
			traceNext(wb, [], 'dependents', at(wb, 'B1')).arrows,
			'dependents',
			at(wb, 'B1'),
		);
		expect(r.arrows.map((a) => a.to.col)).toEqual([2, 3]);
	});

	it('draws error sources red', () => {
		const e = book({ A1: '=1/0', B1: '=A1+1' });
		engine(e).recalculateAll();
		expect(traceNext(e, [], 'precedents', at(e, 'B1')).arrows[0]?.error).toBe(true);
	});
});

describe('FormulaStepper', () => {
	it('reduces a formula step by step', () => {
		const wb = book({ A1: 2, A2: 3, B1: '=(A1+A2)*4' });
		const stepper = new FormulaStepper(wb, engine(wb), at(wb, 'B1'));
		const texts: string[] = [];
		let view = stepper.view();
		expect(view.underline && view.text.slice(view.underline.start, view.underline.end)).toBe('A1');
		while (!view.finished) {
			texts.push(view.text);
			stepper.evaluate();
			view = stepper.view();
		}
		expect(texts).toEqual(['=(A1+A2)*4', '=(2+A2)*4', '=(2+3)*4', '=5*4']);
		expect(view.text).toBe('=20');
		expect(stepper.result()).toBe(20);
	});

	it('steps in to a formula cell and out again', () => {
		const wb = book({ A1: 2, B1: '=A1*10', C1: '=B1+1' });
		const calc = engine(wb);
		calc.recalculateAll();
		const stepper = new FormulaStepper(wb, calc, at(wb, 'C1'));
		expect(stepper.view().canStepIn).toBe(true);
		stepper.stepIn();
		expect(stepper.view().text).toBe('=A1*10');
		expect(stepper.view().canStepOut).toBe(true);
		stepper.stepOut();
		expect(stepper.view().text).toBe('=20+1');
	});

	it('takes only the chosen IF branch', () => {
		const wb = book({ A1: 1, B1: '=IF(A1>0,"yes",1/0)' });
		const stepper = new FormulaStepper(wb, engine(wb), at(wb, 'B1'));
		stepper.evaluate();
		stepper.evaluate();
		stepper.evaluate();
		expect(stepper.result()).toBe('yes');
	});

	it('refuses cells without a formula', () => {
		const wb = book({ A1: 1 });
		expect(FormulaStepper.canEvaluate(wb, at(wb, 'A1'))).toBe(false);
	});
});

describe('findErrors', () => {
	it('applies each rule and honours Ignore Error', () => {
		const wb = book({
			A1: 1,
			A2: 2,
			A3: 3,
			A4: 4,
			A5: '=SUM(A2:A3)',
			B1: '5',
			B2: '=1/0',
			C1: '=A1+1',
			C2: '=A2+1',
			C3: '=A3*9',
			C4: '=A4+1',
		});
		engine(wb).recalculateAll();
		const issues = findErrors(wb, 0);
		const rules = issues.map((i) => `${i.row},${i.col}:${i.rule}`);
		expect(rules).toContain('0,1:number-as-text');
		expect(rules).toContain('1,1:error-value');
		expect(rules).toContain('4,0:omits-adjacent-cells');
		expect(rules).toContain('2,2:inconsistent-formula');
		const first = issues[0];
		if (!first) throw new Error('no issue');
		wb.sheets[0]!.ignoredErrors = withIgnoredError(wb, first);
		expect(findErrors(wb, 0)).toHaveLength(issues.length - 1);
		expect(adjacentIssue(issues, { sheet: 0, row: 99, col: 0 }, 1)).toBe(issues[0]);
		expect(adjacentIssue(issues, { sheet: 0, row: 0, col: 0 }, -1)).toBe(issues[issues.length - 1]);
	});

	it('flags unlocked formula cells', () => {
		const wb = book({ A1: '=1+1' });
		wb.styles.push({
			...(wb.styles[0] as (typeof wb.styles)[number]),
			protection: { locked: false },
		});
		const cell = wb.sheets[0]?.rows.get(0)?.get(0);
		if (cell) cell.styleId = wb.styles.length - 1;
		expect(findErrors(wb, 0).map((i) => i.rule)).toEqual(['unlocked-formula']);
	});
});

describe('describeWatch', () => {
	it('reads value, formula and a defined name', () => {
		const wb = book({ A1: '=1+2' });
		wb.definedNames.push({ name: 'Total', formula: 'Sheet1!$A$1' });
		engine(wb).recalculateAll();
		expect(describeWatch(wb, at(wb, 'A1'))).toEqual({
			sheet: 'Sheet1',
			name: 'Total',
			cell: 'A1',
			value: 3,
			formula: '=1+2',
		});
	});
});
