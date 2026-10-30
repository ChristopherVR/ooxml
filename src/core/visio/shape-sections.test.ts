import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { rewriteControlFormula } from './control-formula';
import { parseActionSteps } from './action-formula';
import { cell, fixture, shape } from './test-fixtures';

describe('User-defined and Scratch cells in the model', () => {
	it('reads rows with their cached values, formulas and prompts', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape(
						'1',
						cell('Width', 2) +
							cell('Height', 1) +
							`<Section N="User"><Row N="Mode"><Cell N="Value" V="3" F="1+2"/><Cell N="Prompt" V="Mode of the shape"/></Row><Row N="Label"><Cell N="Value" V="hello"/></Row></Section>` +
							`<Section N="Scratch"><Row IX="0">${cell('X', 0.25)}${cell('A', 4)}</Row></Section>`,
					)}</Shapes>`,
				},
			],
		});
		const model = (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
		expect(model.userCells).toEqual([
			{ name: 'Mode', value: '3', formula: '1+2', prompt: 'Mode of the shape' },
			{ name: 'Label', value: 'hello' },
		]);
		expect(model.scratch).toEqual([{ row: '0', values: { X: 0.25, A: 4 } }]);
		expect(model.controls).toBeUndefined();
		expect(model.actions).toBeUndefined();
	});
});

describe('rewriting a control position', () => {
	it('keeps BOUND limits and refuses formulas that are not a ratio', () => {
		expect(rewriteControlFormula('BOUND(Width*0.1,0,FALSE,0,Width)', 'x', 0.3)).toBe(
			'BOUND(Width*0.3,0,FALSE,0,Width)',
		);
		expect(rewriteControlFormula('Height*0.5', 'y', 0.25)).toBe('Height*0.25');
		expect(rewriteControlFormula(undefined, 'x', 0.25)).toBe('');
		expect(rewriteControlFormula('No Formula', 'x', 0.25)).toBe('');
		// A Width ratio cannot drive a Y position, and a computed position is not a ratio.
		expect(rewriteControlFormula('Width*0.5', 'y', 0.25)).toBeUndefined();
		expect(rewriteControlFormula('Width*0.5+User.Offset', 'x', 0.25)).toBeUndefined();
		expect(rewriteControlFormula('Width*0.5', 'x', Number.NaN)).toBeUndefined();
	});
});

describe('reading an action formula', () => {
	it('accepts SETF on a cell of the shape, alone or added together', () => {
		expect(parseActionSteps('SETF(GetRef(User.Mode),1)')).toMatchObject({
			steps: [{ target: 'User.Mode' }],
		});
		const two = parseActionSteps('SETF(GetRef(User.A),1)+SETF(GetRef(User.B),NOT(User.A))');
		expect('steps' in two && two.steps.map((step) => step.target)).toEqual(['User.A', 'User.B']);
	});

	it('names the reason for everything else', () => {
		const reason = (formula: string | undefined) => {
			const parsed = parseActionSteps(formula);
			return 'reason' in parsed ? parsed.reason : '';
		};
		expect(reason(undefined)).toMatch(/no formula/);
		expect(reason('OPENFILE("a.chm")')).toMatch(/OPENFILE/);
		expect(reason('SETF(GetRef(Sheet.2!User.A),1)')).toMatch(/owns the action/);
		expect(reason('SETF(User.A,1)')).toMatch(/GetRef/);
		expect(reason('SETF(GetRef(User.A),INDIRECT("x"))')).toMatch(/cannot evaluate/);
	});
});
