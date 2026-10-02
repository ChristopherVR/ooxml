// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import type { FunctionInfo } from '@christophervr/xlsx-core';
import { createFormulaField } from './formula-field.js';
import { createFunctionAssist } from './function-assist.js';

const catalog: FunctionInfo[] = [
	{
		name: 'SUM',
		category: 'Math',
		syntax: 'SUM(number1, [number2], ...)',
		description: 'Adds numbers.',
	},
	{
		name: 'SUMIF',
		category: 'Math',
		syntax: 'SUMIF(range, criteria, [sum_range])',
		description: 'Adds if.',
	},
	{
		name: 'IF',
		category: 'Logical',
		syntax: 'IF(logical_test, value_if_true, [value_if_false])',
		description: 'Chooses.',
	},
];

afterEach(() => document.body.replaceChildren());

function setup() {
	const container = document.createElement('div');
	document.body.append(container);
	const field = createFormulaField(document, 'test', 'Formula');
	container.append(field.root);
	const assist = createFunctionAssist(document, container, (k) => k, catalog);
	const type = (text: string, caret = text.length) => {
		field.setText(text, caret);
		assist.update(field, { left: 0, top: 20 });
	};
	const key = (k: string) =>
		assist.handleKey(new KeyboardEvent('keydown', { key: k, cancelable: true }), field);
	return { container, field, assist, type, key };
}

const options = (root: Element) =>
	[...root.querySelectorAll('[role="option"]')].map((o) => o.textContent);

describe('function assist', () => {
	it('lists functions matching the typed prefix', () => {
		const { container, assist, type } = setup();
		type('=su');
		expect(assist.isOpen()).toBe(true);
		expect(options(container)).toEqual(['SUM', 'SUMIF']);
		expect(container.querySelector('[aria-selected="true"]')?.textContent).toBe('SUM');
		expect(container.querySelector('.xg-assist-desc')?.textContent).toBe('Adds numbers.');
	});

	it('moves with arrows and accepts with Tab', () => {
		const { container, field, type, key } = setup();
		let inputs = 0;
		field.input.addEventListener('input', () => inputs++);
		type('=1+su');
		expect(key('ArrowDown')).toBe(true);
		expect(container.querySelector('[aria-selected="true"]')?.textContent).toBe('SUMIF');
		expect(key('Tab')).toBe(true);
		expect(field.input.value).toBe('=1+SUMIF(');
		expect(inputs).toBe(1);
	});

	it('accepts on click', () => {
		const { container, field, type } = setup();
		type('=I');
		container
			.querySelector<HTMLElement>('[data-name="IF"]')!
			.dispatchEvent(new MouseEvent('mousedown'));
		expect(field.input.value).toBe('=IF(');
	});

	it('shows the argument tooltip with the current parameter in bold', () => {
		const { container, assist, type } = setup();
		type('=IF(A1>0,');
		expect(assist.isOpen()).toBe(true);
		expect(container.querySelector('.xg-assist-name')?.textContent).toBe('IF');
		expect(container.querySelector('.xg-assist-param-active')?.textContent).toBe('value_if_true');
		type('=SUM(1,2,3');
		expect(container.querySelector('.xg-assist-param-active')?.textContent).toBe('[number2]');
	});

	it('closes outside formulas and on Escape, leaving other keys alone', () => {
		const { assist, type, key } = setup();
		type('plain text');
		expect(assist.isOpen()).toBe(false);
		type('=su');
		expect(key('a')).toBe(false);
		expect(key('Escape')).toBe(true);
		expect(assist.isOpen()).toBe(false);
		expect(key('ArrowDown')).toBe(false);
	});

	it('removes itself on destroy', () => {
		const { container, assist } = setup();
		assist.destroy();
		expect(container.querySelector('.xg-assist')).toBeNull();
	});
});
