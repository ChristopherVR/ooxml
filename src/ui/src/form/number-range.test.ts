// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { createNumberRange } from './number-range';

it('pairs values, commits once, and cancels Escape, pointer cancellation and disabled gestures', () => {
	const input = document.createElement('input');
	input.type = 'number';
	input.min = '0';
	input.max = '100';
	input.step = '1';
	input.value = '37';
	let enabled = true;
	let changes = 0;
	let model = 37;
	const previews: Array<number | undefined> = [];
	const range = createNumberRange(input, {
		label: () => 'Transparency',
		enabled: () => enabled,
		onPreview: (value) => previews.push(value),
	});
	document.body.append(input, range.element);
	input.addEventListener('change', () => {
		changes++;
		model = input.valueAsNumber;
	});
	range.refresh();
	const slider = () => range.element.querySelector<HTMLInputElement>('input')!;
	const move = (value: string) => {
		slider().value = value;
		slider().dispatchEvent(new Event('input'));
	};
	move('23');
	expect(input.value).toBe('23');
	expect(model).toBe(37);
	range.element.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
	expect(input.value).toBe('37');
	expect(previews.at(-1)).toBeUndefined();
	slider().dispatchEvent(new Event('change'));
	expect(changes).toBe(0);
	move('23');
	move('24');
	slider().dispatchEvent(new Event('change'));
	expect(changes).toBe(1);
	expect(model).toBe(24);
	move('60');
	range.element.dispatchEvent(new Event('pointercancel'));
	expect(input.value).toBe('24');
	move('60');
	enabled = false;
	slider().dispatchEvent(new Event('change'));
	expect(input.value).toBe('24');
	expect(changes).toBe(1);
	expect(slider().disabled).toBe(true);
	range.element.remove();
	input.remove();
});
