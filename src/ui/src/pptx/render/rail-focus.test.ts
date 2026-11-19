// @vitest-environment jsdom
import { afterEach, describe, expect, it } from 'vitest';
import { focusCurrentRow, railHasFocus } from './rail-focus';

afterEach(() => document.body.replaceChildren());

/** A rail list of `count` rows, the second current, and a button outside it. */
function rail(count = 3) {
	const list = document.createElement('div');
	list.dataset.pptxChrome = 'slide-list';
	const rows = Array.from({ length: count }, (_, index) => {
		const row = document.createElement('button');
		row.dataset.pptxChrome = 'slide-row';
		if (index === 1) row.setAttribute('aria-current', 'true');
		list.append(row);
		return row;
	});
	const outside = document.createElement('button');
	document.body.append(list, outside);
	railHasFocus(list);
	return { list, rows, outside };
}

describe('rail focus', () => {
	it('still holds focus after a render removes the focused row, until focus moves on', () => {
		const { list, rows, outside } = rail();
		rows[2]!.focus();
		rows[2]!.remove();
		expect(document.activeElement).toBe(document.body);
		expect(railHasFocus(list)).toBe(true);
		focusCurrentRow(list);
		expect(document.activeElement).toBe(rows[1]);
		outside.focus();
		expect(railHasFocus(list)).toBe(false);
	});

	it('lets a press outside the rail take focus away from it', () => {
		const { list, rows } = rail();
		rows[0]!.focus();
		rows[0]!.remove();
		document.body.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, composed: true }));
		expect(railHasFocus(list)).toBe(false);
	});
});
