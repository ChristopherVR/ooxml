import { describe, expect, it } from 'vitest';
import {
	canDeleteControl,
	canEditControlContent,
	checkboxState,
	controlDate,
	selectListItem,
	setControlDate,
	toggleCheckbox,
} from './content-control-edit';
import type { ContentControl } from './content-control-model';

const base: ContentControl = { id: '1', level: 'inline', kind: 'checkbox', checked: false };

describe('content control editing', () => {
	it('toggles a checkbox with the default or the file glyphs', () => {
		expect(toggleCheckbox(base)).toMatchObject({ text: '☒', font: 'MS Gothic' });
		const custom = {
			...base,
			checked: true,
			uncheckedGlyph: { code: '25CB', font: 'Segoe UI Symbol' },
		};
		const off = toggleCheckbox(custom);
		expect(off.control.checked).toBe(false);
		expect(off.text).toBe('○');
		expect(off.font).toBe('Segoe UI Symbol');
		expect(checkboxState(base, false).text).toBe('☐');
	});

	it('selects a list item and clears the placeholder flag', () => {
		const list: ContentControl = {
			id: '2',
			level: 'inline',
			kind: 'dropDownList',
			showingPlaceholder: true,
			items: [
				{ display: 'Red', value: 'r' },
				{ display: 'Blue', value: 'b' },
			],
		};
		const change = selectListItem(list, 1)!;
		expect(change.text).toBe('Blue');
		expect(change.control.showingPlaceholder).toBeUndefined();
		expect(selectListItem(list, 5)).toBeUndefined();
	});

	it('formats a chosen date with the control date format and keeps fullDate', () => {
		const control: ContentControl = {
			id: '3',
			level: 'inline',
			kind: 'date',
			dateFormat: 'dd MMMM yyyy',
			dateLocale: 'en-US',
		};
		const change = setControlDate(control, new Date(2026, 2, 5));
		expect(change.text).toBe('05 March 2026');
		expect(change.control.fullDate).toBe('2026-03-05T00:00:00Z');
		expect(controlDate(change.control)?.getDate()).toBe(5);
	});

	it('reports what the lock setting allows', () => {
		expect(canEditControlContent({ ...base, lock: 'contentLocked' })).toBe(false);
		expect(canDeleteControl({ ...base, lock: 'contentLocked' })).toBe(true);
		expect(canDeleteControl({ ...base, lock: 'sdtLocked' })).toBe(false);
		expect(canEditControlContent({ ...base, lock: 'sdtLocked' })).toBe(true);
		expect(canEditControlContent({ ...base, lock: 'sdtContentLocked' })).toBe(false);
		expect(canDeleteControl(undefined)).toBe(true);
	});
});
