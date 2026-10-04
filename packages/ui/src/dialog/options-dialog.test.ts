import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { clampOptionNumber, type OfficeOptionCategory } from './controls.js';
import { registerOfficeUi } from './index.js';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type OptionsDialog = HTMLElement & {
	categories: readonly OfficeOptionCategory[];
	values: Record<string, boolean | number | string>;
	category: string;
	open: boolean;
	show(): void;
};

const CATEGORIES: OfficeOptionCategory[] = [
	{
		id: 'general',
		label: 'General',
		description: 'General options for working with Visio.',
		sections: [
			{
				id: 'ui',
				title: 'User Interface options',
				controls: [
					{ kind: 'toggle', key: 'livePreview', label: 'Enable Live Preview' },
					{
						kind: 'select',
						key: 'screenTips',
						label: 'ScreenTip style',
						choices: [
							{ value: 'descriptions', label: 'Show feature descriptions' },
							{ value: 'off', label: 'Do not show ScreenTips' },
						],
					},
					{
						kind: 'toggle',
						key: 'miniToolbar',
						label: 'Mini Toolbar',
						disabled: 'Needs rich text',
					},
				],
			},
			{
				id: 'you',
				title: 'Personalize',
				controls: [{ kind: 'text', key: 'userName', label: 'User name', maxLength: 64 }],
			},
		],
	},
	{
		id: 'advanced',
		label: 'Advanced',
		sections: [
			{
				id: 'edit',
				title: 'Editing options',
				controls: [{ kind: 'number', key: 'undo', label: 'Maximum undo steps', min: 1, max: 99 }],
			},
		],
	},
	{ id: 'trust', label: 'Trust Center', sections: [], disabled: 'Macros are not supported.' },
];

function mount() {
	const dialog = document.createElement('office-ui-options-dialog') as OptionsDialog;
	document.body.append(dialog);
	dialog.categories = CATEGORIES;
	dialog.values = { livePreview: true, screenTips: 'descriptions', userName: 'Ada', undo: 20 };
	const root = dialog.shadowRoot!;
	const tab = (label: string) =>
		[...root.querySelectorAll<HTMLButtonElement>('[role="tab"]')].find(
			(el) => el.textContent === label,
		)!;
	const field = <T extends HTMLElement>(key: string) =>
		root.querySelector<T>(`[data-key="${key}"] input, [data-key="${key}"] select`)!;
	const action = (name: string) =>
		root.querySelector<HTMLButtonElement>(`[data-action="${name}"]`)!;
	return { dialog, root, tab, field, action };
}

describe('office-ui-options-dialog', () => {
	it('lists categories as tabs and renders the selected pane', () => {
		const { dialog, root, tab, field } = mount();
		dialog.show();
		expect(root.querySelector('office-ui-dialog')!.getAttribute('heading')).toBe('Options');
		expect([...root.querySelectorAll('[role="tab"]')].map((el) => el.textContent)).toEqual([
			'General',
			'Advanced',
			'Trust Center',
		]);
		expect(tab('General').getAttribute('aria-selected')).toBe('true');
		expect(root.querySelector('.lead')!.textContent).toBe(
			'General options for working with Visio.',
		);
		expect(field<HTMLInputElement>('livePreview').checked).toBe(true);
		expect(field<HTMLSelectElement>('screenTips').value).toBe('descriptions');
		const mini = field<HTMLInputElement>('miniToolbar');
		expect(mini.disabled).toBe(true);
		expect(mini.closest('label')!.title).toBe('Needs rich text');
		tab('Trust Center').click();
		expect(root.querySelector('.unavailable')!.textContent).toBe('Macros are not supported.');
	});

	it('emits only on OK with every value and the changed keys', () => {
		const { dialog, field, action } = mount();
		const change = vi.fn();
		dialog.addEventListener('office-options-change', (event) =>
			change((event as CustomEvent).detail),
		);
		dialog.show();
		field<HTMLInputElement>('livePreview').click();
		const name = field<HTMLInputElement>('userName');
		name.value = 'Grace';
		name.dispatchEvent(new Event('input'));
		action('ok').click();
		expect(dialog.open).toBe(false);
		expect(change).toHaveBeenCalledWith({
			values: { livePreview: false, screenTips: 'descriptions', userName: 'Grace', undo: 20 },
			changed: ['livePreview', 'userName'],
		});
		expect(dialog.values.userName).toBe('Grace');
	});

	it('discards the draft on Cancel and clamps numbers', () => {
		const { dialog, tab, field, action } = mount();
		const change = vi.fn();
		dialog.addEventListener('office-options-change', change);
		dialog.show();
		field<HTMLInputElement>('livePreview').click();
		action('cancel').click();
		expect(change).not.toHaveBeenCalled();
		dialog.show();
		expect(field<HTMLInputElement>('livePreview').checked).toBe(true);
		tab('Advanced').click();
		const undo = field<HTMLInputElement>('undo');
		undo.value = '500';
		undo.dispatchEvent(new Event('change'));
		expect(undo.value).toBe('99');
		undo.value = 'abc';
		undo.dispatchEvent(new Event('change'));
		expect(undo.value).toBe('99');
	});

	it('moves between categories with the arrow keys', () => {
		const { dialog, root, tab } = mount();
		dialog.show();
		tab('General').dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
		expect(dialog.category).toBe('advanced');
		expect(root.activeElement).toBe(tab('Advanced'));
		tab('Advanced').dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
		expect(dialog.category).toBe('trust');
	});

	it('clamps option numbers', () => {
		expect(clampOptionNumber('5', 1, 10)).toBe(5);
		expect(clampOptionNumber('0', 1, 10)).toBe(1);
		expect(clampOptionNumber('', 1, 10)).toBeUndefined();
		expect(clampOptionNumber('x', 1, 10)).toBeUndefined();
	});
});
