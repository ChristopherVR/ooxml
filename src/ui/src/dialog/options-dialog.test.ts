import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { clampOptionNumber, type OfficeOptionCategory } from '../controls';
import { registerOfficeUi } from '../index';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

type OptionsDialog = HTMLElement & {
	categories: readonly OfficeOptionCategory[];
	values: Record<string, boolean | number | string>;
	category: string;
	open: boolean;
	live: boolean;
	show(): void;
	defaults: Record<string, boolean | number | string> | null;
	labels: Record<string, string>;
	updateComplete: Promise<unknown>;
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

	it('resets to the defaults, translates its buttons and reports a cancel', async () => {
		const { dialog, root, field, action } = mount();
		dialog.show();
		expect(root.querySelector('[data-action="reset"]')).toBeNull();
		dialog.defaults = {
			livePreview: true,
			screenTips: 'descriptions',
		};
		dialog.labels = {
			ok: 'Aceptar',
			cancel: 'Cancelar',
			reset: 'Restablecer',
		};
		const cancelled = vi.fn();
		dialog.addEventListener('office-options-cancel', cancelled);
		await dialog.updateComplete;
		expect(action('ok').textContent!.trim()).toBe('Aceptar');
		expect(action('reset').textContent!.trim()).toBe('Restablecer');
		field<HTMLInputElement>('livePreview').click();
		expect(field<HTMLInputElement>('livePreview').checked).toBe(false);
		action('reset').click();
		await dialog.updateComplete;
		expect(field<HTMLInputElement>('livePreview').checked).toBe(true);
		action('cancel').click();
		expect(cancelled).toHaveBeenCalledTimes(1);
		action('ok');
		dialog.show();
		action('ok').click();
		expect(cancelled).toHaveBeenCalledTimes(1);
	});

	it('enables a dependent field only while its toggle is on, and projects a custom pane', async () => {
		const { dialog, root, tab } = mount();
		dialog.categories = [
			{
				id: 'calc',
				label: 'Formulas',
				sections: [
					{
						id: 'iter',
						title: 'Iteration',
						controls: [
							{ kind: 'toggle', key: 'iterate', label: 'Iterate' },
							{
								kind: 'number',
								key: 'count',
								label: 'Count',
								min: 1,
								max: 9,
								enabledBy: 'iterate',
							},
						],
					},
				],
			},
			{ id: 'ribbon', label: 'Customize Ribbon', sections: [], custom: true },
		];
		dialog.values = { iterate: false, count: 3 };
		const pane = document.createElement('div');
		pane.slot = 'pane-ribbon';
		pane.textContent = 'commands';
		dialog.append(pane);
		dialog.show();
		const wait = () => dialog.updateComplete;
		await wait();
		const count = root.querySelector<HTMLInputElement>('[data-key="count"] input')!;
		expect(count.disabled).toBe(true);
		root.querySelector<HTMLInputElement>('[data-key="iterate"] input')!.click();
		await wait();
		expect(count.disabled).toBe(false);
		tab('Customize Ribbon').click();
		await wait();
		const slot = root.querySelector<HTMLSlotElement>('slot[name="pane-ribbon"]')!;
		expect(slot.assignedElements()).toEqual([pane]);
	});

	it('applies each edit at once in live mode and leaves reset and cancel to the host', async () => {
		const { dialog, root, tab, field, action } = mount();
		dialog.live = true;
		dialog.defaults = { livePreview: true };
		dialog.categories = [
			...CATEGORIES.slice(0, 1),
			{
				id: 'fonts',
				label: 'Fonts',
				sections: [{ id: 'f', title: 'Fonts', controls: [], slot: 'block-fonts' }],
			},
		];
		const block = document.createElement('div');
		block.slot = 'block-fonts';
		dialog.append(block);
		const input = vi.fn();
		const reset = vi.fn();
		const cancelled = vi.fn();
		const changed = vi.fn();
		const category = vi.fn();
		dialog.addEventListener('office-option-input', (event) => input((event as CustomEvent).detail));
		dialog.addEventListener('office-options-reset', reset);
		dialog.addEventListener('office-options-cancel', cancelled);
		dialog.addEventListener('office-options-change', changed);
		dialog.addEventListener('office-options-category', (event) =>
			category((event as CustomEvent).detail),
		);
		dialog.show();
		await dialog.updateComplete;
		field<HTMLInputElement>('livePreview').click();
		expect(input).toHaveBeenCalledWith({ key: 'livePreview', value: false });
		expect(dialog.values.livePreview).toBe(false);
		action('reset').click();
		expect(reset).toHaveBeenCalledTimes(1);
		expect(dialog.values.livePreview).toBe(false);
		tab('Fonts').click();
		await dialog.updateComplete;
		expect(category).toHaveBeenLastCalledWith({ id: 'fonts' });
		expect(
			root.querySelector<HTMLSlotElement>('slot[name="block-fonts"]')!.assignedElements(),
		).toEqual([block]);
		// Closing from the title bar keeps the edits; only the Cancel button reports a cancel.
		root.querySelector('office-ui-dialog')!.dispatchEvent(new CustomEvent('office-dialog-close'));
		expect(changed).toHaveBeenCalledTimes(1);
		expect(cancelled).not.toHaveBeenCalled();
		dialog.show();
		await dialog.updateComplete;
		action('cancel').click();
		expect(cancelled).toHaveBeenCalledTimes(1);
	});

	it('lets Tab walk through a slotted pane instead of the dialog trap', async () => {
		const { dialog } = mount();
		dialog.categories = [
			{ id: 'ribbon', label: 'Customize Ribbon', sections: [], custom: true },
			{ id: 'other', label: 'Other', sections: [] },
		];
		const field = document.createElement('input');
		field.slot = 'pane-ribbon';
		dialog.append(field);
		dialog.show();
		await dialog.updateComplete;
		const tabKey = new KeyboardEvent('keydown', {
			key: 'Tab',
			bubbles: true,
			composed: true,
			cancelable: true,
		});
		field.dispatchEvent(tabKey);
		expect(tabKey.defaultPrevented).toBe(false);
	});

	it('clamps option numbers', () => {
		expect(clampOptionNumber('5', 1, 10)).toBe(5);
		expect(clampOptionNumber('0', 1, 10)).toBe(1);
		expect(clampOptionNumber('', 1, 10)).toBeUndefined();
		expect(clampOptionNumber('x', 1, 10)).toBeUndefined();
	});
});
