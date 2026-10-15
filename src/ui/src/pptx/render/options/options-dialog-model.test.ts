// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { registerPptxWebControls } from '../../web-components';
import { resolveCustomization } from '../customization/customization-resolve';
import {
	bindPptxOptionsDialog,
	type PptxOptionsDialogElement,
	type PptxOptionsDialogHost,
} from './options-dialog-binding';
import {
	AI_OPTIONS_CATEGORY,
	buildPptxOptionCategories,
	flattenViewerOptions,
	optionsBlockSlot,
	optionsPaneSlot,
	parseOptionValueKey,
	pptxOptionsLabels,
} from './options-dialog-model';
import { DEFAULT_VIEWER_OPTIONS } from './viewer-options';
import { VIEWER_OPTIONS_TABS } from './viewer-options-schema';
import { customizeOptionsTabs } from '../customization/customization-surfaces';

const t = (key: string): string => `<${key}>`;

describe('pptx options categories', () => {
	const categories = buildPptxOptionCategories(VIEWER_OPTIONS_TABS, t, false);

	it('keeps PowerPoint order and the ten category ids', () => {
		expect(categories.map((entry) => entry.id)).toEqual([
			'general',
			'proofing',
			'save',
			'language',
			'accessibility',
			'advanced',
			'ribbon',
			'quickAccess',
			'addIns',
			'trust',
		]);
		expect(categories[0]?.label).toBe('<pptx.settings.general>');
	});

	it('renders every schema control once, keyed group.key, with translated text', () => {
		const advanced = categories.find((entry) => entry.id === 'advanced')!;
		const controls = advanced.sections.flatMap((section) => section.controls);
		const recent = controls.find((entry) => entry.key === 'advanced.recentPresentationsCount');
		expect(recent).toMatchObject({ kind: 'number', label: expect.stringContaining('<pptx.') });
		// The ScreenTip style shows in General and Accessibility, as in PowerPoint; one pane each.
		const keys = categories.flatMap((entry) =>
			entry.sections.flatMap((section) => section.controls.map((control) => control.key)),
		);
		for (const entry of categories) {
			const own = entry.sections.flatMap((section) => section.controls.map((c) => c.key));
			expect(new Set(own).size).toBe(own.length);
		}
		const schemaCount = VIEWER_OPTIONS_TABS.flatMap((tab) => tab.sections).flatMap(
			(section) => section.controls,
		).length;
		expect(keys).toHaveLength(schemaCount);
	});

	it('hands bespoke blocks and panes to slots', () => {
		const slotOf = (category: string, section: string) =>
			categories.find((entry) => entry.id === category)?.sections.find((s) => s.id === section)
				?.slot;
		expect(slotOf('general', 'appearance')).toBe(optionsBlockSlot('themePicker'));
		expect(slotOf('general', 'fonts')).toBe(optionsBlockSlot('customFonts'));
		expect(slotOf('save', 'cache')).toBe(optionsBlockSlot('clearCache'));
		expect(slotOf('language', 'displayLanguage')).toBe(optionsBlockSlot('language'));
		for (const id of ['ribbon', 'addIns', 'quickAccess']) {
			expect(categories.find((entry) => entry.id === id)?.custom).toBe(true);
		}
		expect(categories.find((entry) => entry.id === 'ribbon')?.sections).toEqual([]);
		expect(categories.find((entry) => entry.id === 'quickAccess')?.sections).not.toEqual([]);
		expect(optionsPaneSlot('ribbon')).toBe('pane-ribbon');
	});

	it('appends the AI category only when asked', () => {
		expect(buildPptxOptionCategories(VIEWER_OPTIONS_TABS, t, true).at(-1)?.id).toBe(
			AI_OPTIONS_CATEGORY,
		);
	});

	it('shows host-locked settings disabled with the reason', () => {
		const resolved = resolveCustomization({
			options: { locked: { 'trust.openInProtectedView': true } },
		} as never);
		const tabs = customizeOptionsTabs(VIEWER_OPTIONS_TABS, resolved);
		const trust = buildPptxOptionCategories(tabs, t, false).find((entry) => entry.id === 'trust');
		const locked = trust?.sections
			.flatMap((section) => section.controls)
			.find((control) => control.key === 'trust.openInProtectedView');
		expect(locked?.disabled).toBe('<pptx.options.lockedByHost>');
	});

	it('flattens primitives only and parses keys back', () => {
		const values = flattenViewerOptions(DEFAULT_VIEWER_OPTIONS);
		expect(values['advanced.recentPresentationsCount']).toBe(50);
		expect(values['save.autoSave']).toBe(true);
		expect(Object.keys(values)).not.toContain('ribbon.hiddenTabIds');
		expect(parseOptionValueKey('advanced.recentPresentationsCount')).toEqual({
			group: 'advanced',
			key: 'recentPresentationsCount',
		});
		expect(parseOptionValueKey('nogroup')).toBeNull();
		expect(pptxOptionsLabels(t).ok).toBe('<pptx.common.ok>');
	});
});

describe('bindPptxOptionsDialog', () => {
	beforeAll(() => registerPptxWebControls());
	afterEach(() => document.body.replaceChildren());

	function mount() {
		const dialog = document.createElement('pptx-ui-options-dialog') as PptxOptionsDialogElement;
		document.body.append(dialog);
		const host = {
			setOption: vi.fn(),
			restoreOptions: vi.fn(),
			resetOptions: vi.fn(),
			close: vi.fn(),
			categoryChanged: vi.fn(),
		} satisfies PptxOptionsDialogHost;
		const binding = bindPptxOptionsDialog(dialog, host);
		const input = {
			options: DEFAULT_VIEWER_OPTIONS,
			t,
			customization: resolveCustomization(undefined),
			aiEnabled: false,
		};
		binding.sync(input);
		return { dialog, host, binding, input };
	}

	it('opens live and applies each edit through the host', async () => {
		const { dialog, host } = mount();
		await (dialog as unknown as { updateComplete: Promise<unknown> }).updateComplete;
		const toggle = dialog.shadowRoot!.querySelector<HTMLInputElement>(
			'[data-key="general.showMiniToolbar"] input',
		)!;
		toggle.click();
		expect(host.setOption).toHaveBeenCalledWith('general', 'showMiniToolbar', false);
		expect(host.categoryChanged).toHaveBeenCalledWith('general');
	});

	it('restores the opening snapshot on Cancel and closes on OK', async () => {
		const { dialog, host } = mount();
		await (dialog as unknown as { updateComplete: Promise<unknown> }).updateComplete;
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="cancel"]')!.click();
		expect(host.restoreOptions).toHaveBeenCalledWith(DEFAULT_VIEWER_OPTIONS);
		expect(host.close).toHaveBeenCalledTimes(1);
		dialog.show();
		await (dialog as unknown as { updateComplete: Promise<unknown> }).updateComplete;
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="ok"]')!.click();
		expect(host.close).toHaveBeenCalledTimes(2);
		expect(host.restoreOptions).toHaveBeenCalledTimes(1);
	});

	it('resets through the host and re-labels when the translator changes', async () => {
		const { dialog, host, binding, input } = mount();
		await (dialog as unknown as { updateComplete: Promise<unknown> }).updateComplete;
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="reset"]')!.click();
		expect(host.resetOptions).toHaveBeenCalledTimes(1);
		binding.sync({ ...input, t: (key) => `[${key}]` });
		expect(dialog.categories[0]?.label).toBe('[pptx.settings.general]');
		expect(dialog.heading).toBe('[pptx.options.title]');
	});

	it('stops listening once disposed', async () => {
		const { dialog, host, binding } = mount();
		await (dialog as unknown as { updateComplete: Promise<unknown> }).updateComplete;
		binding.dispose();
		dialog.shadowRoot!.querySelector<HTMLButtonElement>('[data-action="reset"]')!.click();
		expect(host.resetOptions).not.toHaveBeenCalled();
	});
});
