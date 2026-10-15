import { mount } from '@vue/test-utils';
import { createViewerOptionsStore, THEME_CATALOG, VIEWER_OPTIONS_TABS } from 'ooxml-ui/pptx';
import { LOCALE_CATALOG } from 'ooxml-ui/pptx/i18n';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { nextTick } from 'vue';

import SettingsDialog from './SettingsDialog.vue';

beforeEach(() => localStorage.clear());
afterEach(() => {
	document.body.innerHTML = '';
});

async function mountDialog() {
	const store = createViewerOptionsStore();
	const onRestoreOptions = vi.fn();
	const wrapper = mount(SettingsDialog, {
		props: {
			open: true,
			options: store.getOptions(),
			onOptionChange: (group, key, value) => store.setValue(group, key, value),
			onRestoreOptions,
			onRibbonTabHiddenChange: (tabId, hidden) => store.setRibbonTabHidden(tabId, hidden),
			onQuickAccessCommandsChange: (ids) => store.setQuickAccessCommands(ids),
			onResetOptions: (group) => store.reset(group),
			onClearCache: vi.fn(),
			themeKey: 'default',
			onThemeSelect: vi.fn(),
			localeCode: 'en',
			onLocaleSelect: vi.fn(),
			availableThemes: THEME_CATALOG,
			availableLocales: LOCALE_CATALOG,
		},
		attachTo: document.body,
	});
	// The shared dialog is bound once its element has mounted.
	await nextTick();
	await nextTick();
	return { wrapper, store, onRestoreOptions };
}

function shadow(): ShadowRoot {
	return document.body.querySelector('pptx-ui-options-dialog')!.shadowRoot!;
}

function tabs(): HTMLButtonElement[] {
	return Array.from(shadow().querySelectorAll<HTMLButtonElement>('nav [role="tab"]'));
}

function clickTab(label: string): void {
	const tab = tabs().find((b) => b.textContent?.trim() === label);
	if (!tab) {
		throw new Error(`category "${label}" not found`);
	}
	tab.click();
}

function action(name: string): HTMLButtonElement {
	return shadow().querySelector<HTMLButtonElement>(`[data-action="${name}"]`)!;
}

describe('settingsDialog (File > Options)', () => {
	it('renders all ten PowerPoint categories in the rail', async () => {
		await mountDialog();
		const labels = tabs().map((b) => b.textContent?.trim());
		expect(labels).toHaveLength(VIEWER_OPTIONS_TABS.length);
		expect(labels).toContain('Advanced');
		expect(labels).toContain('Trust Center');
	});

	it('writes a toggled option into the store', async () => {
		const { wrapper, store } = await mountDialog();
		clickTab('Advanced');
		await wrapper.vm.$nextTick();
		shadow().querySelector<HTMLInputElement>('[data-key="advanced.showGrid"] input')!.click();
		await wrapper.vm.$nextTick();
		expect(store.getOptions().advanced.showGrid).toBeTruthy();
	});

	it('shows the keyboard shortcut reference on the Customize Ribbon pane', async () => {
		const { wrapper } = await mountDialog();
		clickTab('Customize Ribbon');
		await wrapper.vm.$nextTick();
		expect(document.body.textContent).toContain('Ctrl/Cmd+Z');
	});

	it('restores the opening snapshot on Cancel', async () => {
		const { wrapper, onRestoreOptions } = await mountDialog();
		action('cancel').click();
		await wrapper.vm.$nextTick();
		expect(onRestoreOptions).toHaveBeenCalledOnce();
		expect(wrapper.emitted('close')).toHaveLength(1);
	});

	it('emits close from the OK button', async () => {
		const { wrapper } = await mountDialog();
		action('ok').click();
		await wrapper.vm.$nextTick();
		expect(wrapper.emitted('close')).toHaveLength(1);
	});
});
