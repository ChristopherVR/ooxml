import { mount } from '@vue/test-utils';
import { VIEWER_SHORTCUT_REFERENCE } from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it } from 'vitest';

import { translationsEn } from '../../i18n';
import ShortcutPanel from './ShortcutPanel.vue';

// ShortcutPanel teleports its body to <body> via ModalDialog, so query the
// document rather than the wrapper, and clean up between tests.
afterEach(() => {
	document.body.replaceChildren();
});

function mountPanel(open = true) {
	return mount(ShortcutPanel, {
		props: { open },
		attachTo: document.body,
	});
}

/**
 * The panel teleports its body to `<body>`, so the search `<input>` is outside
 * the component's own DOM tree and `wrapper.find` cannot reach it. Drive it
 * directly via the document, then await the wrapper's reactivity tick.
 */
async function typeSearch(wrapper: ReturnType<typeof mountPanel>, value: string): Promise<void> {
	const input = document.querySelector('.pptx-vue-shortcuts-search') as HTMLInputElement;
	input.value = value;
	input.dispatchEvent(new Event('input'));
	await wrapper.vm.$nextTick();
}

describe('shortcutPanel', () => {
	it('renders nothing when closed', () => {
		mountPanel(false);
		expect(document.querySelector('[data-pptx-shortcuts-panel]')).toBeNull();
	});

	it('renders every shortcut under its heading', () => {
		mountPanel(true);
		const rows = document.querySelectorAll('.pptx-vue-shortcuts-row');
		expect(rows).toHaveLength(VIEWER_SHORTCUT_REFERENCE.length);
		// Group headings are present.
		const titles = Array.from(document.querySelectorAll('.pptx-vue-shortcuts-group-title')).map(
			(el) => el.textContent,
		);
		expect(titles).toContain('History');
		expect(titles).toContain('Clipboard');
		expect(titles).toContain('Editing');
		expect(titles).toContain('Navigation');
		expect(titles).toContain('General');
	});

	it('lists every row of the shared shortcut reference, the keymap the other panels show', () => {
		mountPanel(true);
		const rows = Array.from(document.querySelectorAll('.pptx-vue-shortcuts-row')).map((row) => ({
			description: row.querySelector('.pptx-vue-shortcuts-desc')?.textContent?.trim() ?? '',
			combo: row.querySelector('.pptx-vue-shortcuts-combo')?.textContent?.trim() ?? '',
		}));
		expect(rows).toHaveLength(VIEWER_SHORTCUT_REFERENCE.length);
		expect(new Set(rows.map((row) => row.description))).toStrictEqual(
			new Set(VIEWER_SHORTCUT_REFERENCE.map((item) => translationsEn[item.actionKey])),
		);
		// Ctrl/Cmd+Enter commits a text edit; Vue's own catalog never listed it.
		expect(rows).toContainEqual({
			description: translationsEn['pptx.shortcuts.action.commitTextEdit'],
			combo: 'Ctrl+Enter',
		});
	});

	it('renders a platform-aware combo glyph (Ctrl on non-mac)', () => {
		mountPanel(true);
		const combos = Array.from(document.querySelectorAll('.pptx-vue-shortcuts-combo')).map(
			(el) => el.textContent ?? '',
		);
		// In the (non-mac) test environment, Mod renders as "Ctrl".
		expect(combos.some((text) => text.includes('Ctrl'))).toBeTruthy();
	});

	it('filters shortcuts by the search query', async () => {
		const wrapper = mountPanel(true);
		await typeSearch(wrapper, 'undo');
		const rows = document.querySelectorAll('.pptx-vue-shortcuts-row');
		// Only "Undo" matches; assert at least one and fewer than the full reference.
		expect(rows.length).toBeGreaterThanOrEqual(1);
		expect(rows.length).toBeLessThan(VIEWER_SHORTCUT_REFERENCE.length);
		const desc = document.querySelector('.pptx-vue-shortcuts-desc')?.textContent ?? '';
		expect(desc.toLowerCase()).toContain('undo');
	});

	it('shows an empty message when nothing matches', async () => {
		const wrapper = mountPanel(true);
		await typeSearch(wrapper, 'zzzznotathing');
		expect(document.querySelector('.pptx-vue-shortcuts-empty')).not.toBeNull();
		expect(document.querySelectorAll('.pptx-vue-shortcuts-row')).toHaveLength(0);
	});

	it('emits close from the ModalDialog close button', async () => {
		const wrapper = mountPanel(true);
		const closeBtn = document.querySelector('.pptx-vue-modal-close') as HTMLButtonElement | null;
		expect(closeBtn).not.toBeNull();
		closeBtn?.click();
		await wrapper.vm.$nextTick();
		expect(wrapper.emitted('close')).toBeTruthy();
	});
});
