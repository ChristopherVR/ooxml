// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument } from 'ooxml-core/docx';
import { WORD_RIBBON_TAB_IDS } from 'ooxml-core/docx/ui';
import { DocxEditorElement, registerDocxEditor, type WordCustomization } from './index';
import { createRibbon } from './ribbon';
import { quickAccessButton, ribbonFile, ribbonShadow, ribbonTabs } from './test-support';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());

const NO_FILE: WordCustomization = { ribbon: { hiddenTabs: ['file'] } };

function mount(customization?: WordCustomization): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	const model = createDocument();
	model.warnings = ['Pictures are shown as placeholders.'];
	editor.documentModel = model;
	if (customization) editor.customization = customization;
	document.body.append(editor);
	return editor;
}
type Ribbon = HTMLElement & { noFile: boolean; selected: string; fileButton(): unknown };
const ribbon = (editor: DocxEditorElement) =>
	editor.shadowRoot!.querySelector('office-ui-ribbon') as Ribbon;
const backstage = (editor: DocxEditorElement) =>
	editor.shadowRoot!.querySelector('office-ui-backstage') as HTMLElement & { open: boolean };
/** The tab buttons the row draws, once the ribbon has read its panels again. */
const tabIds = async (editor: DocxEditorElement) => {
	await new Promise((resolve) => setTimeout(resolve));
	return (await ribbonTabs(ribbon(editor))).map((tab) => tab.dataset.tab);
};
const notesButton = async (editor: DocxEditorElement) => {
	const bar = editor.shadowRoot!.querySelector('office-ui-status-bar') as HTMLElement & {
		updateComplete: Promise<unknown>;
	};
	await bar.updateComplete;
	return bar.shadowRoot!.querySelector<HTMLButtonElement>('[data-id="notes"]')!;
};

describe('<docx-editor> ribbon.hiddenTabs', () => {
	it('names every native tab the ribbon builds', () => {
		const panels = [...createRibbon().children]
			.map((child) => (child as HTMLElement).dataset.ribbonTab)
			.filter(Boolean);
		expect(['file', ...panels]).toEqual([...WORD_RIBBON_TAB_IDS]);
	});

	it('removes File and its backstage while the other tabs stay', async () => {
		const editor = mount(NO_FILE);
		expect(ribbon(editor).noFile).toBe(true);
		expect((await ribbonShadow(ribbon(editor))).querySelector<HTMLElement>('.file')!.hidden).toBe(
			true,
		);
		expect(ribbon(editor).fileButton()).toBeNull();
		expect(await tabIds(editor)).toEqual([
			'home',
			'insert',
			'design',
			'layout',
			'references',
			'mailings',
			'review',
			'view',
			'help',
		]);
		// The File button's own event, Alt key tips and the status bar's notes all stay shut.
		ribbon(editor).dispatchEvent(new CustomEvent('office-ribbon-file'));
		expect(backstage(editor).open).toBe(false);
		ribbon(editor).dispatchEvent(new CustomEvent('dve-keytips'));
		const badges = [...(await ribbonShadow(ribbon(editor))).querySelectorAll('.dve-keytip')];
		expect(badges.map((badge) => badge.textContent)).not.toContain('F');
		ribbon(editor).dispatchEvent(new KeyboardEvent('keydown', { key: 'f' }));
		expect(backstage(editor).open).toBe(false);
		expect((await notesButton(editor)).hidden).toBe(true);
	});

	it('keeps Ctrl+S, Quick Access Save and file-command working', async () => {
		const editor = mount(NO_FILE);
		const commands = vi.fn((event: Event) => event.preventDefault());
		editor.addEventListener('file-command', commands);
		editor.dispatchEvent(
			new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true, cancelable: true }),
		);
		(await quickAccessButton(editor, 'Save')).click();
		expect(commands.mock.calls.map(([event]) => (event as CustomEvent).detail.command)).toEqual([
			'save',
			'save',
		]);
		expect(backstage(editor).open).toBe(false);
	});

	it('closes an open backstage when File is hidden at runtime and returns focus to the ribbon', async () => {
		const editor = mount();
		(await ribbonFile(ribbon(editor))).click();
		expect(backstage(editor).open).toBe(true);
		editor.customization = NO_FILE;
		expect(backstage(editor).open).toBe(false);
		const focused = ribbon(editor).shadowRoot!.activeElement as HTMLElement | null;
		expect(focused?.getAttribute('role')).toBe('tab');
		editor.customization = null;
		expect(ribbon(editor).noFile).toBe(false);
		expect((await notesButton(editor)).hidden).toBe(false);
		(await ribbonFile(ribbon(editor))).click();
		expect(backstage(editor).open).toBe(true);
	});

	it('hides other native tabs and moves the selection to the first one left', async () => {
		const editor = mount({ ribbon: { hiddenTabs: ['home', 'mailings'] } });
		expect(ribbon(editor).noFile).toBe(false);
		expect(await tabIds(editor)).toEqual([
			'insert',
			'design',
			'layout',
			'references',
			'review',
			'view',
			'help',
		]);
		expect(ribbon(editor).selected).toBe('insert');
		editor.customization = {};
		expect(await tabIds(editor)).toContain('home');
	});

	it('leaves commands that live only on a hidden tab out of Tell me', async () => {
		const editor = mount();
		const toolbar = ribbon(editor);
		const labels = (tab: string) =>
			[
				...toolbar.querySelectorAll<HTMLElement>(
					`[data-ribbon-tab="${tab}"] button[aria-label]:not([disabled])`,
				),
			].map((control) => control.getAttribute('aria-label')!);
		const elsewhere = new Set(
			[...toolbar.querySelectorAll<HTMLElement>('[data-ribbon-tab] button[aria-label]')]
				.filter((control) => !control.closest('[data-ribbon-tab="references"]'))
				.map((control) => control.getAttribute('aria-label')),
		);
		const only = labels('references').find((label) => !elsewhere.has(label))!;
		const shared = labels('references').find((label) => elsewhere.has(label));
		const bar = editor.shadowRoot!.querySelector('office-ui-title-bar') as HTMLElement & {
			state: { search: { match(query: string): { id: string }[] } };
		};
		const found = (query: string) => bar.state.search.match(query).map((entry) => entry.id);
		expect(found(only)).toContain(only);
		editor.customization = { ribbon: { hiddenTabs: ['references'] } };
		expect(found(only)).not.toContain(only);
		if (shared) expect(found(shared)).toContain(shared);
		editor.customization = null;
		expect(found(only)).toContain(only);
	});
});
