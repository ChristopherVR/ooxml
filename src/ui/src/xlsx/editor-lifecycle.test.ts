// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { parseRange } from 'ooxml-core/xlsx';
import {
	loadEditorStrings,
	registerRibbonTabs,
	resetRibbonTabs,
	ribbonTabs,
} from 'ooxml-core/xlsx/ui';
import type { EditorCore } from './editor-core';
import { mountPopover } from './ribbon/popover';
import { defineXlsxEditor, XlsxEditorElement } from './component';
import { defineElement } from '../registry';

const panel = (editor: Element, id: string) =>
	editor.shadowRoot!.querySelector(`.ribbon-panel[data-tab="${id}"]`);
const tab = (id: string) => ({ id, label: id, groups: [] });

describe('<xlsx-editor> lifecycle', () => {
	beforeAll(() => defineXlsxEditor());

	let builtIn: ReturnType<typeof ribbonTabs> = [];
	afterEach(() => {
		document.body.replaceChildren();
		resetRibbonTabs();
		registerRibbonTabs([...builtIn]);
	});

	function mount() {
		const editor = document.createElement('xlsx-editor');
		document.body.append(editor);
		builtIn = [...ribbonTabs()];
		return editor;
	}

	it('follows ribbon tab registrations only while connected', () => {
		const editor = mount();
		registerRibbonTabs([tab('while-connected')]);
		expect(panel(editor, 'while-connected')).not.toBeNull();
		editor.remove();
		registerRibbonTabs([tab('while-removed')]);
		expect(panel(editor, 'while-removed')).toBeNull();
	});

	it('shows tabs registered while it was removed once reconnected', () => {
		const editor = mount();
		editor.remove();
		registerRibbonTabs([tab('while-removed')]);
		document.body.append(editor);
		expect(panel(editor, 'while-removed')).not.toBeNull();
		registerRibbonTabs([tab('after-return')]);
		expect(panel(editor, 'after-return')).not.toBeNull();
	});

	it('keeps its ribbon when it reconnects with nothing new to show', () => {
		const editor = mount();
		const before = editor.shadowRoot!.querySelector('.ribbon-panel');
		editor.remove();
		document.body.append(editor);
		expect(editor.shadowRoot!.querySelector('.ribbon-panel')).toBe(before);
	});

	it('relabels the mounted shell when the selected locale finishes loading', async () => {
		const editor = mount();
		editor.newWorkbook();
		const root = editor.shadowRoot!;
		const ribbon = root.querySelector('office-ui-ribbon')!;
		const backstage = root.querySelector('office-ui-backstage')!;
		const actions = root.querySelector('office-ui-ribbon-actions')!;
		const comments = () => actions.shadowRoot!.querySelector('[part="comments"]')?.textContent;
		expect(ribbon.getAttribute('file-label')).toBe('File');
		expect(comments()).toBe('Comments');
		editor.locale = 'fr-CA';
		const strings = await loadEditorStrings('fr');
		expect(ribbon.getAttribute('file-label')).toBe('Fichier');
		expect(backstage.getAttribute('back-label')).toBe('Retour au classeur');
		expect(comments()).toBe(strings.Comments);
		expect(comments()).not.toBe('Comments');
	});

	it('applies properties set before the element was defined', () => {
		const tag = 'xlsx-editor-early-read-only';
		const early = document.createElement(tag) as XlsxEditorElement;
		early.readOnly = true;
		early.fileName = 'early.xlsx';
		document.body.append(early);
		defineElement(customElements, tag, class extends XlsxEditorElement {});
		builtIn = [...ribbonTabs()];
		expect(Object.hasOwn(early, 'readOnly')).toBe(false);
		expect(early.readOnly).toBe(true);
		expect(early.fileName).toBe('early.xlsx');
		expect(early.hasAttribute('read-only')).toBe(true);
	});
	it('builds only the grid for a surfaceOnly set before the definition', () => {
		const tag = 'xlsx-editor-early-surface';
		const early = document.createElement(tag) as XlsxEditorElement;
		early.surfaceOnly = true;
		document.body.append(early);
		defineElement(customElements, tag, class extends XlsxEditorElement {});
		builtIn = [...ribbonTabs()];
		expect(early.surfaceOnly).toBe(true);
		expect(early.shadowRoot!.querySelector('[part~="grid"]')).not.toBeNull();
		expect(early.shadowRoot!.querySelector('[part~="ribbon"]')).toBeNull();
		expect(() => (early.surfaceOnly = false)).toThrow(/before mounting/);
	});

	it('keeps an early readOnly over the attribute it was upgraded with', () => {
		const tag = 'xlsx-editor-early-attribute';
		const early = document.createElement(tag) as XlsxEditorElement;
		early.setAttribute('read-only', 'false');
		early.readOnly = true;
		document.body.append(early);
		defineElement(customElements, tag, class extends XlsxEditorElement {});
		builtIn = [...ribbonTabs()];
		expect(early.readOnly).toBe(true);
		expect(early.getAttribute('read-only')).toBe('');
		early.removeAttribute('read-only');
		expect(early.readOnly).toBe(false);
	});

	it('closes an open filter menu when removed', () => {
		const editor = mount();
		editor.newWorkbook();
		const core = (editor as unknown as { core: EditorCore }).core;
		core.session!.setRangeValues(0, { row: 0, col: 0 }, [['Kind'], ['a'], ['b']]);
		core.session!.setAutoFilter(0, parseRange('A1:A3')!);
		const menu = () => editor.shadowRoot!.querySelector('.xg-fmenu');
		editor
			.shadowRoot!.querySelector('.xg-sink-input')!
			.dispatchEvent(
				new KeyboardEvent('keydown', { key: 'ArrowDown', altKey: true, bubbles: true }),
			);
		expect(menu()).not.toBeNull();
		editor.remove();
		expect(menu()).toBeNull();
	});

	it('closes only its own ribbon popover when removed', () => {
		const [first, second] = [mount(), mount()];
		const anchor = second.shadowRoot!.querySelector<HTMLElement>('button')!;
		const pop = document.createElement('div');
		expect(mountPopover(anchor, pop)).toBe(true);
		first.remove();
		expect(pop.isConnected).toBe(true);
		second.remove();
		expect(pop.isConnected).toBe(false);
	});
});
