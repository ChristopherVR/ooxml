// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { registerRibbonTabs, resetRibbonTabs, ribbonTabs } from 'ooxml-core/xlsx/ui';
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
});
