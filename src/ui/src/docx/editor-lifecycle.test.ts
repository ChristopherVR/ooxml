// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createDocument, type DocumentModel } from 'ooxml-core/docx';
import { DocxEditorElement } from './component';
import { defineElement } from '../registry';
import { mountPopover } from './ribbon-popover';

function longDocument(paragraphs: number): DocumentModel {
	const model = createDocument();
	model.blocks = Array.from({ length: paragraphs }, (_, index) => ({
		type: 'paragraph' as const,
		id: `p${index}`,
		runs: [{ text: `Paragraph number ${index} with enough words to fill a line or two.` }],
	}));
	return model;
}

/** How many window listeners of `types` are attached right now. */
function trackWindowListeners(types: readonly string[]) {
	const active = new Set<unknown>();
	const add = window.addEventListener.bind(window);
	const remove = window.removeEventListener.bind(window);
	vi.spyOn(window, 'addEventListener').mockImplementation(((
		type: string,
		listener: EventListener,
		options?: AddEventListenerOptions,
	) => {
		if (types.includes(type)) active.add(listener);
		add(type, listener, options);
	}) as typeof window.addEventListener);
	vi.spyOn(window, 'removeEventListener').mockImplementation(((
		type: string,
		listener: EventListener,
		options?: EventListenerOptions,
	) => {
		if (types.includes(type)) active.delete(listener);
		remove(type, listener, options);
	}) as typeof window.removeEventListener);
	return active;
}

function showPrintLayout(editor: DocxEditorElement) {
	const view = editor.shadowRoot!.querySelector<HTMLSelectElement>('[aria-label="Layout view"]')!;
	view.value = 'print';
	view.dispatchEvent(new Event('change'));
	vi.advanceTimersByTime(500);
}

const pages = (editor: DocxEditorElement) => [
	...editor.shadowRoot!.querySelectorAll<HTMLElement>('.dve-canvas .dve-print-page'),
];

describe('<docx-editor> lifecycle', () => {
	afterEach(() => {
		document.body.replaceChildren();
		vi.useRealTimers();
		vi.restoreAllMocks();
	});

	it('holds no print listener while removed and takes them back on reconnect', () => {
		const active = trackWindowListeners(['beforeprint', 'afterprint']);
		const editor = document.createElement('docx-editor');
		document.body.append(editor);
		expect(active.size).toBe(2);
		editor.remove();
		expect(active.size).toBe(0);
		document.body.append(editor);
		expect(active.size).toBe(2);
	});

	it('closes an open context menu with the editor', () => {
		vi.useFakeTimers();
		const active = trackWindowListeners(['blur', 'resize']);
		const editor = document.createElement('docx-editor');
		document.body.append(editor);
		const surface = editor.shadowRoot!.querySelector('.ProseMirror')!;
		surface.dispatchEvent(
			new MouseEvent('contextmenu', { bubbles: true, composed: true, cancelable: true }),
		);
		vi.advanceTimersByTime(1);
		expect(editor.shadowRoot!.querySelector('office-ui-context-menu')).not.toBeNull();
		expect(active.size).toBeGreaterThan(0);
		editor.remove();
		expect(active.size).toBe(0);
		expect(editor.shadowRoot!.querySelector('office-ui-context-menu')).toBeNull();
	});

	it('still prints every page after a reconnect', () => {
		vi.useFakeTimers();
		const editor = document.createElement('docx-editor');
		editor.documentModel = longDocument(600);
		document.body.append(editor);
		showPrintLayout(editor);
		editor.remove();
		document.body.append(editor);
		vi.advanceTimersByTime(500);
		const sheets = pages(editor);
		expect(sheets.length).toBeGreaterThan(10);
		expect(sheets.some((sheet) => !sheet.childElementCount)).toBe(true);
		window.dispatchEvent(new Event('beforeprint'));
		expect(sheets.every((sheet) => sheet.childElementCount)).toBe(true);
		window.dispatchEvent(new Event('afterprint'));
		expect(sheets.some((sheet) => !sheet.childElementCount)).toBe(true);
	});

	it('applies a read-only flag set before the element was defined', () => {
		const tag = 'docx-editor-early-read-only';
		const early = document.createElement(tag) as DocxEditorElement;
		early.readOnly = true;
		early.fileName = 'set-early.docx';
		document.body.append(early);
		defineElement(customElements, tag, class extends DocxEditorElement {});
		expect(Object.hasOwn(early, 'readOnly')).toBe(false);
		expect(early.readOnly).toBe(true);
		expect(early.fileName).toBe('set-early.docx');
		const surface = early.shadowRoot!.querySelector('.ProseMirror')!;
		expect(surface.getAttribute('contenteditable')).toBe('false');
	});
	it('builds only the editing surface for a surfaceOnly set before the definition', () => {
		const tag = 'docx-editor-early-surface';
		const early = document.createElement(tag) as DocxEditorElement;
		early.surfaceOnly = true;
		document.body.append(early);
		defineElement(customElements, tag, class extends DocxEditorElement {});
		expect(early.surfaceOnly).toBe(true);
		expect(early.shadowRoot!.querySelector('.ProseMirror')).not.toBeNull();
		expect(early.shadowRoot!.querySelector('office-ui-ribbon')).toBeNull();
	});

	it('keeps an early locale over the locale attribute it was upgraded with', () => {
		const tag = 'docx-editor-early-locale';
		const early = document.createElement(tag) as DocxEditorElement;
		early.setAttribute('locale', 'de');
		early.locale = 'fr';
		document.body.append(early);
		defineElement(customElements, tag, class extends DocxEditorElement {});
		expect(early.locale).toBe('fr');
		expect(early.getAttribute('locale')).toBe('fr');
		early.setAttribute('locale', 'es');
		expect(early.locale).toBe('es');
	});

	it('closes only its own ribbon popover when removed', () => {
		const [first, second] = [0, 1].map(() => {
			const editor = document.createElement('docx-editor');
			document.body.append(editor);
			return editor;
		});
		const anchor = second!.shadowRoot!.querySelector<HTMLElement>('button')!;
		const pop = document.createElement('div');
		expect(mountPopover(anchor, pop, anchor)).toBe(true);
		first!.remove();
		expect(pop.isConnected).toBe(true);
		second!.remove();
		expect(pop.isConnected).toBe(false);
	});
});
