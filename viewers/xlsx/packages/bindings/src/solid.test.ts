// @vitest-environment jsdom
import { createRoot, createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { XlsxEditorElement } from 'xlsx-web-component';
import type { EditorHandle } from './index';
import { SpreadsheetEditor } from './solid';
import { emit, fakeWorkbook } from './test-support';

afterEach(() => document.body.replaceChildren());

describe('Solid editor adapter', () => {
	it('mounts once, reacts to option changes, forwards events, and tears down', () => {
		let dispose!: () => void;
		const [readOnly, setReadOnly] = createSignal(false);
		const [className, setClassName] = createSignal('initial');
		const onWorkbookChange = vi.fn();
		const onReady = vi.fn();
		let handle: EditorHandle | undefined;
		const host = document.createElement('div');
		document.body.append(host);
		const props = {
			get readOnly() {
				return readOnly();
			},
			get class() {
				return className();
			},
			locale: 'fr-FR',
			showFormulaBar: false,
			onWorkbookChange,
			onReady,
			editorRef: (next: EditorHandle) => (handle = next),
		};
		createRoot((stop) => {
			dispose = stop;
			render(() => SpreadsheetEditor(props), host);
		});
		const editor = host.querySelector('xlsx-editor') as XlsxEditorElement;
		expect(editor).toBeDefined();
		expect(editor.readOnly).toBe(false);
		expect(editor.locale).toBe('fr');
		expect(editor.showFormulaBar).toBe(false);
		expect(handle?.element).toBe(editor);
		expect(onReady).toHaveBeenCalledWith(editor);
		setReadOnly(true);
		setClassName('updated');
		expect(editor.readOnly).toBe(true);
		expect(host.firstElementChild?.className).toBe('updated');
		const workbook = fakeWorkbook();
		emit(editor, 'workbook-change', { workbook });
		expect(onWorkbookChange).toHaveBeenCalledWith(workbook);
		dispose();
		expect(host.querySelector('xlsx-editor')).toBeNull();
	});

	it('keeps in-editor changes when other props change and reports them', () => {
		let dispose!: () => void;
		const [author, setAuthor] = createSignal('Ada');
		const onReadOnlyChange = vi.fn();
		const onRibbonCustomize = vi.fn();
		const host = document.createElement('div');
		document.body.append(host);
		const props = {
			readOnly: false,
			get authorName() {
				return author();
			},
			onReadOnlyChange,
			onRibbonCustomize,
		};
		createRoot((stop) => {
			dispose = stop;
			render(() => SpreadsheetEditor(props), host);
		});
		const editor = host.querySelector('xlsx-editor') as XlsxEditorElement;
		editor.readOnly = true;
		expect(onReadOnlyChange).toHaveBeenCalledWith(true);
		emit(editor, 'ribbon-customize', { hiddenActions: ['italic'] });
		expect(onRibbonCustomize).toHaveBeenCalledWith(['italic']);
		setAuthor('Grace');
		expect(editor.authorName).toBe('Grace');
		expect(editor.readOnly).toBe(true);
		dispose();
	});
});
