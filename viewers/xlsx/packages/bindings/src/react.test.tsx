// @vitest-environment jsdom
import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { XlsxEditorElement } from 'xlsx-web-component';
import type { EditorHandle } from './index';
import { SpreadsheetEditor, type SpreadsheetEditorProps } from './react';
import { emit, fakeWorkbook } from './test-support';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

describe('React editor adapter', () => {
	it('mounts once, forwards prop changes and events, exposes the handle, and unmounts', () => {
		const container = document.createElement('div');
		document.body.append(container);
		const root = createRoot(container);
		const ref = createRef<EditorHandle>();
		const onWorkbookChange = vi.fn();
		const onDirtyChange = vi.fn();
		const render = (props: SpreadsheetEditorProps) =>
			act(() => root.render(createElement(SpreadsheetEditor, { ...props, ref })));
		render({ className: 'sheet', locale: 'de-DE', onWorkbookChange, onDirtyChange });
		const editor = container.querySelector('xlsx-editor') as XlsxEditorElement;
		expect(editor.parentElement?.className).toBe('sheet');
		expect(editor.locale).toBe('de');
		expect(ref.current?.element).toBe(editor);
		render({ className: 'sheet', readOnly: true, onWorkbookChange, onDirtyChange });
		expect(container.querySelectorAll('xlsx-editor')).toHaveLength(1);
		expect(editor.readOnly).toBe(true);
		expect(editor.locale).toBe('en');
		const workbook = fakeWorkbook();
		emit(editor, 'workbook-change', { workbook });
		emit(editor, 'dirty-change', { dirty: true });
		expect(onWorkbookChange).toHaveBeenCalledWith(workbook);
		expect(onDirtyChange).toHaveBeenCalledWith(true);
		act(() => root.unmount());
		expect(container.querySelector('xlsx-editor')).toBeNull();
		expect(() => ref.current?.element).not.toThrow();
	});
});
