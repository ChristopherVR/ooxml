// @vitest-environment jsdom
import { createApp, h, nextTick, ref } from 'vue';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { XlsxEditorElement } from 'xlsx-web-component';
import { SpreadsheetEditor } from './vue';
import { emit, fakeWorkbook } from './test-support';

afterEach(() => document.body.replaceChildren());

describe('Vue editor adapter', () => {
	it('mounts once, forwards reactive props, re-emits kebab-case events, and unmounts', async () => {
		const container = document.createElement('div');
		document.body.append(container);
		const readOnly = ref(false);
		const onChange = vi.fn();
		const onSelection = vi.fn();
		const onReady = vi.fn();
		const exposed = ref<{ readonly element?: XlsxEditorElement }>();
		const app = createApp({
			render: () =>
				h(SpreadsheetEditor, {
					ref: exposed,
					readOnly: readOnly.value,
					locale: 'es-MX',
					showToolbar: false,
					'onWorkbook-change': onChange,
					'onSelection-change': onSelection,
					onReady,
				}),
		});
		app.mount(container);
		const editor = container.querySelector('xlsx-editor') as XlsxEditorElement;
		expect(editor.locale).toBe('es');
		expect(editor.showToolbar).toBe(false);
		expect(editor.showFormulaBar).toBe(true);
		expect(onReady).toHaveBeenCalledWith(editor);
		expect(exposed.value?.element).toBe(editor);
		readOnly.value = true;
		await nextTick();
		expect(editor.readOnly).toBe(true);
		expect(container.querySelectorAll('xlsx-editor')).toHaveLength(1);
		const workbook = fakeWorkbook();
		emit(editor, 'workbook-change', { workbook });
		emit(editor, 'selection-change', { sheet: 0, ref: 'A1', active: 'A1' });
		expect(onChange).toHaveBeenCalledWith(workbook);
		expect(onSelection).toHaveBeenCalledWith({ sheet: 0, ref: 'A1', active: 'A1' });
		app.unmount();
		expect(container.querySelector('xlsx-editor')).toBeNull();
	});
});
