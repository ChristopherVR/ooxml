// @vitest-environment jsdom
import '@angular/compiler';
import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot } from 'solid-js';
import { render } from 'solid-js/web';
import { flushSync, mount, unmount } from 'svelte';
import { createApp, h, ref } from 'vue';
import { afterEach, describe, expect, it } from 'vitest';
import { EDITOR_HANDLE_KEYS, type EditorHandle } from './index';
import { SpreadsheetEditor as ReactEditor } from './react';
import { SpreadsheetEditor as SolidEditor } from './solid';
import { SpreadsheetEditor as VueEditor } from './vue';
import XlsxEditor from './XlsxEditor.svelte';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());
const missing = (target: object) => EDITOR_HANDLE_KEYS.filter((key) => !(key in target));

describe('every framework exposes the shared handle vocabulary', () => {
	it('lists the members of EditorHandle', () => {
		expect([...EDITOR_HANDLE_KEYS].sort()).toEqual(
			[
				'dirty',
				'download',
				'element',
				'getSelection',
				'load',
				'markClean',
				'newWorkbook',
				'save',
				'saveBytes',
				'select',
				'setActiveSheet',
			].sort(),
		);
	});
	it('React', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const handle = createRef<EditorHandle>();
		const root = createRoot(host);
		act(() => root.render(createElement(ReactEditor, { ref: handle })));
		expect(missing(handle.current!)).toEqual([]);
		expect(handle.current!.element.localName).toBe('xlsx-editor');
		act(() => root.unmount());
	});
	it('Vue', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const editor = ref<EditorHandle>();
		const app = createApp({ render: () => h(VueEditor, { ref: editor }) });
		app.mount(host);
		expect(missing(editor.value!)).toEqual([]);
		expect(editor.value!.element.localName).toBe('xlsx-editor');
		app.unmount();
	});
	it('Solid', () => {
		const host = document.createElement('div');
		document.body.append(host);
		let handle: EditorHandle | undefined;
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(() => SolidEditor({ editorRef: (value) => (handle = value) }), host);
		});
		expect(missing(handle!)).toEqual([]);
		expect(handle!.element.localName).toBe('xlsx-editor');
		dispose();
	});
	it('Angular', async () => {
		const { SpreadsheetEditorComponent } = await import('./angular');
		expect(missing(SpreadsheetEditorComponent.prototype)).toEqual([]);
	});
	it('Svelte', () => {
		const target = document.createElement('div');
		document.body.append(target);
		const component = mount(XlsxEditor, { target });
		flushSync();
		expect(missing(component)).toEqual([]);
		expect(component.element).toBe(target.querySelector('xlsx-editor'));
		expect(component.dirty).toBe(false);
		// The deprecated aliases keep working.
		expect(component.getElement()).toBe(component.element);
		expect(component.isDirty()).toBe(false);
		unmount(component);
	});
});
