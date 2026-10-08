// @vitest-environment jsdom
// Every framework forwards `collaboration`, re-emits `collaboration-change`, exposes `share()` and
// leaves the room when the editor unmounts. Joining itself is the element's job (ooxml-ui tests).
import '@angular/compiler';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot } from 'solid-js';
import { render } from 'solid-js/web';
import { flushSync, mount, unmount } from 'svelte';
import { createApp, h, ref } from 'vue';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { XlsxCollaborationState, XlsxEditorElement } from 'xlsx-web-component';
import { SpreadsheetEditorComponent } from './angular';
import { mountEditor, type EditorHandle } from './index';
import { SpreadsheetEditor as ReactEditor } from './react';
import { SpreadsheetEditor as SolidEditor } from './solid';
import { emit } from './test-support';
import { SpreadsheetEditor as VueEditor } from './vue';
import XlsxEditor from './XlsxEditor.svelte';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

const options = { roomId: 'budget-2026', user: { name: 'Ada', color: '#2563eb' } };
const state: XlsxCollaborationState = {
	active: true,
	roomId: 'budget-2026',
	status: 'connected',
	synced: true,
	people: [{ clientId: 1, name: 'Ada', color: '#2563eb', self: true }],
};

function host() {
	const node = document.createElement('div');
	document.body.append(node);
	return node;
}

/** The shared assertions once a framework has mounted the editor with `options`. */
function check(container: Element, onChange: ReturnType<typeof vi.fn>, share: () => void) {
	const editor = container.querySelector('xlsx-editor') as XlsxEditorElement;
	expect(editor.collaboration).toBe(options);
	emit(editor, 'collaboration-change', state);
	expect(onChange).toHaveBeenCalledWith(state);
	const opened = vi.fn();
	editor.share = opened;
	share();
	expect(opened).toHaveBeenCalledTimes(1);
	return editor;
}

describe('collaboration in every binding', () => {
	it('vanilla (mountEditor)', () => {
		const onCollaborationChange = vi.fn();
		const node = host();
		const binding = mountEditor(node, { collaboration: options, onCollaborationChange });
		const editor = check(node, onCollaborationChange, () => binding.share());
		// The same object again is not a change; null leaves the room.
		const stop = vi.spyOn(editor, 'stopCollaboration');
		binding.update({ collaboration: options, onCollaborationChange });
		expect(editor.collaboration).toBe(options);
		binding.update({ collaboration: null, onCollaborationChange });
		expect(editor.collaboration).toBeNull();
		binding.update({ collaboration: options, onCollaborationChange });
		binding.destroy();
		expect(stop).toHaveBeenCalled();
		expect(editor.collaboration).toBeNull();
	});

	it('React', () => {
		const node = host();
		const root = createRoot(node);
		const handle = createRef<EditorHandle>();
		const onCollaborationChange = vi.fn();
		act(() =>
			root.render(
				createElement(ReactEditor, { ref: handle, collaboration: options, onCollaborationChange }),
			),
		);
		const editor = check(node, onCollaborationChange, () => handle.current!.share());
		act(() => root.unmount());
		expect(editor.collaboration).toBeNull();
	});

	it('Vue', () => {
		const node = host();
		const editorRef = ref<EditorHandle>();
		const onCollaborationChange = vi.fn();
		const app = createApp({
			render: () => h(VueEditor, { ref: editorRef, collaboration: options, onCollaborationChange }),
		});
		app.mount(node);
		const editor = check(node, onCollaborationChange, () => editorRef.value!.share());
		app.unmount();
		expect(editor.collaboration).toBeNull();
	});

	it('Angular', async () => {
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const hostElement = document.createElement('spreadsheet-editor');
		document.body.append(hostElement);
		const component = createComponent(SpreadsheetEditorComponent, {
			environmentInjector: app.injector,
			hostElement,
		});
		component.setInput('collaboration', options);
		const onChange = vi.fn();
		component.instance.collaborationChange.subscribe(onChange);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		const editor = check(hostElement, onChange, () => component.instance.share());
		component.destroy();
		expect(editor.collaboration).toBeNull();
		app.destroy();
	});

	it('Svelte', () => {
		const target = host();
		const oncollaborationchange = vi.fn();
		const component = mount(XlsxEditor, {
			target,
			props: { collaboration: options, oncollaborationchange },
		});
		flushSync();
		const editor = check(target, oncollaborationchange, () => component.share());
		unmount(component);
		expect(editor.collaboration).toBeNull();
	});

	it('Solid', () => {
		const node = host();
		let handle: EditorHandle | undefined;
		let dispose!: () => void;
		const onCollaborationChange = vi.fn();
		createSolidRoot((stop) => {
			dispose = stop;
			render(
				() =>
					SolidEditor({
						collaboration: options,
						onCollaborationChange,
						editorRef: (value) => (handle = value),
					}),
				node,
			);
		});
		const editor = check(node, onCollaborationChange, () => handle!.share());
		dispose();
		expect(editor.collaboration).toBeNull();
	});
});
