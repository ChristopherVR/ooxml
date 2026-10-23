// @vitest-environment jsdom
// Every framework adapter against `EDITOR_CONTRACT`: the same props in, every element event out
// under its handler, and the same handle members.
import '@angular/compiler';
import { reflectComponentType } from '@angular/core';
import {
	angularOutputName,
	handleKeys,
	optionNames,
	propKeys,
	svelteHandlerName,
} from 'ooxml-ui/bindings';
import { vueEmits } from 'ooxml-ui/bindings/vue';
import { act, createElement, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot } from 'solid-js';
import { render } from 'solid-js/web';
import { flushSync, mount, unmount } from 'svelte';
import { createApp, h, ref } from 'vue';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
	EDITOR_CALLBACKS,
	EDITOR_CONTRACT,
	EDITOR_EVENT_NAMES,
	EDITOR_HANDLE_KEYS,
	EDITOR_PROP_KEYS,
	type EditorHandle,
} from './index';
import { SpreadsheetEditor as ReactEditor } from './react';
import { SpreadsheetEditor as SolidEditor } from './solid';
import { SpreadsheetEditor as VueEditor } from './vue';
import XlsxEditor from './XlsxEditor.svelte';

const reactAct = globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean };
beforeEach(() => (reactAct.IS_REACT_ACT_ENVIRONMENT = true));
afterEach(() => {
	delete reactAct.IS_REACT_ACT_ENVIRONMENT;
	document.body.replaceChildren();
});
const sorted = (values: Iterable<string>) => [...values].sort();
const missing = (target: object) => EDITOR_HANDLE_KEYS.filter((key) => !(key in target));
const newHost = () => document.body.appendChild(document.createElement('div'));
const OPTIONS = [...optionNames(EDITOR_CONTRACT), ...EDITOR_CALLBACKS];
const spies = (nameOf: (option: string) => string = (option) => option) =>
	Object.fromEntries(OPTIONS.map((option) => [nameOf(option), vi.fn()]));
/** Dispatches every contract event on the editor; returns the handlers that ran. */
function fireAll(handlers: Record<string, ReturnType<typeof vi.fn>>): string[] {
	const editor = document.querySelector('xlsx-editor')!;
	const detail = { readOnly: false, hostReadOnly: false, warnings: [], hiddenActions: [] };
	for (const name of EDITOR_EVENT_NAMES) editor.dispatchEvent(new CustomEvent(name, { detail }));
	return Object.entries(handlers)
		.filter(([, handler]) => handler.mock.calls.length > 0)
		.map(([name]) => name);
}

describe('EDITOR_CONTRACT', () => {
	it('lists the shared keys, events and handle members exactly once', () => {
		expect(sorted(propKeys(EDITOR_CONTRACT))).toEqual(sorted(EDITOR_PROP_KEYS));
		expect(sorted(Object.keys(EDITOR_CONTRACT.events))).toEqual(sorted(EDITOR_EVENT_NAMES));
		expect(handleKeys(EDITOR_CONTRACT.handle)).toEqual([...EDITOR_HANDLE_KEYS]);
		expect(sorted(EDITOR_HANDLE_KEYS)).toEqual(
			sorted([
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
				'share',
			]),
		);
	});
});

describe('every framework adapter matches the contract', () => {
	it('React: every event and onReady reach their handler props; the ref is the handle', () => {
		const root = createRoot(newHost());
		const handle = createRef<EditorHandle>();
		const handlers = spies();
		act(() => root.render(createElement(ReactEditor, { ref: handle, ...handlers })));
		expect(missing(handle.current!)).toEqual([]);
		expect(handle.current!.element.localName).toBe('xlsx-editor');
		expect(sorted(fireAll(handlers))).toEqual(sorted(OPTIONS));
		act(() => root.unmount());
	});
	it('Vue: declares every prop, emits every event, v-model and ready, exposes the handle', () => {
		expect(sorted(Object.keys(VueEditor.props))).toEqual(sorted(EDITOR_PROP_KEYS));
		expect(VueEditor.emits).toEqual(vueEmits(EDITOR_CONTRACT, { ready: 'onReady' }));
		const editor = ref<EditorHandle>();
		const app = createApp({ render: () => h(VueEditor, { ref: editor }) });
		app.mount(newHost());
		expect(missing(editor.value!)).toEqual([]);
		expect(editor.value!.element.localName).toBe('xlsx-editor');
		app.unmount();
	});
	it('Solid: every event and onReady reach their handler props; editorRef gets the handle', () => {
		let handle: EditorHandle | undefined;
		const handlers = spies();
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(() => SolidEditor({ ...handlers, editorRef: (value) => (handle = value) }), newHost());
		});
		expect(missing(handle!)).toEqual([]);
		expect(handle!.element.localName).toBe('xlsx-editor');
		expect(sorted(fireAll(handlers))).toEqual(sorted(OPTIONS));
		dispose();
	});
	it('Angular: one input per prop, one output per event and ready, the handle on the class', async () => {
		const { SpreadsheetEditorComponent } = await import('./angular');
		const mirror = reflectComponentType(SpreadsheetEditorComponent)!;
		expect(sorted(mirror.inputs.map((input) => input.propName))).toEqual(sorted(EDITOR_PROP_KEYS));
		expect(sorted(mirror.outputs.map((output) => output.propName))).toEqual(
			sorted(OPTIONS.map(angularOutputName)),
		);
		expect(missing(SpreadsheetEditorComponent.prototype)).toEqual([]);
	});
	it('Svelte: every event and onready reach their lower-case props; exports the handle', () => {
		const target = newHost();
		const handlers = spies(svelteHandlerName);
		const component = mount(XlsxEditor, { target, props: handlers });
		flushSync();
		expect(missing(component)).toEqual([]);
		expect(component.element).toBe(target.querySelector('xlsx-editor'));
		expect(component.dirty).toBe(false);
		expect(sorted(fireAll(handlers))).toEqual(sorted(OPTIONS.map(svelteHandlerName)));
		// The deprecated aliases keep working.
		expect(component.getElement()).toBe(component.element);
		expect(component.isDirty()).toBe(false);
		unmount(component);
	});
});
