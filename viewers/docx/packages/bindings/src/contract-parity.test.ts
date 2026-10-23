// @vitest-environment jsdom
// Every framework adapter against `EDITOR_CONTRACT`: the same props in, every element event out
// under its handler, and the same handle members.
import '@angular/compiler';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { reflectComponentType } from '@angular/core';
import {
	angularOutputName,
	handleKeys,
	optionNames,
	propKeys,
	svelteHandlerName,
} from 'ooxml-ui/bindings';
import { vueEmits } from 'ooxml-ui/bindings/vue';
import { act, createElement, createRef, StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot } from 'solid-js';
import { render } from 'solid-js/web';
import { compile } from 'svelte/compiler';
import { createApp, h, ref } from 'vue';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { EDITOR_CONTRACT, EDITOR_EVENT_NAMES, EDITOR_HANDLE_KEYS, EDITOR_PROP_KEYS } from './index';
import type { EditorHandle } from './index';
import { WordEditor as ReactEditor } from './react';
import { WordEditor as SolidEditor } from './solid';
import { WordEditor as VueEditor } from './vue';

const reactAct = globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean };
beforeEach(() => (reactAct.IS_REACT_ACT_ENVIRONMENT = true));
afterEach(() => {
	delete reactAct.IS_REACT_ACT_ENVIRONMENT;
	document.body.replaceChildren();
});
const sorted = (values: Iterable<string>) => [...values].sort();
const missing = (target: object) => EDITOR_HANDLE_KEYS.filter((key) => !(key in target));
const newHost = () => document.body.appendChild(document.createElement('div'));
const spies = () =>
	Object.fromEntries(optionNames(EDITOR_CONTRACT).map((option) => [option, vi.fn()]));
/** Dispatches every contract event on the editor; returns the options whose handler ran. */
function fireAll(handlers: Record<string, ReturnType<typeof vi.fn>>): string[] {
	const editor = document.querySelector('docx-editor')!;
	for (const name of EDITOR_EVENT_NAMES)
		editor.dispatchEvent(
			new CustomEvent(name, { detail: { readOnly: false, hostReadOnly: false } }),
		);
	return Object.entries(handlers)
		.filter(([, handler]) => handler.mock.calls.length > 0)
		.map(([option]) => option);
}

describe('EDITOR_CONTRACT', () => {
	it('lists the shared keys, events and handle members exactly once', () => {
		expect(sorted(propKeys(EDITOR_CONTRACT))).toEqual(sorted(EDITOR_PROP_KEYS));
		expect(sorted(Object.keys(EDITOR_CONTRACT.events))).toEqual(sorted(EDITOR_EVENT_NAMES));
		expect(handleKeys(EDITOR_CONTRACT.handle)).toEqual([...EDITOR_HANDLE_KEYS]);
		expect(sorted(EDITOR_HANDLE_KEYS)).toEqual(
			sorted(['dirty', 'download', 'element', 'load', 'markClean', 'save']),
		);
	});
});

describe('every framework adapter matches the contract', () => {
	it('React: every event reaches its handler prop, and the ref is the handle', async () => {
		const root = createRoot(newHost());
		const handle = createRef<EditorHandle>();
		const handlers = spies();
		await act(async () =>
			root.render(
				createElement(StrictMode, null, createElement(ReactEditor, { ref: handle, ...handlers })),
			),
		);
		expect(missing(handle.current!)).toEqual([]);
		expect(handle.current!.dirty).toBe(false);
		expect(sorted(fireAll(handlers))).toEqual(sorted(optionNames(EDITOR_CONTRACT)));
		act(() => root.unmount());
	});
	it('Vue: declares every prop, emits every event and v-model, exposes the handle', () => {
		expect(sorted(Object.keys(VueEditor.props))).toEqual(sorted(EDITOR_PROP_KEYS));
		expect(VueEditor.emits).toEqual(vueEmits(EDITOR_CONTRACT));
		expect(VueEditor.emits).toContain('update:readOnly');
		const editor = ref<EditorHandle>();
		const app = createApp({ render: () => h(VueEditor, { ref: editor }) });
		app.mount(newHost());
		expect(missing(editor.value!)).toEqual([]);
		expect(editor.value!.dirty).toBe(false);
		app.unmount();
	});
	it('Solid: every event reaches its handler prop, and editorRef gets the handle', () => {
		let handle: EditorHandle | undefined;
		const handlers = spies();
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(() => SolidEditor({ ...handlers, editorRef: (value) => (handle = value) }), newHost());
		});
		expect(missing(handle!)).toEqual([]);
		expect(handle!.dirty).toBe(false);
		expect(sorted(fireAll(handlers))).toEqual(sorted(optionNames(EDITOR_CONTRACT)));
		dispose();
	});
	it('Angular: one input per prop, one output per event, the handle on the class', async () => {
		const { WordEditorComponent } = await import('./angular');
		const mirror = reflectComponentType(WordEditorComponent)!;
		expect(sorted(mirror.inputs.map((input) => input.propName))).toEqual(sorted(EDITOR_PROP_KEYS));
		expect(sorted(mirror.outputs.map((output) => output.propName))).toEqual(
			sorted(optionNames(EDITOR_CONTRACT).map(angularOutputName)),
		);
		expect(missing(WordEditorComponent.prototype)).toEqual([]);
	});
	it('Svelte: one prop per key and handler, and the handle exports', () => {
		// The component is compiled by the package build, not here; check its source instead.
		const source = readFileSync(
			resolve(process.cwd(), 'packages/bindings/src/WordEditor.svelte'),
			'utf8',
		);
		const declared = /let \{([^}]*)\}: Props = \$props\(\)/u.exec(source)?.[1] ?? '';
		const names = declared.split(',').map((entry) => entry.split('=')[0]!.trim());
		expect(sorted(names)).toEqual(
			sorted([...EDITOR_PROP_KEYS, ...optionNames(EDITOR_CONTRACT).map(svelteHandlerName)]),
		);
		const { js } = compile(source, { filename: 'WordEditor.svelte', generate: 'client' });
		const exported = /var \$\$exports = \{([\s\S]*?)\};/u.exec(js.code)?.[1] ?? '';
		for (const key of EDITOR_HANDLE_KEYS)
			expect(exported, key).toMatch(new RegExp(String.raw`\b${key}\b`));
	});
});
