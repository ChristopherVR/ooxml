// @vitest-environment jsdom
import '@angular/compiler';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot, createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { createApp, h, nextTick, shallowRef } from 'vue';
import { afterEach, describe, expect, it } from 'vitest';
import type { DocxEditorElement } from 'docx-web-component';
import type { DocumentStorage } from 'ooxml-core/embedding';
import { WordEditorComponent } from './angular';
import { EDITOR_PROP_KEYS, mountEditor } from './index';
import { WordEditor as ReactEditor } from './react';
import { WordEditor as SolidEditor } from './solid';
import { WordEditor as VueEditor } from './vue';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

const storage = (): DocumentStorage => ({ save: async () => undefined });
function host(): HTMLElement {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
}
const editorIn = (root: ParentNode) => root.querySelector('docx-editor') as DocxEditorElement;

describe('host-owned storage through every binding', () => {
	it('is part of the shared contract', () => {
		expect(EDITOR_PROP_KEYS).toContain('storage');
	});

	it('Vanilla: forwards the storage and clears it when left out', () => {
		const first = storage();
		const binding = mountEditor(host(), { storage: first });
		expect(binding.element.storage).toBe(first);
		binding.update({});
		expect(binding.element.storage).toBeUndefined();
		binding.destroy();
	});

	it('React', () => {
		const container = host();
		const root = createRoot(container);
		const first = storage();
		act(() => root.render(createElement(ReactEditor, { storage: first })));
		expect(editorIn(container).storage).toBe(first);
		act(() => root.render(createElement(ReactEditor, { storage: null })));
		expect(editorIn(container).storage).toBeUndefined();
		act(() => root.unmount());
	});

	it('Vue', async () => {
		const container = host();
		const first = storage();
		const current = shallowRef<DocumentStorage | null>(first);
		const app = createApp({ render: () => h(VueEditor, { storage: current.value }) });
		app.mount(container);
		expect(editorIn(container).storage).toBe(first);
		current.value = null;
		await nextTick();
		expect(editorIn(container).storage).toBeUndefined();
		app.unmount();
	});

	it('Solid', () => {
		const container = host();
		const first = storage();
		const [current, setCurrent] = createSignal<DocumentStorage | undefined>(first);
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(
				() =>
					SolidEditor({
						get storage() {
							return current();
						},
					}),
				container,
			);
		});
		expect(editorIn(container).storage).toBe(first);
		setCurrent(undefined);
		expect(editorIn(container).storage).toBeUndefined();
		dispose();
	});

	it('Angular', async () => {
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const hostElement = document.createElement('word-editor');
		document.body.append(hostElement);
		const component = createComponent(WordEditorComponent, {
			environmentInjector: app.injector,
			hostElement,
		});
		const first = storage();
		component.setInput('storage', first);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).storage).toBe(first);
		component.setInput('storage', null);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).storage).toBeUndefined();
		component.destroy();
		app.destroy();
	});

	it('Svelte', () => {
		// The component is compiled by the package build, not here; check what it forwards instead.
		const source = readFileSync(resolve(__dirname, 'WordEditor.svelte'), 'utf8');
		const attach = /use:attach=\{\{([^}]*)\}\}/.exec(source)?.[1] ?? '';
		expect(attach).toMatch(/\bstorage\b/);
		expect(source).toMatch(/let \{[^}]*\bstorage\b[^}]*\}: Props = \$props\(\)/);
	});
});
