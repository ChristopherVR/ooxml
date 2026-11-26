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
import { WordEditorComponent } from './angular';
import { EDITOR_PROP_KEYS, mountEditor } from './index';
import { WordEditor as ReactEditor } from './react';
import { WordEditor as SolidEditor } from './solid';
import { WordEditor as VueEditor } from './vue';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

const providers = () => ['video.example.com'];
function host(): HTMLElement {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
}
const editorIn = (root: ParentNode) => root.querySelector('docx-editor') as DocxEditorElement;

describe('online video hosts through every binding', () => {
	it('is part of the shared contract', () => {
		expect(EDITOR_PROP_KEYS).toContain('onlineVideoHosts');
	});

	it('Vanilla: forwards the onlineVideoHosts and clears it when left out', () => {
		const first = providers();
		const binding = mountEditor(host(), { onlineVideoHosts: first });
		expect(binding.element.onlineVideoHosts).toEqual(first);
		binding.update({});
		expect(binding.element.onlineVideoHosts).toEqual([]);
		binding.destroy();
	});

	it('React', () => {
		const container = host();
		const root = createRoot(container);
		const first = providers();
		act(() => root.render(createElement(ReactEditor, { onlineVideoHosts: first })));
		expect(editorIn(container).onlineVideoHosts).toEqual(first);
		act(() => root.render(createElement(ReactEditor, { onlineVideoHosts: null })));
		expect(editorIn(container).onlineVideoHosts).toEqual([]);
		act(() => root.unmount());
	});

	it('Vue', async () => {
		const container = host();
		const first = providers();
		const current = shallowRef<readonly string[] | null>(first);
		const app = createApp({ render: () => h(VueEditor, { onlineVideoHosts: current.value }) });
		app.mount(container);
		expect(editorIn(container).onlineVideoHosts).toEqual(first);
		current.value = null;
		await nextTick();
		expect(editorIn(container).onlineVideoHosts).toEqual([]);
		app.unmount();
	});

	it('Solid', () => {
		const container = host();
		const first = providers();
		const [current, setCurrent] = createSignal<readonly string[] | undefined>(first);
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(
				() =>
					SolidEditor({
						get onlineVideoHosts() {
							return current();
						},
					}),
				container,
			);
		});
		expect(editorIn(container).onlineVideoHosts).toEqual(first);
		setCurrent(undefined);
		expect(editorIn(container).onlineVideoHosts).toEqual([]);
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
		const first = providers();
		component.setInput('onlineVideoHosts', first);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).onlineVideoHosts).toEqual(first);
		component.setInput('onlineVideoHosts', null);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).onlineVideoHosts).toEqual([]);
		component.destroy();
		app.destroy();
	});

	it('Svelte', () => {
		// The component is compiled by the package build, not here; check what it forwards instead.
		const source = readFileSync(resolve(__dirname, 'WordEditor.svelte'), 'utf8');
		const attach = /use:attach=\{\{([^}]*)\}\}/.exec(source)?.[1] ?? '';
		expect(attach).toMatch(/\bonlineVideoHosts\b/);
		expect(source).toMatch(/let \{[^}]*\bonlineVideoHosts\b[^}]*\}: Props = \$props\(\)/);
	});
});
