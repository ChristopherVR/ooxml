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
import type { CloudFileProvider } from 'ooxml-core/embedding';
import { WordEditorComponent } from './angular';
import { EDITOR_PROP_KEYS, mountEditor } from './index';
import { WordEditor as ReactEditor } from './react';
import { WordEditor as SolidEditor } from './solid';
import { WordEditor as VueEditor } from './vue';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

const fake = (): CloudFileProvider => ({
	id: 'fake',
	label: 'Fake',
	account: () => null,
	signIn: async () => ({ name: 'Ann' }),
	list: async () => ({ entries: [] }),
	open: async () => {
		throw new Error('unused');
	},
});
const providers = () => [fake()];
function host(): HTMLElement {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
}
const editorIn = (root: ParentNode) => root.querySelector('docx-editor') as DocxEditorElement;

describe('cloud file providers through every binding', () => {
	it('is part of the shared contract', () => {
		expect(EDITOR_PROP_KEYS).toContain('fileProviders');
	});

	it('Vanilla: forwards the fileProviders and clears it when left out', () => {
		const first = providers();
		const binding = mountEditor(host(), { fileProviders: first });
		expect(binding.element.fileProviders).toEqual(first);
		binding.update({});
		expect(binding.element.fileProviders).toEqual([]);
		binding.destroy();
	});

	it('React', () => {
		const container = host();
		const root = createRoot(container);
		const first = providers();
		act(() => root.render(createElement(ReactEditor, { fileProviders: first })));
		expect(editorIn(container).fileProviders).toEqual(first);
		act(() => root.render(createElement(ReactEditor, { fileProviders: null })));
		expect(editorIn(container).fileProviders).toEqual([]);
		act(() => root.unmount());
	});

	it('Vue', async () => {
		const container = host();
		const first = providers();
		const current = shallowRef<readonly CloudFileProvider[] | null>(first);
		const app = createApp({ render: () => h(VueEditor, { fileProviders: current.value }) });
		app.mount(container);
		expect(editorIn(container).fileProviders).toEqual(first);
		current.value = null;
		await nextTick();
		expect(editorIn(container).fileProviders).toEqual([]);
		app.unmount();
	});

	it('Solid', () => {
		const container = host();
		const first = providers();
		const [current, setCurrent] = createSignal<readonly CloudFileProvider[] | undefined>(first);
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(
				() =>
					SolidEditor({
						get fileProviders() {
							return current();
						},
					}),
				container,
			);
		});
		expect(editorIn(container).fileProviders).toEqual(first);
		setCurrent(undefined);
		expect(editorIn(container).fileProviders).toEqual([]);
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
		component.setInput('fileProviders', first);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).fileProviders).toEqual(first);
		component.setInput('fileProviders', null);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).fileProviders).toEqual([]);
		component.destroy();
		app.destroy();
	});

	it('Svelte', () => {
		// The component is compiled by the package build, not here; check what it forwards instead.
		const source = readFileSync(resolve(__dirname, 'WordEditor.svelte'), 'utf8');
		const attach = /use:attach=\{\{([^}]*)\}\}/.exec(source)?.[1] ?? '';
		expect(attach).toMatch(/\bfileProviders\b/);
		expect(source).toMatch(/let \{[^}]*\bfileProviders\b[^}]*\}: Props = \$props\(\)/);
	});
});
