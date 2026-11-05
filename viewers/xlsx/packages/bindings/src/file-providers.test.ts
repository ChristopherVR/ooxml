// @vitest-environment jsdom
import '@angular/compiler';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot, createSignal } from 'solid-js';
import { render } from 'solid-js/web';
import { flushSync, mount, unmount } from 'svelte';
import { createApp, h, nextTick, shallowRef } from 'vue';
import { afterEach, describe, expect, it } from 'vitest';
import type { XlsxEditorElement } from 'xlsx-web-component';
import type { CloudFileProvider } from 'ooxml-core/embedding';
import { SpreadsheetEditorComponent } from './angular';
import { EDITOR_PROP_KEYS, mountEditor } from './index';
import { SpreadsheetEditor as ReactEditor } from './react';
import { SpreadsheetEditor as SolidEditor } from './solid';
import { SpreadsheetEditor as VueEditor } from './vue';
import XlsxEditor from './XlsxEditor.svelte';
import { reactiveProps } from './test-svelte-props.svelte';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

type ProviderList = readonly CloudFileProvider[];
const providers = (): ProviderList => [{ id: 'cloud', label: 'Cloud' } as CloudFileProvider];
function host(): HTMLElement {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
}
const editorIn = (root: ParentNode) => root.querySelector('xlsx-editor') as XlsxEditorElement;

describe('cloud file providers through every binding', () => {
	it('is part of the shared contract', () => {
		expect(EDITOR_PROP_KEYS).toContain('fileProviders');
	});

	it('Vanilla: forwards the fileProviders and clears it when left out', () => {
		const first = providers();
		const binding = mountEditor(host(), { fileProviders: first });
		expect(binding.element.fileProviders).toBe(first);
		binding.update({});
		expect(binding.element.fileProviders).toEqual([]);
		binding.destroy();
	});

	it('React', () => {
		const container = host();
		const root = createRoot(container);
		const first = providers();
		act(() => root.render(createElement(ReactEditor, { fileProviders: first })));
		expect(editorIn(container).fileProviders).toBe(first);
		act(() => root.render(createElement(ReactEditor, { fileProviders: null })));
		expect(editorIn(container).fileProviders).toEqual([]);
		act(() => root.unmount());
	});

	it('Vue', async () => {
		const container = host();
		const first = providers();
		const current = shallowRef<ProviderList | null>(first);
		const app = createApp({ render: () => h(VueEditor, { fileProviders: current.value }) });
		app.mount(container);
		expect(editorIn(container).fileProviders).toBe(first);
		current.value = null;
		await nextTick();
		expect(editorIn(container).fileProviders).toEqual([]);
		app.unmount();
	});

	it('Solid', () => {
		const container = host();
		const first = providers();
		const [current, setCurrent] = createSignal<ProviderList | undefined>(first);
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
		expect(editorIn(container).fileProviders).toBe(first);
		setCurrent(undefined);
		expect(editorIn(container).fileProviders).toEqual([]);
		dispose();
	});

	it('Angular', async () => {
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const hostElement = document.createElement('spreadsheet-editor');
		document.body.append(hostElement);
		const component = createComponent(SpreadsheetEditorComponent, {
			environmentInjector: app.injector,
			hostElement,
		});
		const first = providers();
		component.setInput('fileProviders', first);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).fileProviders).toBe(first);
		component.setInput('fileProviders', null);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).fileProviders).toEqual([]);
		component.destroy();
		app.destroy();
	});

	it('Svelte', () => {
		const target = host();
		const first = providers();
		const props = reactiveProps<{ fileProviders: ProviderList | undefined }>({
			fileProviders: first,
		});
		const component = mount(XlsxEditor, { target, props });
		flushSync();
		// The test's `$state` props proxy the object, so compare what the editor will call.
		expect(editorIn(target).fileProviders[0]?.id).toBe(first[0]!.id);
		props.fileProviders = undefined;
		flushSync();
		expect(editorIn(target).fileProviders).toEqual([]);
		unmount(component);
	});
});
