// @vitest-environment jsdom
// `googleFonts` is a shared prop: every binding forwards it to `<xlsx-editor>`, and `false`
// (downloads off) is a change from unset (the reader decides), not the same value.
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
import type { GoogleFontsOption } from 'ooxml-core/text';
import { SpreadsheetEditorComponent } from './angular';
import { EDITOR_PROP_KEYS, mountEditor } from './index';
import { SpreadsheetEditor as ReactEditor } from './react';
import { SpreadsheetEditor as SolidEditor } from './solid';
import { SpreadsheetEditor as VueEditor } from './vue';
import XlsxEditor from './XlsxEditor.svelte';
import { reactiveProps } from './test-svelte-props.svelte';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

const MIRROR = 'https://fonts.example.test/css2';
type Option = GoogleFontsOption | null | undefined;
function host(): HTMLElement {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
}
const editorIn = (root: ParentNode) => root.querySelector('xlsx-editor') as XlsxEditorElement;

describe('googleFonts through every binding', () => {
	it('is part of the shared contract', () => {
		expect(EDITOR_PROP_KEYS).toContain('googleFonts');
	});

	it('Vanilla: forwards false, then a mirror, then clears it when left out', () => {
		const binding = mountEditor(host(), { googleFonts: false });
		expect(binding.element.googleFonts).toBe(false);
		binding.update({ googleFonts: { baseUrl: MIRROR } });
		expect(binding.element.googleFonts).toEqual({ baseUrl: MIRROR });
		binding.update({});
		expect(binding.element.googleFonts).toBeUndefined();
		binding.destroy();
	});

	it('Vanilla: a fresh literal for the same mirror is not written again', () => {
		const binding = mountEditor(host(), { googleFonts: { baseUrl: MIRROR } });
		const written = binding.element.googleFonts;
		binding.update({ googleFonts: { baseUrl: MIRROR } });
		expect(binding.element.googleFonts).toBe(written);
		binding.destroy();
	});

	it('React', () => {
		const container = host();
		const root = createRoot(container);
		act(() => root.render(createElement(ReactEditor, { googleFonts: false })));
		expect(editorIn(container).googleFonts).toBe(false);
		act(() => root.render(createElement(ReactEditor, { googleFonts: null })));
		expect(editorIn(container).googleFonts).toBeUndefined();
		act(() => root.unmount());
	});

	it('Vue: null leaves it unset rather than turning downloads off', async () => {
		const container = host();
		const current = shallowRef<GoogleFontsOption | null>(null);
		const app = createApp({ render: () => h(VueEditor, { googleFonts: current.value }) });
		app.mount(container);
		expect(editorIn(container).googleFonts).toBeUndefined();
		current.value = false;
		await nextTick();
		expect(editorIn(container).googleFonts).toBe(false);
		current.value = { baseUrl: MIRROR };
		await nextTick();
		expect(editorIn(container).googleFonts).toEqual({ baseUrl: MIRROR });
		app.unmount();
	});

	it('Solid', () => {
		const container = host();
		const [current, setCurrent] = createSignal<Option>(false);
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(
				() =>
					SolidEditor({
						get googleFonts() {
							return current();
						},
					}),
				container,
			);
		});
		expect(editorIn(container).googleFonts).toBe(false);
		setCurrent(undefined);
		expect(editorIn(container).googleFonts).toBeUndefined();
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
		component.setInput('googleFonts', false);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).googleFonts).toBe(false);
		component.setInput('googleFonts', null);
		component.changeDetectorRef.detectChanges();
		expect(editorIn(hostElement).googleFonts).toBeUndefined();
		component.destroy();
		app.destroy();
	});

	it('Svelte', () => {
		const target = host();
		const props = reactiveProps<{ googleFonts: Option }>({ googleFonts: false });
		const component = mount(XlsxEditor, { target, props });
		flushSync();
		expect(editorIn(target).googleFonts).toBe(false);
		props.googleFonts = undefined;
		flushSync();
		expect(editorIn(target).googleFonts).toBeUndefined();
		unmount(component);
	});
});
