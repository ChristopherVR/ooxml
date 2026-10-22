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
import type { ExcelCustomization, XlsxEditorElement } from 'xlsx-web-component';
import { SpreadsheetEditorComponent } from './angular';
import { EDITOR_PROP_KEYS, mountEditor } from './index';
import { SpreadsheetEditor as ReactEditor } from './react';
import { SpreadsheetEditor as SolidEditor } from './solid';
import { SpreadsheetEditor as VueEditor } from './vue';
import XlsxEditor from './XlsxEditor.svelte';
import { reactiveProps } from './test-svelte-props.svelte';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

/** One stable object, as a host would memoise it. */
const kiosk: ExcelCustomization = {
	capabilities: { edit: false, print: false },
	backstage: { extraPages: [{ id: 'about', label: 'About', render: 'acme-about' }] },
};

function host(): HTMLElement {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
}
const editorIn = (root: ParentNode) => root.querySelector('xlsx-editor') as XlsxEditorElement;
/** What the element shows: its customisation, read-only state and File tab entries. */
function shown(editor: XlsxEditorElement) {
	const nav = (
		editor.shadowRoot!.querySelector('office-ui-backstage') as HTMLElement & {
			items: { id: string }[];
		}
	).items.map((item) => item.id);
	return {
		customization: editor.customization,
		readOnly: editor.readOnly,
		print: nav.includes('print'),
		hostPage: nav.includes('host.about'),
	};
}
/** `kiosk` applied: read-only, no Print page, the host page in the nav. */
const KIOSK = { customization: kiosk, readOnly: true, print: false, hostPage: true };
const DEFAULT = { customization: null, readOnly: false, print: true, hostPage: false };

describe('customization through every binding', () => {
	it('is part of the shared contract', () => {
		expect(EDITOR_PROP_KEYS).toContain('customization');
	});

	it('Vanilla: forwards the object, and leaving it out restores the defaults', () => {
		const binding = mountEditor(host(), { customization: kiosk });
		expect(shown(binding.element)).toEqual(KIOSK);
		binding.update({ customization: kiosk });
		expect(shown(binding.element)).toEqual(KIOSK);
		binding.update({});
		expect(shown(binding.element)).toEqual(DEFAULT);
		binding.destroy();
	});

	it('React', () => {
		const container = host();
		const root = createRoot(container);
		const renderWith = (customization: ExcelCustomization | undefined) =>
			act(() => root.render(createElement(ReactEditor, { customization })));
		renderWith(kiosk);
		const editor = editorIn(container);
		expect(shown(editor)).toEqual(KIOSK);
		renderWith(undefined);
		expect(shown(editor)).toEqual(DEFAULT);
		act(() => root.unmount());
	});

	it('Vue', async () => {
		const container = host();
		const current = shallowRef<ExcelCustomization | undefined>(kiosk);
		const app = createApp({
			render: () => h(VueEditor, current.value ? { customization: current.value } : {}),
		});
		app.mount(container);
		const editor = editorIn(container);
		expect(shown(editor)).toEqual(KIOSK);
		current.value = undefined;
		await nextTick();
		expect(shown(editor)).toEqual(DEFAULT);
		app.unmount();
	});

	it('Solid', () => {
		const container = host();
		const [current, setCurrent] = createSignal<ExcelCustomization | undefined>(kiosk);
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(
				() =>
					SolidEditor({
						get customization() {
							return current();
						},
					}),
				container,
			);
		});
		const editor = editorIn(container);
		expect(shown(editor)).toEqual(KIOSK);
		setCurrent(undefined);
		expect(shown(editor)).toEqual(DEFAULT);
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
		component.setInput('customization', kiosk);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		const editor = editorIn(hostElement);
		expect(shown(editor)).toEqual(KIOSK);
		component.setInput('customization', null);
		component.changeDetectorRef.detectChanges();
		expect(shown(editor)).toEqual(DEFAULT);
		component.destroy();
		app.destroy();
	});

	it('Svelte', () => {
		const target = host();
		const props = reactiveProps<{ customization: ExcelCustomization | undefined }>({
			customization: kiosk,
		});
		const component = mount(XlsxEditor, { target, props });
		flushSync();
		const editor = editorIn(target);
		expect(shown(editor)).toEqual(KIOSK);
		props.customization = undefined;
		flushSync();
		expect(shown(editor)).toEqual(DEFAULT);
		unmount(component);
	});
});
