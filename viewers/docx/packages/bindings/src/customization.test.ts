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
import type { DocxEditorElement, WordCustomization } from 'docx-web-component';
import { WordEditorComponent } from './angular';
import { EDITOR_PROP_KEYS, mountEditor } from './index';
import { WordEditor as ReactEditor } from './react';
import { WordEditor as SolidEditor } from './solid';
import { WordEditor as VueEditor } from './vue';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

/** One stable object per test, as a host would memoise it. */
const kiosk: WordCustomization = {
	capabilities: { edit: false, print: false },
	backstage: { extraPages: [{ id: 'about', label: 'About', render: 'acme-about' }] },
};

function host(): HTMLElement {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
}
const editorIn = (root: ParentNode) => root.querySelector('docx-editor') as DocxEditorElement;
/** What the element shows for `kiosk`: read-only, no Print page, the host page in the nav. */
function expectKiosk(editor: DocxEditorElement): void {
	expect(editor.customization).toBe(kiosk);
	expect(editor.readOnly).toBe(true);
	const nav = (
		editor.shadowRoot!.querySelector('office-ui-backstage') as HTMLElement & {
			items: { id: string }[];
		}
	).items.map((item) => item.id);
	expect(nav).not.toContain('print');
	expect(nav).toContain('host.about');
}
function expectDefault(editor: DocxEditorElement): void {
	expect(editor.customization).toBeNull();
	expect(editor.readOnly).toBe(false);
}

describe('customization through every binding', () => {
	it('is part of the shared contract', () => {
		expect(EDITOR_PROP_KEYS).toContain('customization');
	});

	it('Vanilla: forwards the object, and leaving it out restores the defaults', () => {
		const binding = mountEditor(host(), { customization: kiosk });
		expectKiosk(binding.element);
		binding.update({ customization: kiosk });
		expectKiosk(binding.element);
		binding.update({});
		expectDefault(binding.element);
		binding.destroy();
	});

	it('React', () => {
		const container = host();
		const root = createRoot(container);
		const renderWith = (customization: WordCustomization | undefined) =>
			act(() => root.render(createElement(ReactEditor, { customization })));
		renderWith(kiosk);
		const editor = editorIn(container);
		expectKiosk(editor);
		renderWith(undefined);
		expectDefault(editor);
		act(() => root.unmount());
	});

	it('Vue', async () => {
		const container = host();
		const current = shallowRef<WordCustomization | undefined>(kiosk);
		const app = createApp({
			render: () => h(VueEditor, current.value ? { customization: current.value } : {}),
		});
		app.mount(container);
		const editor = editorIn(container);
		expectKiosk(editor);
		current.value = undefined;
		await nextTick();
		expectDefault(editor);
		app.unmount();
	});

	it('Solid', () => {
		const container = host();
		const [current, setCurrent] = createSignal<WordCustomization | undefined>(kiosk);
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
		expectKiosk(editor);
		setCurrent(undefined);
		expectDefault(editor);
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
		component.setInput('customization', kiosk);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		const editor = editorIn(hostElement);
		expectKiosk(editor);
		component.setInput('customization', null);
		component.changeDetectorRef.detectChanges();
		expectDefault(editor);
		component.destroy();
		app.destroy();
	});

	it('Svelte', () => {
		// The component is compiled by the package build, not here; check what it forwards instead.
		const source = readFileSync(
			resolve(process.cwd(), 'packages/bindings/src/WordEditor.svelte'),
			'utf8',
		);
		const props = /let \{([^}]*)\}: Props = \$props\(\)/.exec(source)?.[1] ?? '';
		const attach = /use:attach=\{\{([^}]*)\}\}/.exec(source)?.[1] ?? '';
		expect(props).toMatch(/\bcustomization\b/);
		expect(attach).toMatch(/\bcustomization\b/);
	});
});
