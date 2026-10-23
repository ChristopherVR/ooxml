// @vitest-environment jsdom
// Every framework forwards `collaboration`, re-emits `collaboration-change` and leaves the room
// when the editor unmounts. Joining itself is the element's job (ooxml-ui tests).
import '@angular/compiler';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot } from 'solid-js';
import { render } from 'solid-js/web';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { compile } from 'svelte/compiler';
import { createApp, h } from 'vue';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { DocxCollaborationState, DocxEditorElement } from 'docx-web-component';
import { WordEditorComponent } from './angular';
import { mountEditor } from './index';
import { WordEditor as ReactEditor } from './react';
import { WordEditor as SolidEditor } from './solid';
import { WordEditor as VueEditor } from './vue';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

const options = { roomId: 'report-2026', user: { name: 'Ada', color: '#2563eb' } };
const state: DocxCollaborationState = {
	active: true,
	roomId: 'report-2026',
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
function check(container: Element, onChange: ReturnType<typeof vi.fn>) {
	const editor = container.querySelector('docx-editor') as DocxEditorElement;
	expect(editor.collaboration).toBe(options);
	editor.dispatchEvent(new CustomEvent('collaboration-change', { detail: state }));
	expect(onChange).toHaveBeenCalledWith(state);
	return editor;
}

describe('collaboration in every binding', () => {
	it('vanilla (mountEditor)', () => {
		const onCollaborationChange = vi.fn();
		const node = host();
		const binding = mountEditor(node, { collaboration: options, onCollaborationChange });
		const editor = check(node, onCollaborationChange);
		const stop = vi.spyOn(editor, 'stopCollaboration');
		// The same object again is not a change; null leaves the room.
		binding.update({ collaboration: options, onCollaborationChange });
		expect(editor.collaboration).toBe(options);
		binding.update({ collaboration: null, onCollaborationChange });
		expect(editor.collaboration).toBeNull();
		binding.update({ collaboration: options, onCollaborationChange });
		binding.destroy();
		expect(stop).toHaveBeenCalled();
		expect(editor.collaboration).toBeNull();
	});

	it('does not stop an editor that never joined through the property', () => {
		const node = host();
		const binding = mountEditor(node, {});
		const stop = vi.spyOn(binding.element, 'stopCollaboration');
		binding.destroy();
		expect(stop).not.toHaveBeenCalled();
	});

	it('React', () => {
		const node = host();
		const root = createRoot(node);
		const onCollaborationChange = vi.fn();
		act(() =>
			root.render(createElement(ReactEditor, { collaboration: options, onCollaborationChange })),
		);
		const editor = check(node, onCollaborationChange);
		act(() => root.unmount());
		expect(editor.collaboration).toBeNull();
	});

	it('Vue', () => {
		const node = host();
		const onCollaborationChange = vi.fn();
		const app = createApp({
			render: () => h(VueEditor, { collaboration: options, onCollaborationChange }),
		});
		app.mount(node);
		const editor = check(node, onCollaborationChange);
		app.unmount();
		expect(editor.collaboration).toBeNull();
	});

	it('Angular', async () => {
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const hostElement = document.createElement('word-editor');
		document.body.append(hostElement);
		const component = createComponent(WordEditorComponent, {
			environmentInjector: app.injector,
			hostElement,
		});
		component.setInput('collaboration', options);
		const onChange = vi.fn();
		component.instance.collaborationChange.subscribe(onChange);
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		const editor = check(hostElement, onChange);
		component.destroy();
		expect(editor.collaboration).toBeNull();
		app.destroy();
	});

	it('Svelte', () => {
		// The vitest setup has no Svelte plugin, so check the compiled component's wiring instead:
		// it must accept `collaboration` and `oncollaborationchange` and hand both to the binding.
		const source = readFileSync(resolve(__dirname, 'WordEditor.svelte'), 'utf8');
		expect(() => compile(source, { filename: 'WordEditor.svelte' })).not.toThrow();
		expect(source).toMatch(/collaboration,[^}]*oncollaborationchange/);
		// Handlers reach the binding through the contract (`contract-parity.test.ts`).
		expect(source).toContain('...svelteEditorOptions(value)');
		expect(source).toMatch(
			/ribbonAddIns, collaboration, (customization, )?(storage, )?ondocumentchange/,
		);
		expect(source).toContain('onribbonaddin, oncollaborationchange }}');
	});

	it('Solid', () => {
		const node = host();
		let dispose!: () => void;
		const onCollaborationChange = vi.fn();
		createSolidRoot((stop) => {
			dispose = stop;
			render(() => SolidEditor({ collaboration: options, onCollaborationChange }), node);
		});
		const editor = check(node, onCollaborationChange);
		dispose();
		expect(editor.collaboration).toBeNull();
	});
});
