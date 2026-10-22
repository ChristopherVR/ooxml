// @vitest-environment jsdom
import '@angular/compiler';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot } from 'solid-js';
import { render } from 'solid-js/web';
import { createApp, h, nextTick, shallowRef } from 'vue';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { DocxEditorElement } from 'docx-web-component';
import { WordEditorComponent } from './angular';
import { EDITOR_EVENT_NAMES, EDITOR_PROP_KEYS, mountEditor } from './index';
import { WordEditor as ReactEditor } from './react';
import { WordEditor as SolidEditor } from './solid';
import { WordEditor as VueEditor } from './vue';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

function host(): HTMLElement {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
}
const editorIn = (root: ParentNode) => root.querySelector('docx-editor') as DocxEditorElement;

/** Dispatches a cancelable file command the way the editor chrome does, and reports the result. */
function requestFileCommand(editor: DocxEditorElement, command: string) {
	const event = new CustomEvent('file-command', {
		detail: { command },
		bubbles: true,
		composed: true,
		cancelable: true,
	});
	const proceed = editor.dispatchEvent(event);
	return { event, proceed };
}

describe('shared adapter: file commands', () => {
	it('hands the cancelable file-command event to onFileCommand so a host can take over', () => {
		const onFileCommand = vi.fn((event: CustomEvent<{ command: string }>) => {
			if (event.detail.command === 'save') event.preventDefault();
		});
		const binding = mountEditor(host(), { onFileCommand });
		const save = requestFileCommand(binding.element, 'save');
		expect(onFileCommand.mock.calls[0]?.[0]).toBe(save.event);
		expect(save.event.defaultPrevented).toBe(true);
		expect(save.proceed).toBe(false);
		const open = requestFileCommand(binding.element, 'open');
		expect(open.event.defaultPrevented).toBe(false);
		expect(open.proceed).toBe(true);
		binding.destroy();
	});

	it('leaves the editor default in place when no host handler exists', () => {
		const binding = mountEditor(host(), {});
		expect(requestFileCommand(binding.element, 'save').proceed).toBe(true);
		binding.destroy();
	});
});

const DETAIL = { readOnly: true, locked: false, hostReadOnly: true };

describe('shared adapter: forwarded element events', () => {
	it('forwards warnings, read-only toggles, ribbon actions and ribbon customisation', () => {
		const onDocumentWarning = vi.fn();
		const onReadOnlyChange = vi.fn();
		const onRibbonAction = vi.fn();
		const onRibbonCustomize = vi.fn();
		const binding = mountEditor(host(), {
			onDocumentWarning,
			onReadOnlyChange,
			onRibbonAction,
			onRibbonCustomize,
		});
		const action = { type: 'bold' };
		binding.element.dispatchEvent(new CustomEvent('document-warning', { detail: 'Print issue' }));
		binding.element.dispatchEvent(new CustomEvent('readonly-change', { detail: DETAIL }));
		binding.element.dispatchEvent(new CustomEvent('ribbon-action', { detail: action }));
		binding.element.dispatchEvent(new CustomEvent('ribbon-customize', { detail: ['bold'] }));
		expect(onDocumentWarning).toHaveBeenCalledWith('Print issue');
		expect(onReadOnlyChange).toHaveBeenCalledWith(true, DETAIL);
		expect(onRibbonAction).toHaveBeenCalledWith(action);
		expect(onRibbonCustomize).toHaveBeenCalledWith(['bold']);
		binding.destroy();
		binding.element.dispatchEvent(
			new CustomEvent('readonly-change', { detail: { ...DETAIL, readOnly: false } }),
		);
		binding.element.dispatchEvent(new CustomEvent('file-command', { detail: { command: 'new' } }));
		expect(onReadOnlyChange).toHaveBeenCalledTimes(1);
	});

	it('passes the lock and the host switch so a parent can mirror only the host', () => {
		const onReadOnlyChange = vi.fn();
		const binding = mountEditor(host(), { onReadOnlyChange });
		const locked = { readOnly: true, locked: true, hostReadOnly: false };
		binding.element.dispatchEvent(new CustomEvent('readonly-change', { detail: locked }));
		expect(onReadOnlyChange).toHaveBeenCalledWith(true, locked);
		binding.destroy();
	});

	it('lists the new events and props in the shared contract', () => {
		for (const name of [
			'file-command',
			'document-warning',
			'readonly-change',
			'ribbon-action',
			'ribbon-customize',
		] as const)
			expect(EDITOR_EVENT_NAMES).toContain(name);
		for (const key of ['fileName', 'reviewAuthor', 'themeColors'] as const)
			expect(EDITOR_PROP_KEYS).toContain(key);
	});
});

describe('shared adapter: readOnly follows only its own changes', () => {
	it('keeps the Viewing toggle the user made when the parent re-renders with a stale value', () => {
		const host = document.createElement('div');
		document.body.append(host);
		const onReadOnlyChange = vi.fn();
		const binding = mountEditor(host, { readOnly: false, onReadOnlyChange });
		// The user switches to Viewing; the editor reports it and the parent does not sync.
		binding.element.readOnly = true;
		binding.element.dispatchEvent(new CustomEvent('readonly-change', { detail: DETAIL }));
		expect(onReadOnlyChange).toHaveBeenCalledWith(true, DETAIL);
		binding.update({ readOnly: false, onReadOnlyChange });
		expect(binding.element.readOnly).toBe(true);
		// A changed prop is still applied.
		binding.update({ readOnly: true, onReadOnlyChange });
		expect(binding.element.readOnly).toBe(true);
		binding.update({ readOnly: false, onReadOnlyChange });
		expect(binding.element.readOnly).toBe(false);
		binding.destroy();
	});
});

describe('shared adapter: fileName, reviewAuthor and themeColors', () => {
	it('forwards them and keeps a rename made inside the editor on unrelated updates', () => {
		const binding = mountEditor(host(), {
			fileName: 'Plan.docx',
			reviewAuthor: 'Ada',
			themeColors: { primary: '#123456' },
		});
		const element = binding.element;
		expect(element.fileName).toBe('Plan.docx');
		expect(element.reviewAuthor).toBe('Ada');
		expect(element.themeColors).toEqual({ primary: '#123456' });
		// File > Open renames the document inside the editor; the parent still passes the old value.
		element.fileName = 'Renamed.docx';
		binding.update({
			fileName: 'Plan.docx',
			reviewAuthor: 'Ada',
			themeColors: { primary: '#123456' },
		});
		expect(element.fileName).toBe('Renamed.docx');
		// A changed prop is forwarded again; an unset one leaves the element's own value alone.
		binding.update({ fileName: 'Other.docx' });
		expect(element.fileName).toBe('Other.docx');
		expect(element.reviewAuthor).toBe('Ada');
		binding.destroy();
	});
});

describe('framework bindings: file commands and fileName', () => {
	it('Vanilla: forwards the file-command event and the file name', () => {
		const onFileCommand = vi.fn((event: CustomEvent) => event.preventDefault());
		const binding = mountEditor(host(), { fileName: 'Plan.docx', onFileCommand });
		expect(binding.element.fileName).toBe('Plan.docx');
		const { proceed } = requestFileCommand(binding.element, 'save');
		expect(onFileCommand).toHaveBeenCalledTimes(1);
		expect(proceed).toBe(false);
		binding.destroy();
	});

	it('React', () => {
		const container = host();
		const root = createRoot(container);
		const onFileCommand = vi.fn((event: CustomEvent) => event.preventDefault());
		act(() => root.render(createElement(ReactEditor, { fileName: 'Plan.docx', onFileCommand })));
		const editor = editorIn(container);
		expect(editor.fileName).toBe('Plan.docx');
		const { proceed } = requestFileCommand(editor, 'open');
		expect(onFileCommand).toHaveBeenCalledTimes(1);
		expect(proceed).toBe(false);
		act(() => root.unmount());
	});

	it('Vue', async () => {
		const container = host();
		const onFileCommand = vi.fn((event: CustomEvent) => event.preventDefault());
		const current = shallowRef('Plan.docx');
		const app = createApp({
			render: () => h(VueEditor, { fileName: current.value, 'onFile-command': onFileCommand }),
		});
		app.mount(container);
		const editor = editorIn(container);
		expect(editor.fileName).toBe('Plan.docx');
		const { proceed } = requestFileCommand(editor, 'save');
		expect(onFileCommand).toHaveBeenCalledTimes(1);
		expect(proceed).toBe(false);
		current.value = 'Next.docx';
		await nextTick();
		expect(editor.fileName).toBe('Next.docx');
		app.unmount();
	});

	it('Solid', () => {
		const container = host();
		const onFileCommand = vi.fn((event: CustomEvent) => event.preventDefault());
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(() => SolidEditor({ fileName: 'Plan.docx', onFileCommand }), container);
		});
		const editor = editorIn(container);
		expect(editor.fileName).toBe('Plan.docx');
		const { proceed } = requestFileCommand(editor, 'save');
		expect(onFileCommand).toHaveBeenCalledTimes(1);
		expect(proceed).toBe(false);
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
		const seen: CustomEvent[] = [];
		component.instance.fileCommand.subscribe((event) => {
			seen.push(event);
			event.preventDefault();
		});
		component.setInput('fileName', 'Plan.docx');
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		const editor = editorIn(hostElement);
		expect(editor.fileName).toBe('Plan.docx');
		const { event, proceed } = requestFileCommand(editor, 'save');
		expect(seen).toEqual([event]);
		expect(proceed).toBe(false);
		component.destroy();
		app.destroy();
	});

	it('Svelte', () => {
		// The component is compiled by the package build, not here; check what it forwards instead.
		const source = readFileSync(
			resolve(process.cwd(), 'packages/bindings/src/WordEditor.svelte'),
			'utf8',
		);
		const attach = /use:attach=\{\{([^}]*)\}\}/.exec(source)?.[1] ?? '';
		for (const name of ['fileName', 'reviewAuthor', 'themeColors', 'onfilecommand'])
			expect(attach).toMatch(new RegExp(`\\b${name}\\b`));
		expect(source).toContain("'file-command': value.onfilecommand");
		expect(source).toContain("'readonly-change': value.onreadonlychange");
		expect(source).toContain("'ribbon-customize': value.onribboncustomize");
	});
});
