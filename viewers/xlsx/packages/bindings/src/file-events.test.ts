// @vitest-environment jsdom
import '@angular/compiler';
import { flushSync, mount, unmount } from 'svelte';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { createRoot as createSolidRoot } from 'solid-js';
import { render } from 'solid-js/web';
import { createApp, h, nextTick, shallowRef } from 'vue';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { XlsxEditorElement } from 'xlsx-web-component';
import { SpreadsheetEditorComponent } from './angular';
import { EDITOR_EVENT_NAMES, mountEditor } from './index';
import { SpreadsheetEditor as ReactEditor } from './react';
import { SpreadsheetEditor as SolidEditor } from './solid';
import { SpreadsheetEditor as VueEditor } from './vue';
import XlsxEditor from './XlsxEditor.svelte';
import { emit } from './test-support';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
afterEach(() => document.body.replaceChildren());

function host(): HTMLElement {
	const element = document.createElement('div');
	document.body.append(element);
	return element;
}
const editorIn = (root: ParentNode) => root.querySelector('xlsx-editor') as XlsxEditorElement;

/** Dispatches a cancelable File command the way the editor chrome does, and reports the result. */
function requestFileCommand(editor: XlsxEditorElement, command: string) {
	const event = new CustomEvent('file-command', {
		detail: { command },
		bubbles: true,
		composed: true,
		cancelable: true,
	});
	return { event, proceed: editor.dispatchEvent(event) };
}

describe('shared adapter: file commands and forwarded events', () => {
	it('hands the cancelable file-command event to onFileCommand so a host can take over', () => {
		const onFileCommand = vi.fn((event: CustomEvent<{ command: string }>) => {
			if (event.detail.command === 'save') event.preventDefault();
		});
		const binding = mountEditor(host(), { onFileCommand });
		const save = requestFileCommand(binding.element, 'save');
		expect(onFileCommand.mock.calls[0]?.[0]).toBe(save.event);
		expect(save.proceed).toBe(false);
		const open = requestFileCommand(binding.element, 'open');
		expect(open.proceed).toBe(true);
		binding.destroy();
	});

	it('forwards workbook warnings, ribbon actions and sheet changes', () => {
		const onWorkbookWarning = vi.fn();
		const onRibbonAction = vi.fn();
		const onSheetChange = vi.fn();
		const binding = mountEditor(host(), { onWorkbookWarning, onRibbonAction, onSheetChange });
		emit(binding.element, 'workbook-warning', { warnings: ['Macros removed'] });
		emit(binding.element, 'ribbon-action', { id: 'bold' });
		emit(binding.element, 'sheet-change', { index: 1, name: 'Data' });
		expect(onWorkbookWarning).toHaveBeenCalledWith(['Macros removed']);
		expect(onRibbonAction).toHaveBeenCalledWith({ id: 'bold' });
		expect(onSheetChange).toHaveBeenCalledWith({ index: 1, name: 'Data' });
		binding.destroy();
		emit(binding.element, 'sheet-change', { index: 0, name: 'Sheet1' });
		expect(onSheetChange).toHaveBeenCalledTimes(1);
	});

	it('lists the new events in the shared contract', () => {
		for (const name of [
			'file-command',
			'workbook-warning',
			'ribbon-action',
			'sheet-change',
		] as const)
			expect(EDITOR_EVENT_NAMES).toContain(name);
	});
});

describe('framework bindings: file commands', () => {
	it('Vanilla', () => {
		const onFileCommand = vi.fn((event: CustomEvent) => event.preventDefault());
		const binding = mountEditor(host(), { onFileCommand });
		expect(requestFileCommand(binding.element, 'save').proceed).toBe(false);
		expect(onFileCommand).toHaveBeenCalledTimes(1);
		binding.destroy();
	});

	it('React', () => {
		const container = host();
		const root = createRoot(container);
		const onFileCommand = vi.fn((event: CustomEvent) => event.preventDefault());
		act(() => root.render(createElement(ReactEditor, { onFileCommand })));
		const editor = editorIn(container);
		expect(requestFileCommand(editor, 'save').proceed).toBe(false);
		expect(onFileCommand).toHaveBeenCalledTimes(1);
		act(() => root.unmount());
	});

	it('Vue', async () => {
		const container = host();
		const onFileCommand = vi.fn((event: CustomEvent) => event.preventDefault());
		const current = shallowRef('Book1.xlsx');
		const app = createApp({
			render: () => h(VueEditor, { fileName: current.value, 'onFile-command': onFileCommand }),
		});
		app.mount(container);
		const editor = editorIn(container);
		expect(editor.fileName).toBe('Book1.xlsx');
		expect(requestFileCommand(editor, 'save').proceed).toBe(false);
		expect(onFileCommand).toHaveBeenCalledTimes(1);
		current.value = 'Next.xlsx';
		await nextTick();
		expect(editor.fileName).toBe('Next.xlsx');
		app.unmount();
	});

	it('Solid', () => {
		const container = host();
		const onFileCommand = vi.fn((event: CustomEvent) => event.preventDefault());
		let dispose!: () => void;
		createSolidRoot((stop) => {
			dispose = stop;
			render(() => SolidEditor({ onFileCommand }), container);
		});
		expect(requestFileCommand(editorIn(container), 'save').proceed).toBe(false);
		expect(onFileCommand).toHaveBeenCalledTimes(1);
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
		const seen: CustomEvent[] = [];
		component.instance.fileCommand.subscribe((event) => {
			seen.push(event);
			event.preventDefault();
		});
		app.attachView(component.hostView);
		component.changeDetectorRef.detectChanges();
		const { event, proceed } = requestFileCommand(editorIn(hostElement), 'save');
		expect(seen).toEqual([event]);
		expect(proceed).toBe(false);
		component.destroy();
		app.destroy();
	});

	it('Svelte', () => {
		const target = document.createElement('div');
		document.body.append(target);
		const onfilecommand = vi.fn((event: CustomEvent) => event.preventDefault());
		const onworkbookwarning = vi.fn();
		const onsheetchange = vi.fn();
		const component = mount(XlsxEditor, {
			target,
			props: { onfilecommand, onworkbookwarning, onsheetchange },
		});
		flushSync();
		const editor = target.querySelector('xlsx-editor') as XlsxEditorElement;
		expect(requestFileCommand(editor, 'save').proceed).toBe(false);
		expect(onfilecommand).toHaveBeenCalledTimes(1);
		emit(editor, 'workbook-warning', { warnings: ['Note'] });
		emit(editor, 'sheet-change', { index: 0, name: 'Sheet1' });
		expect(onworkbookwarning).toHaveBeenCalledWith(['Note']);
		expect(onsheetchange).toHaveBeenCalledWith({ index: 0, name: 'Sheet1' });
		unmount(component);
	});
});
