// @vitest-environment jsdom
import '@angular/compiler';
import { createComponent, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { XlsxEditorElement } from 'xlsx-web-component';
import { SpreadsheetEditorComponent } from './angular';
import { emit, fakeWorkbook } from './test-support';

afterEach(() => document.body.replaceChildren());

describe('Angular editor adapter', () => {
	it('mounts into the host, forwards inputs, re-emits outputs, and destroys', async () => {
		const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
		const hostElement = document.createElement('spreadsheet-editor');
		document.body.append(hostElement);
		const ref = createComponent(SpreadsheetEditorComponent, {
			environmentInjector: app.injector,
			hostElement,
		});
		ref.setInput('locale', 'zh-Hans');
		ref.setInput('authorName', 'Grace');
		const onChange = vi.fn();
		const onDirty = vi.fn();
		ref.instance.workbookChange.subscribe(onChange);
		ref.instance.dirtyChange.subscribe(onDirty);
		app.attachView(ref.hostView);
		ref.changeDetectorRef.detectChanges();
		const editor = hostElement.querySelector('xlsx-editor') as XlsxEditorElement;
		expect(editor).not.toBeNull();
		expect(editor.locale).toBe('zh-CN');
		expect(editor.authorName).toBe('Grace');
		expect(ref.instance.element).toBe(editor);
		ref.setInput('readOnly', true);
		ref.changeDetectorRef.detectChanges();
		expect(editor.readOnly).toBe(true);
		const workbook = fakeWorkbook();
		emit(editor, 'workbook-change', { workbook });
		emit(editor, 'dirty-change', { dirty: true });
		expect(onChange).toHaveBeenCalledWith(workbook);
		expect(onDirty).toHaveBeenCalledWith(true);
		ref.destroy();
		expect(hostElement.querySelector('xlsx-editor')).toBeNull();
		app.destroy();
	});
});
