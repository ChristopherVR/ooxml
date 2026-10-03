// @vitest-environment jsdom
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { XlsxEditorElement } from 'xlsx-web-component';
import XlsxEditor from './XlsxEditor.svelte';
import { reactiveProps } from './test-svelte-props.svelte';
import { emit, fakeWorkbook } from './test-support';

afterEach(() => document.body.replaceChildren());

describe('Svelte editor adapter', () => {
	it('mounts once, forwards props, calls lower-case callbacks, and unmounts', () => {
		const target = document.createElement('div');
		document.body.append(target);
		const onworkbookchange = vi.fn();
		const ondirtychange = vi.fn();
		const onready = vi.fn();
		const component = mount(XlsxEditor, {
			target,
			props: { locale: 'fr', readOnly: true, onworkbookchange, ondirtychange, onready },
		});
		flushSync();
		const editor = target.querySelector('xlsx-editor') as XlsxEditorElement;
		expect(editor.locale).toBe('fr');
		expect(editor.readOnly).toBe(true);
		expect(onready).toHaveBeenCalledWith(editor);
		expect(component.getElement()).toBe(editor);
		expect(component.isDirty()).toBe(false);
		const workbook = fakeWorkbook();
		emit(editor, 'workbook-change', { workbook });
		emit(editor, 'dirty-change', { dirty: true });
		expect(onworkbookchange).toHaveBeenCalledWith(workbook);
		expect(ondirtychange).toHaveBeenCalledWith(true);
		unmount(component);
		expect(target.querySelector('xlsx-editor')).toBeNull();
	});

	it('keeps in-editor changes when other props change and calls the new callbacks', () => {
		const target = document.createElement('div');
		document.body.append(target);
		const onreadonlychange = vi.fn();
		const onribboncustomize = vi.fn();
		const props = reactiveProps({ authorName: 'Ada', onreadonlychange, onribboncustomize });
		const component = mount(XlsxEditor, { target, props });
		flushSync();
		const editor = target.querySelector('xlsx-editor') as XlsxEditorElement;
		editor.readOnly = true;
		expect(onreadonlychange).toHaveBeenCalledWith(true);
		emit(editor, 'ribbon-customize', { hiddenActions: ['italic'] });
		expect(onribboncustomize).toHaveBeenCalledWith(['italic']);
		props.authorName = 'Grace';
		flushSync();
		expect(editor.authorName).toBe('Grace');
		expect(editor.readOnly).toBe(true);
		unmount(component);
	});
});
