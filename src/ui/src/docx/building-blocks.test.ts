// @vitest-environment jsdom
import { afterEach, expect, it } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import { createDocument, loadDocx } from 'ooxml-core/docx';
import { registerDocxEditor } from './index';
import type { DocxEditorElement } from './component';
import type { FindReplaceRequest, HyperlinkRequest } from './building-blocks-actions';

registerDocxEditor();
afterEach(() => document.body.replaceChildren());
it('mounts only the surface, edits the retained selection, serializes and enforces permission', async () => {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.surfaceOnly = true;
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text: 'Selected text' }] }];
	editor.documentModel = model;
	document.body.append(editor);
	const root = editor.shadowRoot!;
	expect(root.querySelector('office-ui-ribbon')).toBeNull();
	expect(root.querySelector('office-ui-title-bar')).toBeNull();
	expect(root.querySelector('office-ui-backstage')).toBeNull();
	expect(root.querySelector('.ProseMirror')).not.toBeNull();
	const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 9)));
	let updates = 0;
	const stop = editor.subscribeEditorState(() => updates++);
	const props = editor.getToolbarProps();
	expect(props.onBold()).toBe(true);
	expect(editor.getToolbarProps().bold).toBe(true);
	expect(editor.dirty).toBe(true);
	const saved = await loadDocx(await editor.saveBytes());
	expect(JSON.stringify(saved)).toContain('Selected');
	expect(JSON.stringify(saved)).toContain('"bold":true');
	editor.readOnly = true;
	expect(props.onItalic()).toBe(false);
	expect(editor.getToolbarProps().canEdit).toBe(false);
	expect(updates).toBeGreaterThan(0);
	stop();
	editor.markClean();
	expect(editor.dirty).toBe(false);
});

function textNodeOf(root: Node): Node {
	const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
	return walker.nextNode()!;
}

function mounted(text = 'Selected text', options: { readOnly?: boolean } = {}) {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	editor.surfaceOnly = true;
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text }] }];
	editor.documentModel = model;
	if (options.readOnly) editor.readOnly = true;
	document.body.append(editor);
	const view = (editor as unknown as { view: import('prosemirror-view').EditorView }).view;
	return { editor, view };
}

/** A selection inside a shadow tree, which jsdom's document selection refuses to hold. */
function fakeSelection(root: Node) {
	const selection = {
		anchorNode: null as Node | null,
		anchorOffset: 0,
		focusNode: null as Node | null,
		focusOffset: 0,
		set(anchor: Node, anchorOffset: number, focus: Node, focusOffset: number) {
			Object.assign(selection, { anchorNode: anchor, anchorOffset, focusNode: focus, focusOffset });
		},
		selectAllChildren(node: Node) {
			const text = textNodeOf(node);
			selection.set(text, 0, text, text.textContent!.length);
		},
		collapseToStart() {},
	};
	Object.assign(root, { getSelection: () => selection });
	return selection;
}

it('keeps a read-only surface focusable, selectable and copyable without allowing edits', () => {
	const { editor, view } = mounted('Read only words', { readOnly: true });
	const selection = fakeSelection(editor.shadowRoot!);
	const canvas = editor.shadowRoot!.querySelector<HTMLElement>('.dve-canvas')!;
	expect(canvas.tabIndex).toBe(-1);
	expect(view.editable).toBe(false);
	const text = textNodeOf(view.dom);
	selection.set(text, 0, text, 4);
	document.dispatchEvent(new Event('selectionchange'));
	const props = editor.getToolbarProps();
	expect(props.selection).toEqual({ from: 1, to: 5 });
	expect(props.canEdit).toBe(false);
	expect(props.bold).toBe(false);
	expect(props.onBold()).toBe(false);
	const before = JSON.stringify(view.state.doc.toJSON());
	view.dom.dispatchEvent(new KeyboardEvent('keydown', { key: 'x', bubbles: true }));
	expect(JSON.stringify(view.state.doc.toJSON())).toBe(before);
	expect(editor.dirty).toBe(false);
	// Mod+A selects the document through the browser selection, which is mirrored into the editor.
	const select = new KeyboardEvent('keydown', {
		key: 'a',
		ctrlKey: true,
		bubbles: true,
		cancelable: true,
	});
	canvas.dispatchEvent(select);
	expect(select.defaultPrevented).toBe(true);
	document.dispatchEvent(new Event('selectionchange'));
	expect(editor.getToolbarProps().selection).toEqual({ from: 1, to: 16 });
});

it('stops following the document selection once the element is disconnected', () => {
	const { editor, view } = mounted('Detach test', { readOnly: true });
	const selection = fakeSelection(editor.shadowRoot!);
	editor.remove();
	const text = textNodeOf(view.dom);
	selection.set(text, 2, text, 5);
	const before = view.state.selection.from;
	document.dispatchEvent(new Event('selectionchange'));
	expect(view.state.selection.from).toBe(before);
});

it('types list and alignment state with the values its callbacks accept', () => {
	const { editor, view } = mounted();
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 4)));
	const props = editor.getToolbarProps();
	expect(props.list).toBeNull();
	expect(props.alignment).toBe('left');
	expect(props.onListChange('number')).toBe(true);
	expect(editor.getToolbarProps().list).toBe('number');
	expect(editor.getToolbarProps().onListChange('bullet')).toBe(true);
	expect(editor.getToolbarProps().list).toBe('bullet');
	expect(editor.getToolbarProps().onListChange('remove')).toBe(true);
	expect(editor.getToolbarProps().list).toBeNull();
	expect(editor.getToolbarProps().onAlignmentChange('center')).toBe(true);
	expect(editor.getToolbarProps().alignment).toBe('center');
	expect(Object.keys(props.labels)).toEqual(
		expect.arrayContaining(['Font', 'Align center', 'Bulleted list', 'Insert link']),
	);
});

it('reports page, word and zoom state and drives them through callbacks', () => {
	const { editor } = mounted('One two three four');
	const status = editor.getStatusProps();
	expect(status.words).toBe(4);
	expect(status.zoom).toBe(100);
	expect(status.onZoomChange(150)).toBe(true);
	expect(editor.getStatusProps().zoom).toBe(150);
	expect(editor.getStatusProps().onZoomChange(100000)).toBe(true);
	expect(editor.getStatusProps().zoom).toBe(status.maxZoom);
	expect(editor.getStatusProps().onZoomChange(Number.NaN)).toBe(false);
	const navigation = editor.getNavigationProps();
	expect(navigation).toMatchObject({ page: 1, pageCount: 1, pagesAvailable: false });
	expect(navigation.canGoNext).toBe(false);
	expect(navigation.onNextPage()).toBe(false);
	expect(navigation.onGoToPage(2)).toBe(false);
	expect(navigation.onViewModeChange('print')).toBe(true);
	expect(editor.getNavigationProps().viewMode).toBe('print');
});

it('reports host workflows as not enabled until their callback is supplied', async () => {
	const { editor, view } = mounted('Find me please');
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 1, 5)));
	const none = editor.getToolbarProps();
	expect([none.canInsertLink, none.canFindReplace, none.canInsertPicture]).toEqual([
		false,
		false,
		false,
	]);
	expect([
		none.onOpenHyperlinkDialog(),
		none.onOpenFindReplace(),
		none.onOpenImagePicker(),
	]).toEqual([false, false, false]);
	let link: HyperlinkRequest | undefined;
	let find: FindReplaceRequest | undefined;
	let picker = false;
	editor.surfaceActions = {
		onOpenHyperlinkDialog: (request) => (link = request),
		onOpenFindReplace: (request) => (find = request),
		onOpenImagePicker: () => (picker = true),
	};
	const props = editor.getToolbarProps();
	expect([props.canInsertLink, props.canFindReplace, props.canInsertPicture]).toEqual([
		true,
		true,
		true,
	]);
	expect(props.onOpenHyperlinkDialog()).toBe(true);
	expect(link).toMatchObject({ href: null, text: 'Find' });
	expect(link!.apply('https://example.com')).toBe(true);
	expect(JSON.stringify(view.state.doc.toJSON())).toContain('https://example.com');
	expect(editor.getToolbarProps().onOpenFindReplace()).toBe(true);
	expect(find!.search('please').count).toBe(1);
	expect(find!.replaceAll('now')).toBe(1);
	expect(view.state.doc.textContent).toBe('Find me now');
	expect(editor.getToolbarProps().onOpenImagePicker()).toBe(true);
	expect(picker).toBe(true);
	// Read-only keeps find but closes the editing workflows, even for callbacks retained earlier.
	editor.readOnly = true;
	const retained = editor.getToolbarProps();
	expect([retained.canInsertLink, retained.canFindReplace, retained.canInsertPicture]).toEqual([
		false,
		true,
		false,
	]);
	expect(retained.onOpenHyperlinkDialog()).toBe(false);
	expect(retained.onOpenImagePicker()).toBe(false);
	expect(find!.canReplace).toBe(false);
	expect(find!.replaceAll('x')).toBe(0);
	expect(find!.search('now').count).toBe(1);
	expect(view.state.doc.textContent).toBe('Find me now');
});
