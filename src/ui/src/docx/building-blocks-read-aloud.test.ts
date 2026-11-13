// @vitest-environment jsdom
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { TextSelection } from 'prosemirror-state';
import { createDocument } from 'ooxml-core/docx';
import { registerDocxEditor } from './index';
import type { DocxEditorElement } from './component';
import { stopReadAloud } from './read-aloud';

registerDocxEditor();

interface FakeUtterance {
	text: string;
	listeners: Record<string, (event?: unknown) => void>;
}
let spoken: FakeUtterance[] = [];
let cancel: ReturnType<typeof vi.fn>;

beforeEach(() => {
	spoken = [];
	cancel = vi.fn();
	vi.stubGlobal('speechSynthesis', { speak: (u: FakeUtterance) => spoken.push(u), cancel });
	vi.stubGlobal(
		'SpeechSynthesisUtterance',
		class {
			listeners: FakeUtterance['listeners'] = {};
			constructor(public text: string) {}
			addEventListener(name: string, fn: () => void) {
				this.listeners[name] = fn;
			}
		},
	);
});
afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

function mounted(text = 'Alpha beta gamma', options: { readOnly?: boolean } = {}) {
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
const select = (view: import('prosemirror-view').EditorView, from: number, to: number) =>
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, from, to)));

it('reads the selection, then the caret to the end, and reports playback state', () => {
	const { editor, view } = mounted();
	select(view, 1, 6);
	let updates = 0;
	editor.subscribeEditorState(() => updates++);
	expect(editor.getToolbarProps()).toMatchObject({ canReadAloud: true, readingAloud: false });
	expect(editor.getToolbarProps().onReadAloud()).toBe('started');
	expect(spoken.map((u) => u.text)).toEqual(['Alpha']);
	expect(editor.getToolbarProps().readingAloud).toBe(true);
	expect(updates).toBeGreaterThan(0);
	expect(editor.getToolbarProps().onStopReadAloud()).toBe(true);
	expect(cancel).toHaveBeenCalled();
	expect(editor.getToolbarProps().readingAloud).toBe(false);
	select(view, 7, 7);
	editor.getToolbarProps().onReadAloud();
	expect(spoken.at(-1)?.text).toBe('beta gamma');
	spoken.at(-1)!.listeners.end!();
	expect(editor.getToolbarProps().readingAloud).toBe(false);
});

it('plays in a read-only document without dirtying it or moving the selection', () => {
	const { editor, view } = mounted('Alpha beta', { readOnly: true });
	select(view, 1, 6);
	const retained = editor.getToolbarProps();
	expect(retained.canEdit).toBe(false);
	expect(retained.canReadAloud).toBe(true);
	expect(retained.onReadAloud()).toBe('started');
	expect(editor.dirty).toBe(false);
	expect(view.state.selection.from).toBe(1);
	expect(view.state.selection.to).toBe(6);
	expect(retained.onStopReadAloud()).toBe(true);
});

it('reports an unsupported environment, an empty range and speech errors', () => {
	const { editor, view } = mounted('Hi');
	select(view, 3, 3);
	expect(editor.getToolbarProps().onReadAloud()).toBe('empty');
	expect(editor.getToolbarProps().readAloudIssue).toBe('empty');
	select(view, 1, 3);
	expect(editor.getToolbarProps().onReadAloud()).toBe('started');
	expect(editor.getToolbarProps().readAloudIssue).toBeNull();
	spoken[0]!.listeners.error!({ error: 'synthesis-failed' });
	expect(editor.getToolbarProps()).toMatchObject({ readingAloud: false, readAloudIssue: 'error' });
	// A cancel is not a failure.
	editor.getToolbarProps().onReadAloud();
	spoken[1]!.listeners.error!({ error: 'canceled' });
	expect(editor.getToolbarProps().readAloudIssue).toBeNull();
	vi.stubGlobal('speechSynthesis', undefined);
	expect(editor.getToolbarProps()).toMatchObject({
		canReadAloud: false,
		readAloudIssue: 'unsupported',
	});
	expect(editor.getToolbarProps().onReadAloud()).toBe('unsupported');
});

it('lets one editor own the voice at a time', () => {
	const first = mounted('First one');
	const second = mounted('Second one');
	first.editor.getToolbarProps().onReadAloud();
	expect(first.editor.getToolbarProps().readingAloud).toBe(true);
	second.editor.getToolbarProps().onReadAloud();
	expect(first.editor.getToolbarProps().readingAloud).toBe(false);
	expect(second.editor.getToolbarProps().readingAloud).toBe(true);
	// The first cannot stop what it no longer owns.
	expect(first.editor.getToolbarProps().onStopReadAloud()).toBe(false);
	expect(second.editor.getToolbarProps().readingAloud).toBe(true);
	stopReadAloud((second.editor as unknown as { core: { readAloud: object } }).core.readAloud);
});

it('keeps playing through an ordinary edit by default', () => {
	const { editor, view } = mounted('Alpha beta');
	select(view, 1, 3);
	editor.getToolbarProps().onReadAloud();
	view.dispatch(view.state.tr.insertText('x', 1));
	expect(editor.getToolbarProps().readingAloud).toBe(true);
	expect(editor.getToolbarProps().onStopReadAloud()).toBe(true);
});

it('stops on an ordinary edit when readAloud.stopOnEdit is set', () => {
	const { editor, view } = mounted('Alpha beta');
	editor.readAloud = { stopOnEdit: true };
	expect(editor.readAloud).toEqual({ stopOnEdit: true });
	select(view, 1, 3);
	editor.getToolbarProps().onReadAloud();
	view.dispatch(view.state.tr.insertText('x', 1));
	expect(editor.getToolbarProps().readingAloud).toBe(false);
	editor.readAloud = null;
	select(view, 1, 3);
	editor.getToolbarProps().onReadAloud();
	view.dispatch(view.state.tr.insertText('y', 1));
	expect(editor.getToolbarProps().readingAloud).toBe(true);
});

it('always stops when the document is replaced or the editor unmounts', () => {
	const { editor, view } = mounted('Alpha beta');
	select(view, 1, 3);
	editor.getToolbarProps().onReadAloud();
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p2', runs: [{ text: 'Other' }] }];
	editor.documentModel = model;
	expect(editor.getToolbarProps().readingAloud).toBe(false);
	editor.getToolbarProps().onReadAloud();
	expect(editor.getToolbarProps().readingAloud).toBe(true);
	editor.remove();
	expect(editor.getToolbarProps().readingAloud).toBe(false);
	expect(cancel).toHaveBeenCalled();
});
