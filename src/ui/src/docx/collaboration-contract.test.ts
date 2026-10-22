// @vitest-environment jsdom
// The shared embedding contract on <docx-editor>: the `collaboration` property joins a room over
// the in-memory hub, reports state through `collaboration-change`, and cleans up by ownership.
import { afterEach, describe, expect, it } from 'vitest';
import { createDocument } from 'ooxml-core/docx';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from 'ooxml-core/collab';
import type { EditorView } from 'prosemirror-view';
import type { DocxCollaborationState } from './index';
import { DocxEditorElement } from './index';
import './index';

const wait = (ms = 20) => new Promise<void>((resolve) => setTimeout(resolve, ms));
const sessions: CollabSession[] = [];
const editors: DocxEditorElement[] = [];

afterEach(() => {
	for (const editor of editors.splice(0)) {
		editor.stopCollaboration(true);
		editor.remove();
	}
	for (const session of sessions.splice(0)) session.destroy();
});

function mount(): DocxEditorElement {
	const editor = document.createElement('docx-editor') as DocxEditorElement;
	document.body.append(editor);
	editors.push(editor);
	return editor;
}
const viewOf = (editor: DocxEditorElement) => (editor as unknown as { view: EditorView }).view;
const textOf = (editor: DocxEditorElement) => viewOf(editor).state.doc.textContent;

function hostSession(hub: ReturnType<typeof createMemoryHub>, name: string, role?: 'viewer') {
	const session = createCollabSession({
		roomId: 'word',
		provider: transportProvider({ transport: hub.createTransport('word') }),
		user: { name, ...(role ? { role } : {}) },
		heartbeatMs: 0,
		syncGraceMs: 0,
		teardown: false,
	});
	sessions.push(session);
	return session;
}

describe('docx-editor collaboration property', () => {
	it('is off by default and rejects nothing when cleared', () => {
		const editor = mount();
		expect(editor.collaboration).toBeNull();
		expect(editor.collaborationState.active).toBe(false);
		expect(editor.collaborationState.status).toBe('off');
		editor.collaboration = null;
		expect(editor.collaborationState.people).toEqual([]);
	});

	it('shares one document between two editors and reports state', async () => {
		const hub = createMemoryHub();
		const a = mount();
		const b = mount();
		const events: DocxCollaborationState[] = [];
		a.addEventListener('collaboration-change', (event) => events.push(event.detail));
		a.collaboration = { session: hostSession(hub, 'Ada'), intent: 'create' };
		await wait();
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('hello', 1));
		b.collaboration = { session: hostSession(hub, 'Bob'), intent: 'join' };
		await wait();
		expect(textOf(b)).toBe('hello');
		viewOf(b).dispatch(viewOf(b).state.tr.insertText('B', 1));
		await wait();
		expect(textOf(a)).toBe(textOf(b));
		expect(textOf(a)).toContain('B');

		const state = a.collaborationState;
		expect(state.active).toBe(true);
		expect(state.roomId).toBe('word');
		expect(state.status).toBe('connected');
		expect(state.people.map((person) => [person.name, person.self])).toEqual([
			['Ada', true],
			['Bob', false],
		]);
		expect(events.length).toBeGreaterThan(0);
		expect(events.some((detail) => detail.people.length === 2)).toBe(true);
	});

	it('leaves on null and on stopCollaboration, never destroying a host session', async () => {
		const hub = createMemoryHub();
		const a = mount();
		const session = hostSession(hub, 'Ada');
		const events: DocxCollaborationState[] = [];
		a.addEventListener('collaboration-change', (event) => events.push(event.detail));
		a.collaboration = { session };
		await wait();
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('kept', 1));
		a.collaboration = null;
		expect(a.collaborationState.active).toBe(false);
		expect(events[events.length - 1]?.status).toBe('off');
		expect(session.canWrite()).toBe(true);
		expect(textOf(a)).toBe('kept');

		a.collaboration = { session };
		await wait();
		expect(a.collaborationState.active).toBe(true);
		a.stopCollaboration();
		expect(a.collaboration).toBeNull();
		expect(a.collaborationState.active).toBe(false);
		expect(session.canWrite()).toBe(true);
	});

	it('opens and destroys its own session for a room name', async () => {
		const editor = mount();
		const errors: Error[] = [];
		editor.addEventListener('document-error', (event) => errors.push(event.detail));
		editor.collaboration = { roomId: 'own-room', intent: 'create' };
		const opened = editor.collaborationState;
		expect(opened.active).toBe(true);
		expect(opened.roomId).toBe('own-room');
		editor.collaboration = null;
		expect(editor.collaborationState.active).toBe(false);
		expect(errors).toEqual([]);
	});

	it('reports an invalid room as document-error without throwing', async () => {
		const editor = mount();
		const errors: Error[] = [];
		editor.addEventListener('document-error', (event) => errors.push(event.detail));
		editor.collaboration = { roomId: 'not a room!' };
		await wait();
		expect(errors).toHaveLength(1);
		expect(editor.collaborationState.active).toBe(false);
	});

	it('leaves before a document is replaced and rejoins after', async () => {
		const hub = createMemoryHub();
		const a = mount();
		a.collaboration = { session: hostSession(hub, 'Ada') };
		await wait();
		expect(a.collaborationState.active).toBe(true);
		a.documentModel = createDocument();
		await wait();
		expect(a.collaborationState.active).toBe(true);
		expect(a.collaboration).not.toBeNull();
	});

	it('keeps the imperative Yjs API reporting state', async () => {
		const hub = createMemoryHub();
		const a = mount();
		const b = mount();
		a.collaboration = { session: hostSession(hub, 'Ada') };
		await wait();
		const events: DocxCollaborationState[] = [];
		b.addEventListener('collaboration-change', (event) => events.push(event.detail));
		b.startYjsCollaboration(hostSession(hub, 'Bob'), { documentId: 'word' });
		expect(b.collaborationState.active).toBe(true);
		expect(b.collaborationState.people).toHaveLength(2);
		expect(b.collaboration).toBeNull();
		expect(events.length).toBeGreaterThan(0);
		b.stopCollaboration();
		expect(b.collaborationState.active).toBe(false);
	});

	it('a viewer role is read-only like readOnly: the host cannot lift it and leaving restores it', async () => {
		const hub = createMemoryHub();
		const owner = mount();
		owner.collaboration = { session: hostSession(hub, 'Ada') };
		await wait();
		const a = mount();
		const events: { readOnly: boolean; locked: boolean; hostReadOnly: boolean }[] = [];
		a.addEventListener('readonly-change', (event) =>
			events.push((event as CustomEvent<(typeof events)[number]>).detail),
		);
		a.collaboration = { session: hostSession(hub, 'Vic', 'viewer'), intent: 'join' };
		await wait();
		expect(a.readOnly).toBe(true);
		expect(events).toEqual([{ readOnly: true, locked: true, hostReadOnly: false }]);
		a.readOnly = false;
		expect(a.readOnly).toBe(true);
		// The File-menu toggle reports the effective value only when it changed, plus the host switch.
		const shell = (
			a as unknown as { core: { shell: { chrome: { host: { setReadOnly(v: boolean): void } } } } }
		).core.shell.chrome.host;
		events.length = 0;
		shell.setReadOnly(true);
		expect(events).toEqual([{ readOnly: true, locked: true, hostReadOnly: true }]);
		shell.setReadOnly(true);
		expect(events).toHaveLength(1);
		shell.setReadOnly(false);
		events.length = 0;
		a.stopCollaboration();
		expect(a.readOnly).toBe(false);
		expect(events).toEqual([{ readOnly: false, locked: false, hostReadOnly: false }]);
		a.readOnly = true;
		expect(a.readOnly).toBe(true);
	});

	it('a viewer role blocks local edits', async () => {
		const hub = createMemoryHub();
		const a = mount();
		const b = mount();
		a.collaboration = { session: hostSession(hub, 'Ada') };
		await wait();
		viewOf(a).dispatch(viewOf(a).state.tr.insertText('text', 1));
		b.collaboration = { session: hostSession(hub, 'Vic', 'viewer'), intent: 'join' };
		await wait();
		const before = textOf(b);
		viewOf(b).dispatch(viewOf(b).state.tr.insertText('X', 1));
		await wait();
		expect(textOf(b)).toBe(before);
		expect(textOf(a)).toBe(before);
	});
});
