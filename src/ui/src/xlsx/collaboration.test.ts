// @vitest-environment jsdom
// Two real <xlsx-editor> elements sharing one room over the collab area's in-memory hub: edits
// travel, undo is per window, selections become outlines in the other grid, and stopping cleans up.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	UPDATE_REFUSAL_MESSAGES,
	type CollabSession,
} from 'ooxml-core/collab';
import { collabStrings } from 'ooxml-core/i18n';
import { getCell } from 'ooxml-core/xlsx';
import { sanitizeXlsxPresence, type XlsxPresence } from 'ooxml-core/xlsx/collab';
import type { EditorCore } from './editor-core';
import { defineXlsxEditor, type XlsxEditorElement } from './index';
import type { OfficeUiShareDialog } from '../dialog/share-dialog';

const wait = (ms = 20) => new Promise<void>((resolve) => setTimeout(resolve, ms));
const sessions: CollabSession<XlsxPresence>[] = [];

beforeAll(() => defineXlsxEditor());
afterEach(() => {
	for (const session of sessions.splice(0)) session.destroy();
	document.body.replaceChildren();
});

function join(hub: ReturnType<typeof createMemoryHub>, name: string, color: string) {
	const session = createCollabSession<XlsxPresence>({
		roomId: 'book',
		provider: transportProvider({ transport: hub.createTransport('book') }),
		user: { name, color },
		initialPresence: {},
		sanitizePayload: sanitizeXlsxPresence,
		heartbeatMs: 0,
		syncGraceMs: 0,
	});
	sessions.push(session);
	return session;
}

function editor() {
	const element = document.createElement('xlsx-editor') as XlsxEditorElement;
	document.body.append(element);
	element.newWorkbook();
	const core = (element as unknown as { core: EditorCore }).core;
	return { element, core, edit: () => core.session! };
}

async function pair() {
	const hub = createMemoryHub();
	const a = editor();
	const b = editor();
	a.edit().setCellInput(0, 0, 0, 'seed');
	await a.element.startCollaboration({ session: join(hub, 'Ada', '#2563eb') });
	await wait();
	await b.element.startCollaboration({ session: join(hub, 'Bob', '#dc2626') });
	await wait();
	return { hub, a, b };
}

const value = (side: ReturnType<typeof editor>, row: number, col: number) =>
	getCell(side.core.workbook!.sheets[0]!, row, col)?.value;

describe('<xlsx-editor> collaboration', () => {
	it('a viewer role stays read-only whatever the host sets, and leaving restores the host', async () => {
		const hub = createMemoryHub();
		const a = editor();
		const viewer = createCollabSession<XlsxPresence>({
			roomId: 'book',
			provider: transportProvider({ transport: hub.createTransport('book') }),
			user: { name: 'Vera', color: '#2563eb', role: 'viewer' },
			initialPresence: {},
			sanitizePayload: sanitizeXlsxPresence,
			heartbeatMs: 0,
			syncGraceMs: 0,
		});
		sessions.push(viewer);
		await a.element.startCollaboration({ session: viewer });
		await wait();
		expect(a.element.readOnly).toBe(true);
		a.element.readOnly = false;
		expect(a.element.readOnly).toBe(true);
		a.element.stopCollaboration();
		expect(a.element.readOnly).toBe(false);
		a.element.readOnly = true;
		expect(a.element.readOnly).toBe(true);
	});

	it('a customization never lifts a viewer role, and leaving keeps edit: false', async () => {
		const hub = createMemoryHub();
		const a = editor();
		const viewer = createCollabSession<XlsxPresence>({
			roomId: 'book',
			provider: transportProvider({ transport: hub.createTransport('book') }),
			user: { name: 'Vera', color: '#2563eb', role: 'viewer' },
			initialPresence: {},
			sanitizePayload: sanitizeXlsxPresence,
			heartbeatMs: 0,
			syncGraceMs: 0,
		});
		sessions.push(viewer);
		await a.element.startCollaboration({ session: viewer });
		await wait();
		a.element.customization = { capabilities: { print: false } };
		a.element.customization = null;
		expect(a.element.readOnly).toBe(true);
		a.element.customization = { capabilities: { edit: false } };
		a.element.stopCollaboration();
		expect(a.element.readOnly).toBe(true);
		a.element.readOnly = false;
		expect(a.element.readOnly).toBe(true);
		const events: unknown[] = [];
		a.element.addEventListener('readonly-change', (event) =>
			events.push((event as CustomEvent).detail),
		);
		a.element.customization = null;
		expect(a.element.readOnly).toBe(false);
		expect(events).toEqual([{ readOnly: false, locked: false, hostReadOnly: false }]);
	});

	it('keeps a read-only attribute set under edit: false', () => {
		const a = editor();
		a.element.customization = { capabilities: { edit: false } };
		a.element.setAttribute('read-only', '');
		a.element.customization = null;
		expect(a.element.readOnly).toBe(true);
		a.element.removeAttribute('read-only');
		expect(a.element.readOnly).toBe(false);
	});

	it('explains a refused update in the display language and stops sharing', async () => {
		const { hub, a } = await pair();
		a.element.locale = 'de';
		const errors: string[] = [];
		a.element.addEventListener('workbook-error', (event) =>
			errors.push((event as CustomEvent<{ message: string }>).detail.message),
		);
		const evil = hub.createTransport('book');
		evil.connect({ open: () => {}, close: () => {}, message: () => {}, error: () => {} });
		// A whole sync update whose struct has an unknown content type.
		evil.send(new Uint8Array([0, 2, 5, 1, 1, 1, 0, 0x1f]));
		await wait();
		const german = collabStrings.de[UPDATE_REFUSAL_MESSAGES.malformed]!;
		const root = a.element.shadowRoot ?? a.element;
		const toasts = [...root.querySelectorAll('.xve-toast-text')].map((node) => node.textContent);
		expect(toasts).toContain(german);
		expect(toasts.some((text) => text?.startsWith('Verbindung'))).toBe(false);
		expect(errors).toContain(german);
		expect(a.element.collaborationState).toMatchObject({ status: 'error', refusal: 'malformed' });
	});

	it('adopts the room and carries edits both ways', async () => {
		const { a, b } = await pair();
		expect(value(b, 0, 0)).toBe('seed');
		a.edit().setCellInput(0, 1, 0, '41');
		b.edit().setCellInput(0, 1, 1, '=A2+1');
		await wait();
		expect(value(b, 1, 0)).toBe(41);
		expect(value(a, 1, 1)).toBe(42);
		expect(a.element.collaborationState).toMatchObject({ active: true, roomId: 'book' });
		expect(a.element.collaborationState.people.map((p) => [p.name, p.self])).toEqual([
			['Ada', true],
			['Bob', false],
		]);
	});

	it('undoes only this window edits, through the binding', async () => {
		const { a, b } = await pair();
		b.edit().setCellInput(0, 3, 1, 'Bob was here');
		a.edit().setCellInput(0, 3, 0, 'Ada was here');
		await wait();
		const sessionUndo = a.edit().undo;
		let sessionUndos = 0;
		a.edit().undo = () => {
			sessionUndos++;
			return sessionUndo();
		};
		// The ribbon, Quick Access Toolbar and Ctrl+Z all run this command.
		a.element.undo();
		await wait();
		expect(sessionUndos).toBe(0);
		for (const side of [a, b]) {
			expect(value(side, 3, 0)).toBeUndefined();
			expect(value(side, 3, 1)).toBe('Bob was here');
		}
		a.element.redo();
		await wait();
		expect(value(b, 3, 0)).toBe('Ada was here');
	});

	it('publishes the selection and outlines collaborators in the grid', async () => {
		const { a, b } = await pair();
		a.element.select('C3:D4');
		await wait(120);
		expect(b.core.ctx.remoteSelections?.()).toMatchObject([
			{
				userName: 'Ada',
				userColor: '#2563eb',
				sheet: 0,
				range: { start: { row: 2, col: 2 }, end: { row: 3, col: 3 } },
			},
		]);
		await wait(60);
		const outline = b.element.shadowRoot!.querySelector<HTMLElement>('.xg-remote:not([hidden])');
		expect(outline?.querySelector('.xg-remote-tag')?.textContent).toBe('Ada');
		expect(outline?.title).toBe('Ada');
		const people = b.element.shadowRoot!.querySelector('.xve-title-people') as HTMLElement & {
			participants: { name: string }[];
		};
		expect(people.hidden).toBe(false);
		expect(people.participants.map((p) => p.name).sort()).toEqual(['Ada', 'Bob']);
	});

	it('stops cleanly: edits stay local and remote outlines disappear', async () => {
		const { a, b } = await pair();
		const events: unknown[] = [];
		a.element.addEventListener('collaboration-change', (event) => events.push(event.detail));
		a.element.stopCollaboration();
		expect(a.element.collaborationState.active).toBe(false);
		expect(events.at(-1)).toMatchObject({ active: false, status: 'off', people: [] });
		expect(a.core.ctx.remoteSelections?.()).toEqual([]);
		a.edit().setCellInput(0, 5, 0, 'private');
		await wait();
		expect(value(b, 5, 0)).toBeUndefined();
		// Back on the session's own history.
		a.element.undo();
		expect(value(a, 5, 0)).toBeUndefined();
	});

	it('rejoins with a newly opened workbook while `collaboration` is set', async () => {
		const hub = createMemoryHub();
		const a = editor();
		const b = editor();
		a.edit().setCellInput(0, 0, 0, 'room');
		a.element.collaboration = { session: join(hub, 'Ada', '#2563eb') };
		await wait();
		b.element.collaboration = { session: join(hub, 'Bob', '#dc2626') };
		await wait();
		expect(value(b, 0, 0)).toBe('room');
		b.element.collaboration = null;
		expect(b.element.collaborationState.active).toBe(false);
		await expect(b.element.startCollaboration({ roomId: 'not a room!' })).rejects.toThrow(
			/128 letters/,
		);
	});

	it('shares from the Share dialog, and File > Share opens it', async () => {
		const a = editor();
		a.element.authorName = 'Ada';
		a.element.share();
		const root = a.element.shadowRoot!;
		const dialog = root.querySelector<OfficeUiShareDialog>('office-ui-share-dialog')!;
		expect(dialog.open).toBe(true);
		dialog.ensureRendered();
		expect(dialog.formOptions()).toEqual({
			roomId: expect.stringMatching(/^book-[a-z0-9]{8}$/),
			user: { name: 'Ada' },
			intent: 'create',
		});
		// A failed start is shown in the dialog, and nothing is shared.
		dialog.dispatchEvent(
			new CustomEvent('office-share-start', { detail: { options: { roomId: 'bad room' } } }),
		);
		await wait();
		expect(dialog.error).toMatch(/128 letters/);
		expect(a.element.collaborationState.active).toBe(false);
		dialog.open = false;
		// File > Share's "Share with people" runs the same command.
		const announced: string[] = [];
		a.element.addEventListener('file-command', (event) =>
			announced.push((event as CustomEvent<{ command: string }>).detail.command),
		);
		const shell = a.element as unknown as { shell: { backstage: { open(page: string): void } } };
		shell.shell.backstage.open('share');
		root
			.querySelector<HTMLButtonElement>('[data-backstage-page="share"] [data-share="open"]')!
			.click();
		await wait();
		expect(announced).toEqual(['share']);
		expect(dialog.open).toBe(true);
	});
});

describe('<xlsx-editor> collaboration while removed', () => {
	const until = async (check: () => boolean) => {
		for (let i = 0; i < 100 && !check(); i++) await wait(20);
		expect(check()).toBe(true);
	};

	it('keeps a room it opened through a move, leaves it while removed and rejoins when back', async () => {
		const roomId = `moved-${Math.random().toString(36).slice(2, 10)}`;
		const a = editor();
		const b = editor();
		a.edit().setCellInput(0, 0, 0, 'room');
		a.element.collaboration = { roomId, user: { name: 'Ada' } };
		await until(() => a.element.collaborationState.active);
		b.element.collaboration = { roomId, intent: 'join', user: { name: 'Bob' } };
		const people = () => b.element.collaborationState.people.length;
		await until(() => people() === 2 && value(b, 0, 0) === 'room');
		const other = document.createElement('div');
		document.body.append(other);
		other.append(a.element);
		await wait();
		expect(a.element.collaborationState.active).toBe(true);
		a.element.remove();
		await wait();
		expect(a.element.collaborationState.active).toBe(false);
		await until(() => people() === 1);
		b.edit().setCellInput(0, 1, 0, 'while away');
		document.body.append(a.element);
		await until(() => people() === 2 && value(a, 1, 0) === 'while away');
		a.element.collaboration = null;
		b.element.collaboration = null;
	});

	it('rejoins a room started with startCollaboration once attached again', async () => {
		const hub = createMemoryHub();
		const a = editor();
		await a.element.startCollaboration({ session: join(hub, 'Ada', '#2563eb') });
		a.element.remove();
		await wait();
		expect(a.element.collaborationState.active).toBe(false);
		document.body.append(a.element);
		await until(() => a.element.collaborationState.active);
		a.element.stopCollaboration();
		a.element.remove();
		document.body.append(a.element);
		await wait();
		expect(a.element.collaborationState.active).toBe(false);
	});
});

describe('two Y.Docs synced by hand', () => {
	it('converge through Y.applyUpdate', async () => {
		const a = editor();
		const b = editor();
		const docs = [new Y.Doc(), new Y.Doc()] as const;
		const relay = (from: Y.Doc, to: Y.Doc) =>
			from.on('update', (update: Uint8Array, origin: unknown) => {
				if (origin !== 'relay') Y.applyUpdate(to, update, 'relay');
			});
		relay(docs[0], docs[1]);
		relay(docs[1], docs[0]);
		const hub = createMemoryHub();
		const session = (doc: Y.Doc, name: string) => {
			const s = createCollabSession<XlsxPresence>({
				roomId: `solo-${name}`,
				doc,
				provider: transportProvider({ transport: hub.createTransport(`solo-${name}`) }),
				user: { name },
				sanitizePayload: sanitizeXlsxPresence,
				heartbeatMs: 0,
				syncGraceMs: 0,
			});
			sessions.push(s);
			return s;
		};
		a.edit().setCellInput(0, 0, 0, '7');
		await a.element.startCollaboration({ session: session(docs[0], 'Ada') });
		await wait();
		await b.element.startCollaboration({ session: session(docs[1], 'Bob') });
		await wait();
		expect(value(b, 0, 0)).toBe(7);
		b.edit().setCellInput(0, 0, 1, '=A1*6');
		await wait();
		expect(value(a, 0, 1)).toBe(42);
	});
});
