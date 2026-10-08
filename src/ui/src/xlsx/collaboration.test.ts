// @vitest-environment jsdom
// Two real <xlsx-editor> elements sharing one room over the collab area's in-memory hub: edits
// travel, undo is per window, selections become outlines in the other grid, and stopping cleans up.
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollabSession,
} from 'ooxml-core/collab';
import { getCell } from 'ooxml-core/xlsx';
import { sanitizeXlsxPresence, type XlsxPresence } from 'ooxml-core/xlsx/collab';
import type { EditorCore } from './editor-core';
import { defineXlsxEditor, type XlsxEditorElement } from './index';

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

	it('shares from File > Share', async () => {
		const a = editor();
		a.element.share();
		const page = a.element.shadowRoot!.querySelector<HTMLElement>('[data-backstage-page="share"]')!;
		expect(page.hidden).toBe(false);
		const room = page.querySelector<HTMLInputElement>('.xve-share-room')!;
		expect(room.value).toMatch(/^book-[a-z0-9]+$/);
		room.value = 'bad room';
		page.querySelector<HTMLButtonElement>('[data-share="start"]')!.click();
		await wait();
		expect(page.querySelector('.xve-share-status')?.textContent).toMatch(/128 letters/);
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
