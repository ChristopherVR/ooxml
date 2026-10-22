// @vitest-environment jsdom
// Collaboration needs no application UI: two surface-only editors share a room, the toolbar props
// report who is in it, and undo through the props undoes only this window's edits.
import { afterEach, beforeAll, expect, it } from 'vitest';
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

function join(hub: ReturnType<typeof createMemoryHub>, name: string) {
	const session = createCollabSession<XlsxPresence>({
		roomId: 'book',
		provider: transportProvider({ transport: hub.createTransport('book') }),
		user: { name, color: '#2563eb' },
		initialPresence: {},
		sanitizePayload: sanitizeXlsxPresence,
		heartbeatMs: 0,
		syncGraceMs: 0,
	});
	sessions.push(session);
	return session;
}
function surface() {
	const element = document.createElement('xlsx-editor') as XlsxEditorElement;
	element.surfaceOnly = true;
	document.body.append(element);
	element.newWorkbook();
	return { element, core: (element as unknown as { core: EditorCore }).core };
}

it('shares a workbook between surface-only editors and reports it in the props', async () => {
	const hub = createMemoryHub();
	const a = surface();
	const b = surface();
	a.core.session!.setCellInput(0, 0, 0, 'seed');
	await a.element.startCollaboration({ session: join(hub, 'Ada') });
	await wait();
	await b.element.startCollaboration({ session: join(hub, 'Bob') });
	await wait();
	expect(getCell(b.core.workbook!.sheets[0]!, 0, 0)?.value).toBe('seed');
	b.core.session!.setCellInput(0, 1, 0, '7');
	await wait();
	expect(getCell(a.core.workbook!.sheets[0]!, 1, 0)?.value).toBe(7);
	const state = a.element.getToolbarProps().collaboration;
	expect(state).toMatchObject({ active: true, roomId: 'book' });
	expect(state.people.map((p) => p.name).sort()).toEqual(['Ada', 'Bob']);
	// Undo through the props reverts only this window's edit.
	a.core.session!.setCellInput(0, 2, 0, '9');
	await wait();
	expect(await a.element.getToolbarProps().onUndo()).toBe(true);
	await wait();
	expect(getCell(a.core.workbook!.sheets[0]!, 2, 0)?.value).toBeUndefined();
	expect(getCell(a.core.workbook!.sheets[0]!, 1, 0)?.value).toBe(7);
	a.element.stopCollaboration();
	expect(a.element.getToolbarProps().collaboration.active).toBe(false);
});
