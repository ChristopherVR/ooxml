import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollaborationState,
} from 'ooxml-core/collab';
import { createBackstage } from './backstage';
import { ViewerController } from './controller';
import { demoDocument } from 'ooxml-core/visio/ui';
import { registerViewerControls } from './office-ui';
import { ViewerShare } from './viewer-share';
import type { CancellableEditor } from './worker-editor';

const disposers: (() => void)[] = [];
afterEach(() => {
	disposers.splice(0).forEach((dispose) => dispose());
	document.body.replaceChildren();
});

const until = async (check: () => boolean) => {
	for (let i = 0; i < 100 && !check(); i++) await new Promise((r) => setTimeout(r, 10));
	expect(check()).toBe(true);
};

function window(name: string) {
	registerViewerControls();
	const parsed: number[][] = [];
	const editor: CancellableEditor = async (bytes) => ({
		bytes: new Uint8Array([...bytes, 7]),
		document: structuredClone(demoDocument),
		changedParts: ['visio/pages/page1.xml'],
		diagnostics: [],
	});
	const controller = new ViewerController(
		async (bytes) => {
			parsed.push([...new Uint8Array(bytes)]);
			return structuredClone(demoDocument);
		},
		() => {},
		editor,
	);
	const host = document.createElement('div');
	document.body.append(host);
	const root = host.attachShadow({ mode: 'open' });
	root.append(createBackstage(document));
	const share = new ViewerShare(root, controller, () => ({
		displayName: name,
		avatarColor: '#16a34a',
	}));
	const dispose = share.wire();
	disposers.push(() => {
		dispose();
		controller.destroy();
	});
	const people = () =>
		(
			root.querySelector('office-ui-presence') as HTMLElement & {
				participants: { name: string }[];
			}
		).participants.map((person) => person.name);
	return { controller, share, root, parsed, people };
}

describe('File > Share', () => {
	it('rejects session names the core would refuse', async () => {
		const a = window('Ada');
		await a.share.start('no spaces');
		expect(a.share.active).toBe(false);
		expect(a.root.querySelector('.share-status')!.textContent).toMatch(/letters, digits/);
	});

	it('shares the drawing with another window and applies its edits as history steps', async () => {
		const room = `test-${Math.random().toString(36).slice(2, 8)}`;
		const a = window('Ada');
		await a.controller.load(new Uint8Array([1]));
		await a.share.start(room);
		expect(a.share.active).toBe(true);
		expect(a.root.querySelector<HTMLElement>('[data-share="stop"]')!.hidden).toBe(false);

		const b = window('Grace');
		await b.share.start(room);
		// The joiner had nothing open: it loads the room's drawing.
		await until(() => b.parsed.some((bytes) => bytes.join() === '1'));
		await until(() => b.people().includes('Ada') && a.people().includes('Grace'));

		await a.controller.applyEdits([]);
		await until(() => b.parsed.some((bytes) => bytes.join() === '1,7'));
		await until(() => !b.controller.state.edit.busy && b.controller.state.edit.canUndo);
		expect(b.controller.exportVsdx().bytes).toEqual(new Uint8Array([1, 7]));

		// The remote step is undoable, and the undo is shared back.
		await b.controller.undo();
		await until(() => a.parsed.some((bytes) => bytes.join() === '1'));
		b.share.stop();
		expect(b.people()).toEqual([]);
	});

	it('shares a drawing loaded after joining even when an earlier listener changes state on load', async () => {
		const room = `test-${Math.random().toString(36).slice(2, 8)}`;
		const a = window('Ada');
		// A host that fits the page on load changes the zoom inside the load event, which makes the
		// controller drop that event for listeners registered later, such as the share session.
		a.controller.onEvent((name) => {
			if (name === 'document-load') a.controller.setZoom(2);
		});
		await a.share.start(room);
		const b = window('Grace');
		await b.share.start(room);
		await until(() => b.people().includes('Ada') && a.people().includes('Grace'));

		await a.controller.load(new Uint8Array([4, 2]));
		expect(a.controller.state.zoom).toBe(2);
		await until(() => b.parsed.some((bytes) => bytes.join() === '4,2'));
		// A later local edit still reaches the peer, and neither window echoes it back to A.
		await a.controller.applyEdits([]);
		await until(() => b.parsed.some((bytes) => bytes.join() === '4,2,7'));
		expect(a.parsed.map((bytes) => bytes.join())).toEqual(['4,2']);
	});
});

describe('the collaboration contract on the share controller', () => {
	const hostSession = (hub: ReturnType<typeof createMemoryHub>, name: string, role?: 'viewer') =>
		createCollabSession({
			roomId: 'contract',
			provider: transportProvider({ transport: hub.createTransport('contract') }),
			user: { name, ...(role ? { role } : {}) },
			heartbeatMs: 0,
			syncGraceMs: 0,
		});

	it('converges two windows through host sessions and never destroys them', async () => {
		const hub = createMemoryHub();
		const sessions = [hostSession(hub, 'Ada'), hostSession(hub, 'Grace')];
		const [first, second] = sessions as [(typeof sessions)[0], (typeof sessions)[0]];
		const destroyed = sessions.map((session) => vi.spyOn(session, 'destroy'));
		const a = window('Ada');
		const b = window('Grace');
		const states: CollaborationState[] = [];
		a.share.onChange = (state) => states.push(state);
		await a.controller.load(new Uint8Array([5]));
		const options = { session: first };
		a.share.want(options);
		await until(() => a.share.active);
		expect(a.share.wanted).toBe(options);
		b.share.want({ session: second });
		await until(() => b.parsed.some((bytes) => bytes.join() === '5'));
		await until(() => a.share.state.people.length === 2);
		expect(a.share.state).toMatchObject({ active: true, roomId: 'contract' });
		expect(a.share.state.people.map((person) => person.name).sort()).toEqual(['Ada', 'Grace']);
		expect(a.share.state.people[0]?.self).toBe(true);
		expect(states.at(-1)?.active).toBe(true);
		expect(a.people()).toContain('Grace');

		a.share.stop();
		expect(a.share.wanted).toBeNull();
		expect(a.share.state.active).toBe(false);
		expect(states.at(-1)).toMatchObject({ active: false, status: 'off', people: [] });
		expect(a.root.querySelector('.share-status')!.textContent).toBe('Not sharing.');
		b.share.stop();
		expect(destroyed[0]).not.toHaveBeenCalled();
		expect(destroyed[1]).not.toHaveBeenCalled();
		for (const session of sessions) session.destroy();
	});

	it('keeps a panel-started room reported and owned by the controller', async () => {
		const a = window('Ada');
		const states: CollaborationState[] = [];
		a.share.onChange = (state) => states.push(state);
		const room = `panel-${Math.random().toString(36).slice(2, 8)}`;
		await a.share.start(room);
		expect(a.share.state).toMatchObject({ active: true, roomId: room });
		expect(a.share.wanted).toBeNull();
		expect(states.at(-1)?.active).toBe(true);
		a.root.querySelector<HTMLElement>('[data-share="stop"]')!.click();
		expect(a.share.state.active).toBe(false);
		expect(states.at(-1)?.active).toBe(false);
	});

	it('does not seed an empty room when joining, and a viewer role is read-only', async () => {
		const hub = createMemoryHub();
		const session = hostSession(hub, 'Vera', 'viewer');
		const a = window('Vera');
		await a.controller.load(new Uint8Array([8]));
		a.share.want({ session, intent: 'join' });
		await until(() => a.share.active);
		await new Promise((r) => setTimeout(r, 30));
		expect(session.doc.getMap('package').size).toBe(0);
		expect(a.controller.canEdit).toBe(false);
		a.share.stop();
		expect(a.controller.canEdit).toBe(true);
		session.destroy();
	});

	it('abandons a join that is stopped while the runtime loads', async () => {
		const hub = createMemoryHub();
		const session = hostSession(hub, 'Ada');
		const a = window('Ada');
		a.share.want({ session });
		a.share.stop();
		await new Promise((r) => setTimeout(r, 30));
		expect(a.share.active).toBe(false);
		session.destroy();
	});

	it('reports a failed declarative join', async () => {
		const a = window('Ada');
		const failures: Error[] = [];
		a.share.onError = (error) => failures.push(error);
		a.share.want({ roomId: 'no spaces' });
		await until(() => failures.length === 1);
		expect(failures[0]!.message).toMatch(/letters, digits/);
		expect(a.share.active).toBe(false);
	});
});
