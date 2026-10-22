import { afterEach, describe, expect, it, vi } from 'vitest';
import {
	createCollabSession,
	createMemoryHub,
	transportProvider,
	type CollaborationState,
} from 'ooxml-core/collab';
import { createBackstage } from './backstage';
import { ViewerController } from './controller';
import { createVsdx, editVsdx, parseVsdx, type VisioEdit } from 'ooxml-core/visio';
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

const rect = (id: string, x: number): VisioEdit => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId: id,
	x,
	y: 2,
	width: 1,
	height: 1,
});
const move = (id: string, x: number): VisioEdit => ({
	type: 'move-shape',
	pageId: '0',
	shapeId: id,
	x,
	y: 2,
});
const pins = (w: { controller: ViewerController }) =>
	(w.controller.state.document?.pages[0]?.shapes ?? []).map(
		(shape) => shape.rotation?.pinX ?? shape.transform[4],
	);

function window(name: string) {
	registerViewerControls();
	const editor: CancellableEditor = async (bytes, commands) => {
		const result = await editVsdx(bytes, commands);
		return { ...result, document: await parseVsdx(result.bytes) };
	};
	const controller = new ViewerController(parseVsdx, () => {}, editor);
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
	return { controller, share, root, people };
}

describe('File > Share', () => {
	it('rejects session names the core would refuse', async () => {
		const a = window('Ada');
		await a.share.start('no spaces');
		expect(a.share.active).toBe(false);
		expect(a.root.querySelector('.share-status')!.textContent).toMatch(/letters, digits/);
	});

	it('merges concurrent edits of different shapes and undoes only local edits', async () => {
		const room = `test-${Math.random().toString(36).slice(2, 8)}`;
		const a = window('Ada');
		await a.controller.load(await createVsdx());
		await a.controller.applyEdits([rect('1', 1), rect('2', 3)]);
		await a.share.start(room);
		expect(a.share.active).toBe(true);
		expect(a.root.querySelector<HTMLElement>('[data-share="stop"]')!.hidden).toBe(false);

		const b = window('Grace');
		await b.share.start(room);
		// The joiner had nothing open: it loads the room's drawing.
		await until(() => pins(b).length === 2);
		await until(() => b.people().includes('Ada') && a.people().includes('Grace'));

		// Both windows edit different shapes at the same moment: both edits survive.
		await Promise.all([
			a.controller.applyEdits([move('1', 5)]),
			b.controller.applyEdits([move('2', 7)]),
		]);
		await until(() => pins(a).join() === '5,7' && pins(b).join() === '5,7');
		await until(() => !a.controller.state.edit.busy && !b.controller.state.edit.busy);

		// Undo reverts only this window's own edit.
		await new Promise((r) => setTimeout(r, 300));
		await until(() => b.controller.state.edit.canUndo);
		await b.controller.undo();
		await until(() => pins(a).join() === '5,3' && pins(b).join() === '5,3');
		b.share.stop();
		expect(b.people()).toEqual([]);
		// Not sharing any more: the package history is back.
		expect(b.controller.state.edit.canUndo).toBeDefined();
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

		const source = await createVsdx();
		await a.controller.load(await editVsdx(source, [rect('1', 2)]).then((r) => r.bytes));
		expect(a.controller.state.zoom).toBe(2);
		await until(() => pins(b).length === 1);
		// A later local edit still reaches the peer.
		await a.controller.applyEdits([move('1', 4)]);
		await until(() => pins(b).join() === '4');
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
		await a.controller.load(await createVsdx());
		const options = { session: first };
		a.share.want(options);
		await until(() => a.share.active);
		expect(a.share.wanted).toBe(options);
		b.share.want({ session: second });
		await until(() => b.controller.state.document !== null);
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
		await a.controller.load(await createVsdx());
		a.share.want({ session, intent: 'join' });
		await until(() => a.share.active);
		await new Promise((r) => setTimeout(r, 30));
		expect(session.doc.getMap('visio:pages').size).toBe(0);
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
