// The binding on real collab sessions over the in-memory transport, with a package-backed editor.
import { afterEach, describe, expect, it } from 'vitest';
import { createMemoryHub } from '../../collab/memory-transport';
import { type CollabSession, createCollabSession } from '../../collab/session';
import { transportProvider } from '../../collab/transport-provider';
import { type VisioEdit, createVsdx, editVsdx, parseVsdx } from '../index';
import { type DrawingBinding, type DrawingEditor, bindDrawingSession } from './bind';

const open: CollabSession<object>[] = [];
const bindings: DrawingBinding[] = [];
afterEach(() => {
	for (const binding of bindings.splice(0)) binding.dispose();
	for (const session of open.splice(0)) session.destroy();
});

const wait = (ms = 30) => new Promise((resolve) => setTimeout(resolve, ms));
const until = async (check: () => boolean | Promise<boolean>) => {
	for (let i = 0; i < 200 && !(await check()); i++) await wait(10);
	expect(await check()).toBe(true);
};

/** An editor over package bytes: edits go through the real VSDX writer. */
function makeEditor(initial: Uint8Array | undefined) {
	let bytes = initial;
	let version = 0;
	let working = false;
	const local = new Set<() => void>();
	const idle = new Set<() => void>();
	const editor: DrawingEditor = {
		exportBytes: () => bytes,
		applyRemote: async (next) => {
			bytes = next;
			version++;
		},
		onLocalChange: (listener) => {
			local.add(listener);
			return () => local.delete(listener);
		},
		onIdle: (listener) => {
			idle.add(listener);
			return () => idle.delete(listener);
		},
		busy: () => working,
		version: () => version,
	};
	return {
		editor,
		async edit(...edits: VisioEdit[]) {
			working = true;
			try {
				bytes = (await editVsdx(bytes!, edits)).bytes;
			} finally {
				working = false;
			}
			version++;
			for (const listener of [...local]) listener();
			for (const listener of [...idle]) listener();
		},
		async model() {
			return parseVsdx(bytes!);
		},
	};
}

const rect = (id: string, x: number, pageId = '0'): VisioEdit => ({
	type: 'create-rectangle',
	pageId,
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
const pin = (s: { rotation?: { pinX: number } | undefined; transform: readonly number[] }) =>
	s.rotation?.pinX ?? s.transform[4]!;

function join(hub: ReturnType<typeof createMemoryHub>, name: string, initial?: Uint8Array) {
	const session = createCollabSession({
		roomId: 'drawing',
		provider: transportProvider({ transport: hub.createTransport('drawing') }),
		user: { name },
		heartbeatMs: 0,
		syncGraceMs: 0,
	});
	open.push(session);
	const side = makeEditor(initial);
	const errors: unknown[] = [];
	const binding = bindDrawingSession(session, side.editor, { onError: (e) => errors.push(e) });
	bindings.push(binding);
	return { session, binding, errors, side };
}
type Peer = ReturnType<typeof join>;

async function pair() {
	const hub = createMemoryHub();
	const a = join(hub, 'Ada', await createVsdx());
	await a.side.edit(rect('1', 1), rect('2', 3));
	await wait();
	const b = join(hub, 'Bob', await createVsdx({ width: 4 }));
	await until(async () => (await b.side.model()).pages[0]!.shapes.length === 2);
	return { hub, a, b };
}
const ids = async (peer: Peer) => (await peer.side.model()).pages[0]!.shapes.map((s) => s.id);
const pins = async (peer: Peer) => (await peer.side.model()).pages[0]!.shapes.map(pin);
const pageNames = async (peer: Peer) => (await peer.side.model()).pages.map((p) => p.name).join();

describe('bindDrawingSession', () => {
	it('seeds an empty room and a joiner adopts it, replacing its own drawing', async () => {
		const { a, b } = await pair();
		expect(await ids(b)).toEqual(['1', '2']);
		expect(a.errors).toEqual([]);
		expect(b.errors).toEqual([]);
		expect((await b.side.model()).pages[0]!.width).toBe((await a.side.model()).pages[0]!.width);
	});

	it('keeps concurrent edits to different shapes', async () => {
		const { a, b } = await pair();
		await Promise.all([a.side.edit(move('1', 5)), b.side.edit(move('2', 7))]);
		await until(async () => [await pins(a), await pins(b)].every((p) => p.join() === '5,7'));
	});

	it('resolves two edits of one shape to the same value on both peers', async () => {
		const { a, b } = await pair();
		await Promise.all([a.side.edit(move('1', 5)), b.side.edit(move('1', 8))]);
		await wait(150);
		await a.binding.idle();
		await b.binding.idle();
		const [pa, pb] = [await pins(a), await pins(b)];
		expect(pa).toEqual(pb);
		expect([5, 8]).toContain(pa[0]);
		expect(pa[1]).toBe(3);
	});

	it('lets a deletion win over a concurrent edit of the same shape', async () => {
		const { a, b } = await pair();
		await Promise.all([
			a.side.edit({ type: 'delete-shape', pageId: '0', shapeId: '1' }),
			b.side.edit(move('1', 6)),
		]);
		await until(async () => (await ids(a)).join() === '2' && (await ids(b)).join() === '2');
		await wait(100);
		expect(await ids(a)).toEqual(['2']);
		expect(await ids(b)).toEqual(['2']);
	});

	it('gives shapes with the same id added concurrently distinct ids on every peer', async () => {
		const { a, b } = await pair();
		await Promise.all([a.side.edit(rect('9', 1.5)), b.side.edit(rect('9', 6.5))]);
		await until(async () => (await ids(a)).length === 4 && (await ids(b)).length === 4);
		const [ia, ib] = [await ids(a), await ids(b)];
		expect(new Set(ia).size).toBe(4);
		expect([...ia].sort()).toEqual([...ib].sort());
		expect([...(await pins(a))].sort()).toEqual([...(await pins(b))].sort());
		expect((await pins(a)).filter((x) => x === 1.5 || x === 6.5)).toHaveLength(2);
	});

	it('syncs pages: add, rename and delete', async () => {
		const { a, b } = await pair();
		await a.side.edit({ type: 'insert-page', pageId: '5', afterPageId: '0', name: 'Second' });
		await until(async () => (await pageNames(b)) === 'Page-1,Second');
		await b.side.edit({ type: 'rename-page', pageId: '5', name: 'Renamed' });
		await until(async () => (await pageNames(a)) === 'Page-1,Renamed');
		await a.side.edit({ type: 'delete-page', pageId: '5' });
		await until(async () => (await b.side.model()).pages.length === 1);
	});

	it('undo reverts only the local edit, never a collaborator', async () => {
		const { a, b } = await pair();
		await a.side.edit(move('1', 5));
		await until(async () => (await pins(b))[0] === 5);
		await b.side.edit(move('2', 7));
		await until(async () => (await pins(a))[1] === 7);
		expect(b.binding.canUndo()).toBe(true);
		expect(await b.binding.undo()).toBe(true);
		await until(async () => (await pins(a))[1] === 3);
		expect(await pins(a)).toEqual([5, 3]);
		expect(await pins(b)).toEqual([5, 3]);
		expect(await b.binding.redo()).toBe(true);
		await until(async () => (await pins(a))[1] === 7);
		expect(await a.binding.undo()).toBe(true);
		await until(async () => (await pins(b))[0] !== 5);
		expect((await pins(b))[1]).toBe(7);
	});

	it('does not seed an empty room when told to join', async () => {
		const hub = createMemoryHub();
		const session = createCollabSession({
			roomId: 'drawing',
			provider: transportProvider({ transport: hub.createTransport('drawing') }),
			user: { name: 'Vera' },
			heartbeatMs: 0,
			syncGraceMs: 0,
		});
		open.push(session);
		const side = makeEditor(await createVsdx());
		bindings.push(bindDrawingSession(session, side.editor, { seed: false }));
		await wait(60);
		expect(session.doc.getMap('visio:pages').size).toBe(0);
	});
});
