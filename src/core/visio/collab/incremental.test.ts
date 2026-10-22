// What the binding tells the editor about a remote change: just the pages whose content differs
// while the page list, masters and shared parts stay as they are, the whole package otherwise.
import { afterEach, describe, expect, it } from 'vitest';
import { createVsdx, editVsdx } from '../index';
import { addMaster, masterShape } from './test-support/masters';
import { createRoom, settle, until, wait } from './test-support/peers';

const rect = (id: string, x: number, pageId = '0') =>
	({ type: 'create-rectangle', pageId, shapeId: id, x, y: 2, width: 1, height: 1 }) as const;
const move = (id: string, x: number, pageId = '0') =>
	({ type: 'move-shape', pageId, shapeId: id, x, y: 2 }) as const;

const rooms: ReturnType<typeof createRoom>[] = [];
afterEach(() => {
	for (const room of rooms.splice(0)) room.destroy();
});

async function pair() {
	const room = createRoom();
	rooms.push(room);
	let start = (await editVsdx(await createVsdx(), [rect('1', 1), rect('2', 3)])).bytes;
	start = (
		await editVsdx(start, [{ type: 'insert-page', pageId: '5', afterPageId: '0', name: 'Two' }])
	).bytes;
	start = (await editVsdx(start, [rect('1', 2, '5')])).bytes;
	const a = room.join('Ada', start);
	await wait();
	// A joiner never seeds: a blank drawing racing the first peer's seed would also change page 0.
	const b = room.join('Bob', await createVsdx(), { seed: false });
	await until(async () => (await b.side.model()).pages.length === 2, 'join');
	await settle(a, b);
	return { a, b, room };
}
type Peer = Awaited<ReturnType<typeof pair>>['a'];
const last = (peer: Peer) => peer.side.changes[peer.side.changes.length - 1];
const pinX = async (peer: Peer, pageId: string, id: string) =>
	(await peer.side.model()).pages.find((p) => p.id === pageId)!.shapes.find((s) => s.id === id)!
		.transform[4];

describe('the change handed to the editor', () => {
	it('is the whole package for the first drawing a joiner adopts', async () => {
		const { b } = await pair();
		expect(b.side.changes[0]).toBeUndefined();
		expect(b.side.changes.length).toBeGreaterThan(0);
	});

	it('names only the page whose cell a collaborator edited', async () => {
		const { a, b } = await pair();
		const before = b.side.applied;
		await a.side.edit(move('1', 6, '5'));
		await until(async () => (await pinX(b, '5', '1')) === 5.5, 'edit');
		await settle(a, b);
		expect(b.side.applied).toBe(before + 1);
		expect(last(b)).toEqual({ pages: ['5'] });
		expect(await pinX(b, '0', '1')).toBe(0.5);
		expect(b.errors).toEqual([]);
	});

	it('names every page a burst of edits touched', async () => {
		const { a, b } = await pair();
		await a.side.edit(move('1', 6, '5'));
		await a.side.edit(move('2', 7, '0'));
		await until(async () => (await pinX(b, '0', '2')) === 6.5 && (await pinX(b, '5', '1')) === 5.5);
		await settle(a, b);
		const named = new Set(b.side.changes.flatMap((c) => c?.pages ?? []));
		expect(named).toEqual(new Set(['0', '5']));
	});

	it('is the whole package when a master changes', async () => {
		const { a, b } = await pair();
		await a.side.patch((zip) =>
			addMaster(zip, { id: '2', unique: '{M}', shapes: masterShape('1', 1) }),
		);
		await until(async () => (await b.side.model()).masters?.length === 1, 'master');
		await settle(a, b);
		expect(last(b)).toBeUndefined();
	});

	it('is the whole package when the page list changes', async () => {
		const { a, b } = await pair();
		await a.side.edit({ type: 'rename-page', pageId: '5', name: 'Renamed' });
		await until(async () => (await b.side.model()).pages[1]!.name === 'Renamed', 'rename');
		await settle(a, b);
		expect(last(b)).toBeUndefined();
		await a.side.edit({ type: 'delete-page', pageId: '5' });
		await until(async () => (await b.side.model()).pages.length === 1, 'delete');
		await settle(a, b);
		expect(last(b)).toBeUndefined();
	});

	it('is the whole package when a shared part changes', async () => {
		const { a, b } = await pair();
		await a.side.patch(async (zip) => {
			const path = 'visio/document.xml';
			zip.file(path, `${await zip.file(path)!.async('string')}<!-- changed -->`);
		});
		await until(
			async () => ((await b.side.text('visio/document.xml')) ?? '').includes('<!-- changed -->'),
			'part',
		);
		await settle(a, b);
		expect(last(b)).toBeUndefined();
	});

	it('is the whole package after the person edited here, then names pages again', async () => {
		const { a, b } = await pair();
		await b.side.edit(move('2', 4, '0'));
		await until(async () => (await pinX(a, '0', '2')) === 3.5, 'local edit published');
		await settle(a, b);
		await a.side.edit(move('1', 6, '5'));
		await until(async () => (await pinX(b, '5', '1')) === 5.5, 'remote edit');
		await settle(a, b);
		expect(last(b)).toBeUndefined();
		await a.side.edit(move('1', 7, '5'));
		await until(async () => (await pinX(b, '5', '1')) === 6.5, 'next remote edit');
		await settle(a, b);
		expect(last(b)).toEqual({ pages: ['5'] });
	});
});
