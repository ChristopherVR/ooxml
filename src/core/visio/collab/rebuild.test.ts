// Rebuild cost: bursts of remote changes share one rebuild, a change that leaves the drawing as it
// is rebuilds nothing, unchanged pages are not serialised again, and a room of another schema
// version is left alone.
import { afterEach, describe, expect, it } from 'vitest';
import { createVsdx, editVsdx } from '../index';
import { composeDrawing } from './compose';
import { SchemaVersionError, readSharedDrawing } from './index';
import { createRoom, settle, until, wait } from './test-support/peers';
import { seedPeer } from './test-support/doc-peers';

const rect = (id: string, x: number) =>
	({
		type: 'create-rectangle',
		pageId: '0',
		shapeId: id,
		x,
		y: 2,
		width: 1,
		height: 1,
	}) as const;

const rooms: ReturnType<typeof createRoom>[] = [];
afterEach(() => {
	for (const room of rooms.splice(0)) room.destroy();
});
const newRoom = () => {
	const room = createRoom();
	rooms.push(room);
	return room;
};

async function pair() {
	const room = newRoom();
	const queue: (() => void)[] = [];
	const a = room.join('Ada', (await editVsdx(await createVsdx(), [rect('1', 1)])).bytes);
	await wait();
	const b = room.join('Bob', await createVsdx(), { schedule: (run) => void queue.push(run) });
	await until(async () => (await b.side.model()).pages[0]!.shapes.length === 1, 'join');
	await settle(a, b);
	return { a, b, queue };
}

describe('coalesced rebuilds', () => {
	it('applies a burst of remote edits with one rebuild', async () => {
		const { a, b, queue } = await pair();
		const before = b.side.applied;
		for (let i = 0; i < 8; i++) await a.side.edit(rect(String(10 + i), 2 + i));
		await wait(150);
		expect(queue).toHaveLength(1);
		expect(b.side.applied).toBe(before);
		queue.shift()!();
		await b.binding.idle();
		expect(b.side.applied).toBe(before + 1);
		expect((await b.side.model()).pages[0]!.shapes).toHaveLength(9);
		expect(b.errors).toEqual([]);
	});

	it('does not rebuild for a change that leaves the drawing as it is', async () => {
		const { a, b, queue } = await pair();
		const before = b.side.applied;
		a.session.doc.getMap('visio:meta').set('noise', 1);
		await wait(100);
		expect(queue).toHaveLength(1);
		queue.shift()!();
		await b.binding.idle();
		expect(b.side.applied).toBe(before);
	});

	it('rebuilds again after the person edited locally', async () => {
		const { a, b, queue } = await pair();
		const before = b.side.applied;
		await a.side.edit(rect('20', 5));
		await wait(100);
		queue.shift()!();
		await b.binding.idle();
		expect(b.side.applied).toBe(before + 1);
		expect((await b.side.model()).pages[0]!.shapes).toHaveLength(2);
	});
});

describe('page cache', () => {
	it('serialises only the pages that changed', async () => {
		let start = (await editVsdx(await createVsdx(), [rect('1', 1)])).bytes;
		start = (
			await editVsdx(start, [{ type: 'insert-page', pageId: '5', afterPageId: '0', name: 'Two' }])
		).bytes;
		start = (await editVsdx(start, [{ ...rect('2', 3), pageId: '5' }])).bytes;
		const peer = await seedPeer(start, 'a');
		const build = async () => {
			const composed = await composeDrawing(readSharedDrawing(peer.shared)!, peer.bytes, peer.keys);
			composed.commit();
			return new Map([...peer.keys.cache.pageParts].map(([k, v]) => [k, v.bytes]));
		};
		const first = await build();
		expect(first.size).toBe(2);
		await peer.edit({ type: 'move-shape', pageId: '5', shapeId: '2', x: 6, y: 2 });
		const second = await build();
		const [p0, p5] = ['s0', 's5'];
		expect(second.get(p0)).toBe(first.get(p0));
		expect(second.get(p5)).not.toBe(first.get(p5));
	});
});

describe('schema version in a session', () => {
	it('leaves a room of another version alone and reports it once', async () => {
		const room = newRoom();
		const a = room.join('Ada', await createVsdx());
		await wait();
		a.session.doc.getMap('visio:meta').set('v', 1);
		const b = room.join('Bob', (await editVsdx(await createVsdx(), [rect('1', 1)])).bytes);
		await until(() => b.errors.length > 0, 'version error');
		expect(b.errors[0]).toBeInstanceOf(SchemaVersionError);
		const cells = b.session.doc.getMap('visio:cells').size;
		await b.side.edit(rect('2', 2));
		await wait(100);
		expect(b.errors).toHaveLength(1);
		expect(b.session.doc.getMap('visio:cells').size).toBe(cells);
		expect(b.session.doc.getMap('visio:meta').get('v')).toBe(1);
	});
});
