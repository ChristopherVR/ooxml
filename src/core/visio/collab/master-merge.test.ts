// Edits inside one master merge per cell: the shapes of a master's content are shape and cell
// entries under the master's scope, so two peers editing the same master both keep their change.
import JSZip from 'jszip';
import { afterEach, describe, expect, it } from 'vitest';
import { createVsdx, editVsdx } from '../index';
import { type DocPeer, converge, joinFrom, seedPeer } from './test-support/doc-peers';
import { createRoom, settle, until, wait } from './test-support/peers';
import {
	addMaster,
	masterCell,
	masterPartPaths,
	masterShape,
	removeMaster,
	setMasterCell,
} from './test-support/masters';

async function pair() {
	const start = (
		await editVsdx(await createVsdx(), [
			{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 1, y: 2, width: 1, height: 1 },
		])
	).bytes;
	const a = await seedPeer(start, 'a');
	await a.patch((zip) =>
		addMaster(zip, {
			id: '2',
			unique: '{M}',
			shapes: masterShape('1', 1) + masterShape('2', 3),
		}),
	);
	const b = await joinFrom(a, 'b');
	return { a, b };
}
const part = async (peer: DocPeer) => {
	const [path] = masterPartPaths(await peer.names());
	return path === undefined ? undefined : peer.text(path);
};
const cell = async (peer: DocPeer, shape: string, name: string) =>
	masterCell(await part(peer), shape, name);

describe('edits inside one master', () => {
	it('shares the shapes of a master as entries under the master scope', async () => {
		const { a } = await pair();
		const keys = [...a.shared.shapes.keys()].filter((k) => k.startsWith('m~'));
		expect(keys).toHaveLength(2);
		expect(a.shared.masters.size).toBe(1);
		expect(a.shared.masters.get([...a.shared.masters.keys()][0]!)).not.toHaveProperty('shapes');
	});

	it('keeps edits of two different shapes of the same master', async () => {
		const { a, b } = await pair();
		await a.patch((zip) => setMasterCell(zip, '1', 'PinX', 5));
		await b.patch((zip) => setMasterCell(zip, '2', 'PinX', 7));
		await converge(a, b);
		for (const peer of [a, b]) {
			expect(await cell(peer, '1', 'PinX')).toBe('5');
			expect(await cell(peer, '2', 'PinX')).toBe('7');
		}
		expect(await part(a)).toBe(await part(b));
	});

	it('keeps edits of two different cells of the same master shape', async () => {
		const { a, b } = await pair();
		await a.patch((zip) => setMasterCell(zip, '1', 'PinX', 5));
		await b.patch((zip) => setMasterCell(zip, '1', 'PinY', 9));
		await converge(a, b);
		for (const peer of [a, b]) {
			expect(await cell(peer, '1', 'PinX')).toBe('5');
			expect(await cell(peer, '1', 'PinY')).toBe('9');
		}
	});

	it('resolves one master cell edited by both to the same value everywhere', async () => {
		const { a, b } = await pair();
		await a.patch((zip) => setMasterCell(zip, '1', 'PinX', 5));
		await b.patch((zip) => setMasterCell(zip, '1', 'PinX', 8));
		await converge(a, b);
		const [xa, xb] = [await cell(a, '1', 'PinX'), await cell(b, '1', 'PinX')];
		expect(xa).toBe(xb);
		expect(['5', '8']).toContain(xa);
		expect(await cell(a, '2', 'PinX')).toBe('3');
	});

	it('writes only the cell that changed', async () => {
		const { a } = await pair();
		const before = new Map([...a.shared.cells].map(([k, v]) => [k, JSON.stringify(v)]));
		const masters = JSON.stringify([...a.shared.masters]);
		await a.patch((zip) => setMasterCell(zip, '1', 'PinX', 5));
		const changed = [...a.shared.cells].filter(([k, v]) => before.get(k) !== JSON.stringify(v));
		expect(changed).toHaveLength(1);
		expect(changed[0]![0]).toMatch(/^m~.*\/c\/PinX$/);
		expect(JSON.stringify([...a.shared.masters])).toBe(masters);
	});

	it('lets a deleted master win over a concurrent edit inside it', async () => {
		const { a, b } = await pair();
		await a.patch((zip) => removeMaster(zip, '2'));
		await b.patch((zip) => setMasterCell(zip, '1', 'PinX', 6));
		await converge(a, b);
		for (const peer of [a, b]) {
			expect(masterPartPaths(await peer.names())).toEqual([]);
			expect(await peer.text('visio/masters/masters.xml')).not.toContain('<Master ');
		}
		expect(a.shared.masters.size).toBe(0);
	});

	it('adds and removes a shape inside a master on one peer while another edits a sibling', async () => {
		const { a, b } = await pair();
		await a.patch(async (zip) => {
			const [path] = masterPartPaths(Object.keys(zip.files));
			const xml = (await zip.file(path!)!.async('string')) as string;
			zip.file(path!, xml.replace('</Shapes>', `${masterShape('3', 4)}</Shapes>`));
		});
		await b.patch((zip) => setMasterCell(zip, '2', 'PinY', 4));
		await converge(a, b);
		for (const peer of [a, b]) {
			expect(await cell(peer, '3', 'PinX')).toBe('4');
			expect(await cell(peer, '2', 'PinY')).toBe('4');
		}
	});
});

describe('edits inside one master in live sessions', () => {
	const rooms: ReturnType<typeof createRoom>[] = [];
	afterEach(() => {
		for (const room of rooms.splice(0)) room.destroy();
	});

	it('keeps concurrent edits of two shapes of a master on both peers', async () => {
		const room = createRoom();
		rooms.push(room);
		const a = room.join('Ada', await createVsdx());
		await a.side.patch((zip) =>
			addMaster(zip, { id: '2', unique: '{M}', shapes: masterShape('1', 1) + masterShape('2', 3) }),
		);
		await wait();
		const b = room.join('Bob', await createVsdx());
		const text = async (p: typeof a) => {
			const zip = await JSZip.loadAsync(p.side.bytes);
			const [path] = masterPartPaths(Object.keys(zip.files));
			return path === undefined ? undefined : zip.file(path)!.async('string');
		};
		await until(async () => (await text(b)) !== undefined, 'master joined');
		await settle(a, b);
		await Promise.all([
			a.side.patch((zip) => setMasterCell(zip, '1', 'PinX', 5)),
			b.side.patch((zip) => setMasterCell(zip, '2', 'PinX', 7)),
		]);
		await until(
			async () =>
				masterCell(await text(a), '2', 'PinX') === '7' &&
				masterCell(await text(b), '1', 'PinX') === '5',
			'edits',
		);
		await settle(a, b);
		for (const p of [a, b]) {
			expect(masterCell(await text(p), '1', 'PinX')).toBe('5');
			expect(masterCell(await text(p), '2', 'PinX')).toBe('7');
			expect(p.errors).toEqual([]);
		}
	});
});
