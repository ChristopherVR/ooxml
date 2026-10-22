// A formula names the shape it means by key, so a collision never retargets it; the winner of a
// colliding id is settled by the room (a persisted claim), so a later concurrent add cannot
// displace it; and the reader resolves formulas without rewriting the author's entry, so the
// author's undo still reverts the edit.
import JSZip from 'jszip';
import { afterEach, describe, expect, it } from 'vitest';
import { createVsdx, editVsdx } from '../index';
import { readSharedDrawing } from './index';
import {
	type DocPeer,
	converge,
	convergeStable,
	joinFrom,
	seedPeer,
} from './test-support/doc-peers';
import { createRoom, settle, until, wait } from './test-support/peers';

const pageName = (zip: JSZip): string =>
	Object.keys(zip.files).find((name) => /^visio\/pages\/page\d+\.xml$/.test(name))!;
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

/** The id of the shape whose PinX is `x`, read from the page part of `zip`. */
async function idAtPinX(zip: JSZip, x: number): Promise<string> {
	const xml = (await zip.file(pageName(zip))!.async('string')) as string;
	for (const m of xml.matchAll(/<Shape ID="(\d+)"[^>]*>[\s\S]*?<Cell N="PinX" V="([^"]*)"/g))
		if (Number(m[2]) === x) return m[1]!;
	throw new Error(`no shape at ${x}`);
}
/** Sets the LocPinY formula of shape `id` to `f`. */
async function setFormula(zip: JSZip, id: string, f: string): Promise<void> {
	const xml = (await zip.file(pageName(zip))!.async('string')) as string;
	const next = xml.replace(
		new RegExp(`(<Shape ID="${id}"[^>]*>[\\s\\S]*?<Cell N="LocPinY" V="[^"]*")(?: F="[^"]*")?`),
		`$1 F="${f}"`,
	);
	if (next === xml) throw new Error('formula not set');
	zip.file(pageName(zip), next);
}
const formulaOf = async (peer: { bytes: Uint8Array }, id: string) => {
	const zip = await JSZip.loadAsync(peer.bytes);
	const xml = (await zip.file(pageName(zip))!.async('string')) as string;
	const cell = new RegExp(`<Shape ID="${id}"[^>]*>[\\s\\S]*?(<Cell N="LocPinY"[^>]*>)`).exec(
		xml,
	)?.[1];
	return cell ? /\sF="([^"]*)"/.exec(cell)?.[1] : undefined;
};
const zipOf = (peer: { bytes: Uint8Array }) => JSZip.loadAsync(peer.bytes);
const idAt = async (peer: DocPeer, x: number): Promise<string> => idAtPinX(await zipOf(peer), x);
const ids = async (peer: DocPeer) =>
	[...((await peer.text('visio/pages/page1.xml')) ?? '').matchAll(/<Shape ID="(\d+)"/g)]
		.map((m) => m[1]!)
		.sort();

async function pair() {
	const start = (await editVsdx(await createVsdx(), [rect('1', 1), rect('2', 3)])).bytes;
	const a = await seedPeer(start, 'a');
	const b = await joinFrom(a, 'b');
	return { a, b };
}

describe('a formula written during a collision keeps its target', () => {
	it('points a formula of another shape at the author own shape, not the id keeper', async () => {
		const { a, b } = await pair();
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5));
		// b edits an older shape (another batch) to follow its own shape 9.
		await b.patch(async (zip) => setFormula(zip, '2', 'Sheet.9!Height*0.5'));
		await convergeStable(a, b);
		const mine = await idAtPinX(await zipOf(b), 6.5);
		expect(mine).not.toBe('9');
		for (const peer of [a, b]) {
			expect(await idAtPinX(await zipOf(peer), 6.5)).toBe(mine);
			expect(await formulaOf(peer, '2')).toBe(`Sheet.${mine}!Height*0.5`);
		}
		// The room still holds exactly what the author wrote, plus the shapes it names.
		const entry = [...a.shared.cells].find(([, v]) => JSON.stringify(v).includes('Sheet.9'))!;
		expect(JSON.stringify(entry[1])).toContain('"r":["9"');
		expect(readSharedDrawing(a.shared)!.fixups).toEqual([]);
	});

	it('keeps a formula that names the shape that kept the id on that shape', async () => {
		const { a, b } = await pair();
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5));
		await a.patch(async (zip) => setFormula(zip, '2', 'Sheet.9!Height*0.5'));
		await convergeStable(a, b);
		for (const peer of [a, b]) expect(await formulaOf(peer, '2')).toBe('Sheet.9!Height*0.5');
	});

	it('follows a shape that is renumbered in a later round, whatever the batch', async () => {
		const { a, b } = await pair();
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5), rect('10', 7.5));
		await b.patch(async (zip) => setFormula(zip, '10', 'Sheet.9!Height*0.5'));
		await b.patch(async (zip) => setFormula(zip, '2', 'Sheet.10!Height*0.5'));
		await convergeStable(a, b);
		const [nine, ten] = [await idAtPinX(await zipOf(b), 6.5), await idAtPinX(await zipOf(b), 7.5)];
		for (const peer of [a, b]) {
			expect(await formulaOf(peer, ten)).toBe(`Sheet.${nine}!Height*0.5`);
			expect(await formulaOf(peer, '2')).toBe(`Sheet.${ten}!Height*0.5`);
		}
	});
});

describe('the winner of a colliding id is settled by the room', () => {
	it('does not displace a persisted id for a concurrent add that sorts first', async () => {
		const { a, b } = await pair();
		const c = await joinFrom(a, '0'); // keys `0.n` sort before `a.n`, and c has not seen the collision
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5));
		await convergeStable(a, b);
		const winner = await idAtPinX(await zipOf(a), 1.5);
		const loser = await idAtPinX(await zipOf(a), 6.5);
		expect(winner).toBe('9');
		await c.edit(rect('9', 8.5)); // numbered alike, unaware of what a and b settled
		await convergeStable(a, b, c);
		for (const peer of [a, b, c]) {
			expect(await idAtPinX(await zipOf(peer), 1.5)).toBe(winner);
			expect(await idAtPinX(await zipOf(peer), 6.5)).toBe(loser);
			expect(await idAt(peer, 8.5)).not.toBe(winner);
			expect(new Set(await ids(peer)).size).toBe(5);
		}
		expect(JSON.stringify([...a.shared.renum])).toBe(JSON.stringify([...c.shared.renum]));
	});

	it('gives peers with the same room the same ids, persisted or not', async () => {
		const { a, b } = await pair();
		const c = await joinFrom(a, '0');
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5));
		await c.edit(rect('9', 8.5));
		// Nobody has persisted anything yet: every key order decides alike.
		await converge(a, b, c);
		const before = await Promise.all([a, b, c].map(async (p) => (await ids(p)).join()));
		expect(new Set(before).size).toBe(1);
		const at = async (p: DocPeer) => [await idAt(p, 1.5), await idAt(p, 6.5), await idAt(p, 8.5)];
		const expected = await at(a);
		expect(expected[2]).toBe('9'); // the lowest key wins when nothing is persisted
		await convergeStable(a, b, c);
		for (const peer of [a, b, c]) expect(await at(peer)).toEqual(expected);
	});
});

describe('undo of a formula edit that a collision re-targeted', () => {
	const rooms: ReturnType<typeof createRoom>[] = [];
	afterEach(() => {
		for (const room of rooms.splice(0)) room.destroy();
	});

	it('reverts the author edit and leaves the room entry the author wrote', async () => {
		const room = createRoom();
		rooms.push(room);
		const seed = (await editVsdx(await createVsdx(), [rect('1', 1), rect('2', 3)])).bytes;
		const a = room.join('Ada', seed);
		await wait();
		const b = room.join('Bob', await createVsdx(), { seed: false });
		await until(async () => (await b.side.model()).pages[0]!.shapes.length === 2, 'join');
		await settle(a, b);
		await Promise.all([
			a.side.edit(rect('9', 8.5)),
			(async () => {
				await b.side.edit(rect('9', 6.5));
				await b.side.patch(async (zip) =>
					setFormula(zip, '2', `Sheet.${await idAtPinX(zip, 6.5)}!Height*0.5`),
				);
			})(),
		]);
		await until(async () => (await a.side.model()).pages[0]!.shapes.length === 4, 'all shapes');
		await settle(a, b);
		const shapeAt = async (p: typeof a, x: number) => idAtPinX(await zipOf(p.side), x);
		const bobs = await shapeAt(b, 6.5);
		expect(await shapeAt(a, 6.5)).toBe(bobs);
		for (const p of [a, b]) expect(await formulaOf(p.side, '2')).toBe(`Sheet.${bobs}!Height*0.5`);
		// Bob undoes the formula edit: it reverts to the original on both peers, whatever the collision did.
		expect(await b.binding.undo()).toBe(true);
		await settle(a, b);
		for (const p of [a, b]) expect(await formulaOf(p.side, '2')).toBe('Height*0.5');
		expect(await b.binding.redo()).toBe(true);
		await settle(a, b);
		for (const p of [a, b]) expect(await formulaOf(p.side, '2')).toBe(`Sheet.${bobs}!Height*0.5`);
		expect(a.errors).toEqual([]);
		expect(b.errors).toEqual([]);
	});
});
