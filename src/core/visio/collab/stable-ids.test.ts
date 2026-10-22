// Renumbered ids are persisted by a fix-up, so they stay put: a later concurrent add with a higher
// id does not shift them, formulas written afterwards keep their meaning, two peers writing the
// fix-up at once converge, and neither user's undo reverts it.
import type JSZip from 'jszip';
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
import { addMaster, masterShape } from './test-support/masters';
import { createRoom, settle, until, wait } from './test-support/peers';

const rect = (id: string, x: number, pageId = '0') =>
	({
		type: 'create-rectangle',
		pageId,
		shapeId: id,
		x,
		y: 2,
		width: 1,
		height: 1,
	}) as const;
const PAGE = 'visio/pages/page1.xml';

async function pair() {
	const start = (await editVsdx(await createVsdx(), [rect('1', 1), rect('2', 3)])).bytes;
	const a = await seedPeer(start, 'a');
	const b = await joinFrom(a, 'b');
	return { a, b };
}
const ids = async (peer: DocPeer) =>
	[...((await peer.text(PAGE)) ?? '').matchAll(/<Shape ID="(\d+)"/g)].map((m) => m[1]!);
/** The id of the shape created at x (inches) on the page. */
async function idAt(peer: DocPeer, x: number): Promise<string> {
	const model = await peer.model();
	return model.pages[0]!.shapes.find((s) => Math.abs(s.transform[4]! - (x - 0.5)) < 1e-6)!.id;
}
const formula = async (peer: DocPeer, id: string, name: string) =>
	new RegExp(`<Shape ID="${id}"[^>]*>[\\s\\S]*?<Cell N="${name}"[^>]*F="([^"]*)"`).exec(
		(await peer.text(PAGE)) ?? '',
	)?.[1];
const setFormula = (id: string, name: string, f: string) => async (zip: JSZip) => {
	const xml = (await zip.file(PAGE)!.async('string')) as string;
	zip.file(
		PAGE,
		xml.replace(
			new RegExp(`(<Shape ID="${id}"[^>]*>[\\s\\S]*?<Cell N="${name}" V="[^"]*" F=")[^"]*(")`),
			`$1${f}$2`,
		),
	);
};

describe('persisted renumbering of shapes', () => {
	it('converges two concurrent adds of one id to the same stable ids', async () => {
		const { a, b } = await pair();
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5));
		await convergeStable(a, b);
		expect([...(await ids(a))].sort()).toEqual([...(await ids(b))].sort());
		expect(new Set(await ids(a)).size).toBe(4);
		expect(a.shared.renum.size).toBe(2);
		expect(JSON.stringify([...a.shared.renum])).toBe(JSON.stringify([...b.shared.renum]));
		// Nothing left to fix: the ids come from the room, not from the read.
		expect(readSharedDrawing(a.shared)!.fixups).toEqual([]);
		expect(a.fixup()).toBe(0);
		expect(await idAt(a, 1.5)).toBe('9'); // the lowest key (a's) keeps the id
		expect(await idAt(b, 6.5)).toBe(await idAt(a, 6.5));
	});

	it('does not shift a renumbered id when a higher id is added later, at the same time', async () => {
		const { a, b } = await pair();
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5));
		await convergeStable(a, b);
		const moved = await idAt(a, 6.5);
		await a.edit(rect('30', 8.5));
		await b.edit(rect('40', 9.5));
		await convergeStable(a, b);
		for (const peer of [a, b]) {
			expect(await idAt(peer, 6.5)).toBe(moved);
			expect(await idAt(peer, 1.5)).toBe('9');
		}
	});

	it('keeps a persisted id against a newcomer that numbered its shape alike', async () => {
		const { a, b } = await pair();
		const c = await joinFrom(a, 'c'); // stays behind while the others settle the collision
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5));
		await convergeStable(a, b);
		const moved = await idAt(a, 6.5);
		await c.edit(rect(moved, 8.5)); // c never saw the persisted id and numbers its shape alike
		await convergeStable(a, b, c);
		for (const peer of [a, b, c]) {
			expect(await idAt(peer, 6.5)).toBe(moved);
			expect(await idAt(peer, 1.5)).toBe('9');
			expect(await idAt(peer, 8.5)).not.toBe(moved);
			expect(new Set(await ids(peer)).size).toBe(5);
		}
	});

	it('rewrites the formulas of the batch once and keeps later references as written', async () => {
		const { a, b } = await pair();
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5), rect('10', 7.5));
		await b.patch(setFormula('10', 'LocPinY', 'Sheet.9!Height*0.5'));
		await convergeStable(a, b);
		const loser = await idAt(b, 6.5);
		expect(loser).not.toBe('9');
		for (const peer of [a, b])
			expect(await formula(peer, await idAt(peer, 7.5), 'LocPinY')).toBe(
				`Sheet.${loser}!Height*0.5`,
			);
		// The author now points a formula at the shape that kept id 9; it must stay that way.
		const mate = await idAt(b, 7.5);
		await b.patch(setFormula(mate, 'LocPinY', 'Sheet.9!Height*0.25'));
		await convergeStable(a, b);
		for (const peer of [a, b])
			expect(await formula(peer, await idAt(peer, 7.5), 'LocPinY')).toBe('Sheet.9!Height*0.25');
	});

	it('writes the fix-up from both peers at once without disagreement', async () => {
		const { a, b } = await pair();
		await a.edit(rect('9', 1.5));
		await b.edit(rect('9', 6.5));
		await converge(a, b);
		expect(a.fixup()).toBeGreaterThan(0);
		expect(b.fixup()).toBeGreaterThan(0);
		await converge(a, b);
		expect(JSON.stringify([...a.shared.renum])).toBe(JSON.stringify([...b.shared.renum]));
		expect(await a.text(PAGE)).toBe(await b.text(PAGE));
		expect(a.fixup() + b.fixup()).toBe(0);
	});
});

describe('persisted renumbering of pages and masters', () => {
	it('keeps a renumbered page id when a higher page id is added later', async () => {
		const { a, b } = await pair();
		await a.edit({ type: 'insert-page', pageId: '5', afterPageId: '0', name: 'A' });
		await b.edit({ type: 'insert-page', pageId: '5', afterPageId: '0', name: 'B' });
		await convergeStable(a, b);
		const pagesOf = async (peer: DocPeer) =>
			(await peer.model()).pages.map((p) => `${p.id}:${p.name}`).sort();
		const before = await pagesOf(a);
		expect(new Set(before.map((p) => p.split(':')[0])).size).toBe(3);
		expect(await pagesOf(b)).toEqual(before);
		await a.edit({ type: 'insert-page', pageId: '40', afterPageId: '0', name: 'C' });
		await convergeStable(a, b);
		const after = await pagesOf(b);
		expect(after.filter((p) => before.includes(p))).toEqual(before);
	});

	it('keeps a renumbered master id and the Master references of its batch', async () => {
		const { a, b } = await pair();
		const shapes = masterShape('1', 1);
		await a.patch((zip) => addMaster(zip, { id: '2', unique: '{A}', shapeId: '5', shapes }));
		await b.patch((zip) => addMaster(zip, { id: '2', unique: '{B}', shapeId: '6', shapes }));
		await convergeStable(a, b);
		const idOf = async (peer: DocPeer, u: string) =>
			new RegExp(`<Master ID="(\\d+)"[^>]*UniqueID="${u}"`).exec(
				(await peer.text('visio/masters/masters.xml')) ?? '',
			)![1]!;
		const [ma, mb] = [await idOf(a, '{A}'), await idOf(a, '{B}')];
		expect(ma).not.toBe(mb);
		expect(await idOf(b, '{B}')).toBe(mb);
		expect(a.shared.renum.size).toBeGreaterThan(0);
		expect(readSharedDrawing(a.shared)!.fixups).toEqual([]);
		// A master with a higher id arriving later does not move them.
		await a.patch((zip) => addMaster(zip, { id: '90', unique: '{C}', shapes }));
		await convergeStable(a, b);
		expect(await idOf(b, '{B}')).toBe(mb);
		expect(await idOf(b, '{A}')).toBe(ma);
		const page = (await b.text(PAGE))!;
		expect(new RegExp(`<Shape ID="6"[^>]*Master="${mb}"`).test(page)).toBe(true);
	});
});

const rooms: ReturnType<typeof createRoom>[] = [];
afterEach(() => {
	for (const room of rooms.splice(0)) room.destroy();
});

describe('the fix-up in live sessions', () => {
	it('is written by a peer, converges, and is not undoable by either user', async () => {
		const room = createRoom();
		rooms.push(room);
		const a = room.join('Ada', (await editVsdx(await createVsdx(), [rect('1', 1)])).bytes);
		await wait();
		const b = room.join('Bob', await createVsdx(), { seed: false });
		await until(async () => (await b.side.model()).pages[0]!.shapes.length === 1, 'join');
		await settle(a, b);
		await Promise.all([a.side.edit(rect('9', 3.5)), b.side.edit(rect('9', 6.5))]);
		await until(async () => {
			const [ia, ib] = [await a.side.model(), await b.side.model()];
			return ia.pages[0]!.shapes.length === 3 && ib.pages[0]!.shapes.length === 3;
		}, 'both shapes');
		await settle(a, b);
		const shapes = async (p: typeof a) =>
			(await p.side.model()).pages[0]!.shapes.map((s) => `${s.id}@${s.transform[4]}`).sort();
		expect(await shapes(a)).toEqual(await shapes(b));
		expect(new Set((await shapes(a)).map((s) => s.split('@')[0])).size).toBe(3);
		expect(a.session.doc.getMap('visio:renum').size).toBe(2);
		expect(b.session.doc.getMap('visio:renum').size).toBe(2);
		// Each user's history holds their own add only: the fix-up is not in it.
		expect(a.binding.canUndo()).toBe(true);
		expect(b.binding.canUndo()).toBe(true);
		expect(await a.binding.undo()).toBe(true);
		expect(a.binding.canUndo()).toBe(false);
		await settle(a, b);
		expect(await shapes(a)).toEqual(await shapes(b));
		expect((await shapes(a)).length).toBe(2);
		expect(await b.binding.undo()).toBe(true);
		expect(b.binding.canUndo()).toBe(false);
		await settle(a, b);
		expect((await shapes(a)).length).toBe(1);
		expect(await shapes(b)).toEqual(await shapes(a));
		expect(a.errors).toEqual([]);
		expect(b.errors).toEqual([]);
	});
});
