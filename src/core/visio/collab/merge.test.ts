// Property-level merging: two peers edit one shape at once, shapes are renumbered, and the
// schema version rule. Updates are exchanged by hand so the edits are truly concurrent.
import { describe, expect, it } from 'vitest';
import * as Y from 'yjs';
import { createVsdx, editVsdx } from '../index';
import { SCHEMA_VERSION, readSharedDrawing, sharedDrawing } from './index';
import { DocPeer, converge, joinFrom, seedPeer } from './test-support/doc-peers';
import { buildShape, splitShape } from './pieces';

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
const pinX = (s: { transform: readonly number[] }) => s.transform[4]!;

async function pair() {
	const start = (await editVsdx(await createVsdx(), [rect('1', 1), rect('2', 3)])).bytes;
	const a = await seedPeer(start, 'a');
	const b = await joinFrom(a, 'b');
	return { a, b };
}
const shapeOf = async (peer: DocPeer, id: string) =>
	(await peer.model()).pages[0]!.shapes.find((s) => s.id === id)!;

describe('pieces', () => {
	it('splits a shape into cells, sections, rows and rebuilds it in a stable order', () => {
		const xml =
			'<Shape ID="1" Type="Shape"><Cell N="PinX" V="1"/><Section N="Geometry" IX="0"><Cell N="NoFill" V="0"/>' +
			'<Row IX="1" T="MoveTo"><Cell N="X" V="0"/></Row></Section><Text>a<cp IX="0"/>b</Text>' +
			'<Shapes><Shape ID="2"><Cell N="PinX" V="2"/></Shape></Shapes></Shape>';
		const split = splitShape(xml)!;
		expect(split.pieces.map((p) => p.path)).toEqual([
			'c/PinX',
			's/Geometry:0',
			's/Geometry:0/c/NoFill',
			's/Geometry:0/r/1',
			's/Geometry:0/r/1/c/X',
			'x/Text#0',
		]);
		expect(split.nested).toHaveLength(1);
		const entries = new Map(
			split.pieces.map((p, pos) => [
				p.path,
				p.kind === 'leaf' ? { pos, xml: p.text } : { pos, open: p.text },
			]),
		);
		expect(buildShape(split.open, entries, `<Shapes>${split.nested[0]}</Shapes>`)).toBe(xml);
	});
});

describe('one shape, different properties', () => {
	it('keeps a move and a text and fill change made at the same time', async () => {
		const { a, b } = await pair();
		await a.edit({ type: 'move-shape', pageId: '0', shapeId: '1', x: 5, y: 2 });
		await b.edit(
			{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'Hello' },
			{ type: 'format-shape', pageId: '0', shapeId: '1', fillColor: '#00FF00' },
		);
		await converge(a, b);
		for (const peer of [a, b]) {
			const shape = await shapeOf(peer, '1');
			expect(pinX(shape)).toBe(4.5);
			expect(shape.text.plainText.trim()).toBe('Hello');
			expect(shape.style.fill.toLowerCase()).toBe('#00ff00');
		}
		expect((await a.text('visio/pages/page1.xml'))!.length).toBe(
			(await b.text('visio/pages/page1.xml'))!.length,
		);
	});

	it('resolves one cell edited by both to the same value everywhere', async () => {
		const { a, b } = await pair();
		await a.edit({ type: 'move-shape', pageId: '0', shapeId: '1', x: 5, y: 2 });
		await b.edit({ type: 'move-shape', pageId: '0', shapeId: '1', x: 8, y: 2 });
		await converge(a, b);
		const [xa, xb] = [pinX(await shapeOf(a, '1')), pinX(await shapeOf(b, '1'))];
		expect(xa).toBe(xb);
		expect([4.5, 7.5]).toContain(xa);
		expect(pinX(await shapeOf(a, '2'))).toBe(2.5);
	});

	it('writes only the entries that changed', async () => {
		const { a } = await pair();
		const before = new Map([...a.shared.cells].map(([k, v]) => [k, JSON.stringify(v)]));
		await a.edit({ type: 'move-shape', pageId: '0', shapeId: '1', x: 5, y: 2 });
		const changed = [...a.shared.cells].filter(([k, v]) => before.get(k) !== JSON.stringify(v));
		expect(changed.length).toBeGreaterThan(0);
		expect(changed.every(([k]) => k.includes('/s1/') || k.includes('/s1'))).toBe(true);
		expect(changed.length).toBeLessThan(5);
	});

	it('lets a deleted shape win over a concurrent edit of its cells', async () => {
		const { a, b } = await pair();
		await a.edit({ type: 'delete-shape', pageId: '0', shapeId: '1' });
		await b.edit({ type: 'format-shape', pageId: '0', shapeId: '1', fillColor: '#FF0000' });
		await converge(a, b);
		for (const peer of [a, b])
			expect((await peer.model()).pages[0]!.shapes.map((s) => s.id)).toEqual(['2']);
	});

	it('rebuilds the same page text on every peer', async () => {
		const { a, b } = await pair();
		await a.edit({ type: 'move-shape', pageId: '0', shapeId: '2', x: 6, y: 2 });
		await b.edit({ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'x' });
		await converge(a, b);
		expect(await a.text('visio/pages/page1.xml')).toBe(await b.text('visio/pages/page1.xml'));
	});
});

describe('renumbered shapes and their formulas', () => {
	async function addNine(peer: DocPeer, x: number, ref: boolean) {
		await peer.edit(rect('9', x));
		if (ref)
			await peer.patch(async (zip) => {
				const path = 'visio/pages/page1.xml';
				const xml = await zip.file(path)!.async('string');
				zip.file(
					path,
					xml.replace(
						/(<Shape ID="9"[^>]*>[\s\S]*?)<Cell N="LocPinX" V="0.5" F="Width\*0.5"\/>/,
						'$1<Cell N="LocPinX" V="0.5" F="Sheet.9!Width*0.5"/>',
					),
				);
			});
	}

	it('rewrites Sheet references inside the renumbered shape and its batch mates', async () => {
		const { a, b } = await pair();
		await addNine(a, 1.5, true);
		// Same batch on b: shape 9 (the one renumbered) and shape 10 that points at it.
		await b.edit(rect('9', 6.5), rect('10', 7.5));
		await b.patch(async (zip) => {
			const path = 'visio/pages/page1.xml';
			const xml = await zip.file(path)!.async('string');
			zip.file(
				path,
				xml
					.replace(
						/(<Shape ID="9"[^>]*>[\s\S]*?)<Cell N="LocPinX" V="0.5" F="Width\*0.5"\/>/,
						'$1<Cell N="LocPinX" V="0.5" F="Sheet.9!Width*0.5"/>',
					)
					.replace(
						/(<Shape ID="10"[^>]*>[\s\S]*?)<Cell N="LocPinY" V="0.5" F="Height\*0.5"\/>/,
						'$1<Cell N="LocPinY" V="0.5" F="Sheet.9!Height*0.5"/>',
					),
			);
		});
		await converge(a, b);
		const [pa, pb] = [await a.text('visio/pages/page1.xml'), await b.text('visio/pages/page1.xml')];
		expect(pa).toBe(pb);
		const ids = [...pa!.matchAll(/<Shape ID="(\d+)"/g)].map((m) => m[1]);
		expect(new Set(ids).size).toBe(ids.length);
		// Whoever lost id 9 kept formulas that point at its own new id, never at the other shape.
		const formulas = [...pa!.matchAll(/F="Sheet\.(\d+)!(Width|Height)\*0\.5"/g)].map((m) => m[1]!);
		expect(formulas).toHaveLength(3);
		const refsTo = (id: string) => formulas.filter((f) => f === id).length;
		expect(Math.max(...formulas.map(Number))).toBeGreaterThan(9);
		// Shape 9 (kept) is referenced by exactly one formula; the renumbered shape's id by two.
		const renumbered = formulas.find((f) => f !== '9')!;
		expect(refsTo('9')).toBe(1);
		expect(refsTo(renumbered)).toBe(2);
	});
});

describe('schema version', () => {
	it('stamps the room with the current version', async () => {
		const { a } = await pair();
		expect(a.shared.meta.get('v')).toBe(SCHEMA_VERSION);
		expect(SCHEMA_VERSION).toBe(4);
	});

	it('refuses to read or write a room of another version', async () => {
		for (const v of [2, 3, 5]) {
			const { a } = await pair();
			a.shared.meta.set('v', v);
			expect(() => readSharedDrawing(a.shared)).toThrow(/schema version/);
			await expect(a.publish()).rejects.toThrow(/schema version/);
		}
		const doc = new Y.Doc();
		expect(readSharedDrawing(sharedDrawing(doc))).toBeUndefined();
	});
});
