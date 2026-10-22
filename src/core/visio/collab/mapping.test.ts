import * as Y from 'yjs';
import { describe, expect, it } from 'vitest';
import { createVsdx, editVsdx, parseVsdx } from '../index';
import { composeDrawing } from './compose';
import { DrawingKeys } from './keys';
import { assignPositions } from './order';
import { readSharedDrawing, renumberShape } from './read';
import { scanContainer } from './scan';
import { sharedDrawing } from './schema';
import { snapshotDrawing } from './snapshot';
import { writeSharedDrawing } from './write';

describe('scanContainer', () => {
	it('returns the direct children by span, skipping comments, CDATA and quoted ">"', () => {
		const xml = `<R><S a="x>y"><!-- <S/> --><S ID="2"/></S><S ID="3"/><![CDATA[<S>]]></R>`;
		const box = scanContainer(xml, 'R')!;
		expect(box.children.map((c) => xml.slice(c.start, c.end))).toEqual([
			`<S a="x>y"><!-- <S/> --><S ID="2"/></S>`,
			`<S ID="3"/>`,
		]);
	});
	it('reports a malformed container as undefined', () => {
		expect(scanContainer('<R><S></R>', 'R')).toBeUndefined();
	});
});

describe('assignPositions', () => {
	it('keeps increasing positions and places new items between neighbours', () => {
		const current = new Map([
			['a', 1],
			['c', 2],
		]);
		const result = assignPositions(['a', 'b', 'c'], (k) => current.get(k));
		expect(result.get('a')).toBe(1);
		expect(result.get('c')).toBe(2);
		expect(result.get('b')).toBeGreaterThan(1);
		expect(result.get('b')).toBeLessThan(2);
	});
});

describe('renumberShape', () => {
	it('rewrites shape ids and Sheet references of the renumbered ids only', () => {
		const xml = `<Shape ID="7"><Cell N="PinX" F="Sheet.8!Width+Sheet.3!Width"/><Shapes><Shape ID="8"/></Shapes></Shape>`;
		expect(renumberShape(xml, new Map([['8', '12']]))).toBe(
			`<Shape ID="7"><Cell N="PinX" F="Sheet.12!Width+Sheet.3!Width"/><Shapes><Shape ID="12"/></Shapes></Shape>`,
		);
	});
});

describe('shared drawing round trip', () => {
	it('splits a package into the shared maps and rebuilds the same drawing', async () => {
		const shapes = (
			await editVsdx(await createVsdx(), [
				{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 1, y: 2, width: 1, height: 1 },
				{ type: 'create-ellipse', pageId: '0', shapeId: '2', x: 4, y: 2, width: 1, height: 1 },
			])
		).bytes;
		const source = (
			await editVsdx(shapes, [{ type: 'insert-page', pageId: '3', afterPageId: '0', name: 'Two' }])
		).bytes;
		const keys = new DrawingKeys('t');
		keys.seeding = true;
		const doc = new Y.Doc();
		const shared = sharedDrawing(doc);
		const writes = writeSharedDrawing(shared, await snapshotDrawing(source, keys), keys);
		expect(writes).toBeGreaterThan(0);
		expect(shared.shapes.size).toBe(2);
		expect([...shared.pages.keys()]).toEqual(['s0', 's3']);
		// A second write of the same drawing changes nothing.
		expect(writeSharedDrawing(shared, await snapshotDrawing(source, keys), keys)).toBe(0);

		const rebuilt = await composeDrawing(
			readSharedDrawing(shared)!,
			undefined,
			new DrawingKeys('u'),
		);
		const [before, after] = [await parseVsdx(source), await parseVsdx(rebuilt.bytes)];
		expect(after.pages.map((p) => [p.id, p.name, p.shapes.length])).toEqual(
			before.pages.map((p) => [p.id, p.name, p.shapes.length]),
		);
		expect(after.pages[0]!.shapes.map((s) => s.geometry)).toEqual(
			before.pages[0]!.shapes.map((s) => s.geometry),
		);
	});
});
