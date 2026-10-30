import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

// Optional: `scripts/record-visio-controls.ps1 -OutputDirectory <dir>` saves real drops of Basic
// Shapes masters with a control handle (Can, Parallelogram, Trapezoid, Chevron, Cube, Cone) before
// and after Visio moves the handle's free axis to 30% of the shape size. Each drag must write what
// Visio wrote, including the Cone, whose BOUND limit is another geometry cell.
const native = process.env.VISIO_NATIVE_CONTROLS_DIR;
const PAGE = 'visio/pages/page1.xml';
type Cells = Map<string, [number | string, string | null, string | null]>;

async function shapeCells(bytes: Uint8Array, id: string): Promise<Cells> {
	const xml = await (await JSZip.loadAsync(bytes)).file(PAGE)!.async('string');
	const root = new DOMParser().parseFromString(xml, 'text/xml').documentElement!;
	const shape = Array.from(root.getElementsByTagName('Shape')).find(
		(node) => node.getAttribute('ID') === id,
	)!;
	const result: Cells = new Map();
	for (const cell of Array.from(shape.getElementsByTagName('Cell'))) {
		const path: string[] = [];
		for (
			let node = cell.parentNode as typeof shape | null;
			node && node !== shape.parentNode;
			node = node.parentNode as typeof shape | null
		) {
			if (node.localName === 'Shape') path.unshift(`Shape#${node.getAttribute('ID')}`);
			else if (node.localName !== 'Shapes')
				path.unshift(
					`${node.localName}[${node.getAttribute('N') ?? ''}/${node.getAttribute('T') ?? ''}/${node.getAttribute('IX') ?? ''}]`,
				);
		}
		const value = cell.getAttribute('V') ?? '';
		result.set([...path, cell.getAttribute('N')].join('.'), [
			/^-?[\d.]+(e-?\d+)?$/i.test(value) ? Number(value) : value,
			cell.getAttribute('U'),
			cell.getAttribute('F'),
		]);
	}
	return result;
}

/** Menu, data and scratch caches may need text metrics; they are left for Visio to recalculate. */
const PASSIVE = /\.Section\[(Actions|User|Property|Scratch)\//;

interface Entry {
	master: string;
	shapeId: string;
	axis: 'X' | 'Y';
	ratio: number;
}

describe.skipIf(!native)('native Visio control handles', () => {
	it('writes what Visio writes when a handle is dragged', async () => {
		const before = new Uint8Array(await readFile(join(native!, 'before.vsdx')));
		const after = new Uint8Array(await readFile(join(native!, 'after.vsdx')));
		const manifest = JSON.parse(
			(await readFile(join(native!, 'manifest.json'), 'utf8')).replace(/^﻿/, ''),
		) as Entry[];
		const model = await parseVsdx(before);
		expect(manifest.length).toBeGreaterThan(3);
		for (const entry of manifest) {
			const shape = model.pages[0]!.shapes.find((item) => item.id === entry.shapeId)!;
			const handle = shape.controls![0]!;
			const size = entry.axis === 'X' ? shape.width : shape.height;
			const saved = await editVsdx(before, [
				{
					type: 'set-control',
					pageId: '0',
					shapeId: entry.shapeId,
					row: handle.row,
					[entry.axis.toLowerCase()]: size * entry.ratio,
				},
			]);
			const ours = await shapeCells(saved.bytes, entry.shapeId);
			const theirs = await shapeCells(after, entry.shapeId);
			const start = await shapeCells(before, entry.shapeId);
			for (const [key, [value, unit, formula]] of ours) {
				if (start.has(key) && !theirs.has(key)) continue;
				const expected = theirs.get(key);
				if (!expected) {
					expect(PASSIVE.test(key), `${entry.master} ${key}`).toBe(true);
					continue;
				}
				if (typeof value === 'number')
					expect(value, `${entry.master} ${key}`).toBeCloseTo(expected[0] as number, 9);
				else expect(value, key).toBe(expected[0]);
				expect([unit, formula], `${entry.master} ${key}`).toEqual([expected[1], expected[2]]);
			}
			// Everything Visio refreshed is refreshed here too.
			for (const key of theirs.keys())
				if (!start.has(key) || JSON.stringify(start.get(key)) !== JSON.stringify(theirs.get(key)))
					expect(ours.has(key), `${entry.master} missing ${key}`).toBe(true);
			const moved = (await parseVsdx(saved.bytes)).pages[0]!.shapes.find(
				(item) => item.id === entry.shapeId,
			)!;
			const real = (await parseVsdx(after)).pages[0]!.shapes.find(
				(item) => item.id === entry.shapeId,
			)!;
			expect(JSON.stringify(moved.geometry), entry.master).toBe(JSON.stringify(real.geometry));
			expect(moved.controls![0]!.x, entry.master).toBeCloseTo(real.controls![0]!.x, 9);
			expect(moved.controls![0]!.y, entry.master).toBeCloseTo(real.controls![0]!.y, 9);
		}
	});
});
