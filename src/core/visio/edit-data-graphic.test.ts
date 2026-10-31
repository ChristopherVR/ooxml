import { describe, expect, it } from 'vitest';
import { writeFileSync } from 'node:fs';
import { createVsdx } from './create-document';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import type { VisioEdit } from './edit-commands';

const row = (name: string, value: string, type: 'number' | 'string' = 'number'): VisioEdit => ({
	type: 'set-shape-data',
	pageId: '0',
	shapeId: '1',
	row: name,
	data: { label: name, prompt: '', type, format: '', value },
});
async function owner(): Promise<Uint8Array> {
	const drawn = await editVsdx(await createVsdx(), [
		{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 3, y: 4, width: 2, height: 1 },
		row('Sales', '42'),
		row('Region', 'North', 'string'),
	]);
	return drawn.bytes;
}
const part = async (bytes: Uint8Array, name: string) =>
	new TextDecoder().decode(await (await VisioPackage.open(bytes)).readBytes(name));
const graphic = (items: Extract<VisioEdit, { type: 'set-data-graphic' }>['items']): VisioEdit => ({
	type: 'set-data-graphic',
	pageId: '0',
	shapeId: '1',
	items,
});
const ITEMS = [
	{ kind: 'text', field: 'Region' },
	{ kind: 'bar', field: 'Sales', min: 0, max: 50 },
	{ kind: 'icon', field: 'Sales', min: 0, max: 50, position: 'top-right' },
] as const;
const first = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;

describe('native data graphics', () => {
	it('writes the shape as a group of callout instances named by a hidden data graphic master', async () => {
		const result = await editVsdx(await owner(), [graphic([...ITEMS])]);
		if (process.env.DG_OUT) writeFileSync(process.env.DG_OUT, result.bytes);
		const shape = await first(result.bytes);
		expect(shape.kind).toBe('group');
		expect(shape.children.map((child) => child.name)).toEqual([
			'Text callout',
			'Data bar 1',
			'Status icons',
		]);
		const display = shape.userCells!.find((cell) => cell.name === 'visDGDisplayFormat')!;
		expect(display.formula).toMatch(/^USE\(\{[0-9A-F-]{36}\}\)$/);
		const masters = await part(result.bytes, 'visio/masters/masters.xml');
		const uid = /\{[0-9A-F-]{36}\}/.exec(display.formula!)![0];
		expect(masters).toMatch(
			new RegExp(`Hidden="1"[^>]*MasterType="2"|MasterType="2"[^>]*Hidden="1"|UniqueID="${uid}"`),
		);
		expect(masters).toContain(`UniqueID="${uid}"`);
		for (const name of ['Text callout', 'Data bar 1', 'Status icons'])
			expect(masters).toContain(`NameU="${name}"`);
		const page = await part(result.bytes, 'visio/pages/page1.xml');
		expect(page).toContain('F="Sheet.1!Prop.Sales"');
		expect(page).toContain('F="Sheet.1!Prop.Region"');
		// Each callout holds a value, a field it reads and its own place.
		const bar = shape.children[1]!;
		expect(bar.userCells!.find((cell) => cell.name === 'visDGField0')!.value).toBe('{Sales}');
		expect(bar.userCells!.find((cell) => cell.name === 'PercentFilled')!.value).toBe('0.84');
		expect(bar.shapeData!.find((cell) => cell.name === 'msvCalloutField')!.value).toBe(42);
		expect(result.changedParts).toContain('visio/masters/masters.xml');
	});

	it('keeps the callouts, their IDs and the group when Shape Data changes', async () => {
		const drawn = await editVsdx(await owner(), [graphic([...ITEMS])]);
		const before = (await first(drawn.bytes)).children.map((child) => child.id);
		const changed = await editVsdx(drawn.bytes, [
			row('Sales', '10'),
			row('Region', 'South', 'string'),
		]);
		const shape = await first(changed.bytes);
		expect(shape.children.map((child) => child.id)).toEqual(before);
		expect(shape.children[1]!.userCells!.find((c) => c.name === 'PercentFilled')!.value).toBe(
			'0.2',
		);
		const page = await part(changed.bytes, 'visio/pages/page1.xml');
		expect(page).toContain('V="Region: South"');
		expect(page).not.toContain('Region: North');
		// The data graphic master is reused, not added again.
		const masters = await part(changed.bytes, 'visio/masters/masters.xml');
		expect(masters.match(/MasterType="2"/g)).toHaveLength(1);
	});

	it('rewrites the callouts for a new list and removes the graphic with an empty one', async () => {
		let bytes = (await editVsdx(await owner(), [graphic([...ITEMS])])).bytes;
		bytes = (await editVsdx(bytes, [graphic([{ kind: 'text', field: 'Sales', position: 'top' }])]))
			.bytes;
		let shape = await first(bytes);
		expect(shape.children).toHaveLength(1);
		expect(shape.children[0]!.userCells!.find((c) => c.name === 'visDGDefaultPos')!.value).toBe(
			'PNT(2,5)',
		);
		bytes = (await editVsdx(bytes, [graphic([])])).bytes;
		shape = await first(bytes);
		expect(shape.kind).toBe('shape');
		expect(shape.children).toEqual([]);
		expect(shape.userCells?.some((cell) => cell.name === 'visDGDisplayFormat') ?? false).toBe(
			false,
		);
		expect(shape.shapeData!.map((data) => data.name)).toEqual(['Sales', 'Region']);
	});

	it('refuses a shape that is an instance of a master and a row the shape does not have', async () => {
		const dropped = await editVsdx(await createVsdx(), [
			{ type: 'drop-stencil-master', pageId: '0', shapeId: '1', master: 'rectangle', x: 2, y: 2 },
			row('Sales', '1'),
		]);
		await expect(
			editVsdx(dropped.bytes, [graphic([{ kind: 'text', field: 'Sales' }])]),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_DATA_GRAPHIC' });
		await expect(
			editVsdx(await owner(), [graphic([{ kind: 'text', field: 'Missing' }])]),
		).rejects.toMatchObject({ code: 'EDIT_TARGET_NOT_FOUND' });
		await expect(
			editVsdx(await owner(), [graphic([{ kind: 'text', field: 'Bad Name!' }])]),
		).rejects.toMatchObject({ code: 'INVALID_EDIT' });
	});

	it('runs the graphic first when other edits come with it', async () => {
		const result = await editVsdx(await owner(), [
			graphic([{ kind: 'bar', field: 'Sales', min: 0, max: 100 }]),
			row('Sales', '30'),
		]);
		const shape = await first(result.bytes);
		expect(shape.children).toHaveLength(1);
		expect(shape.children[0]!.userCells!.find((c) => c.name === 'PercentFilled')!.value).toBe(
			'0.3',
		);
	});
});
