import { DOMParser } from '@xmldom/xmldom';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit.js';
import { parseVsdx } from './parser.js';
import { cell, fixture, rectangle, row, section, shape } from './test-fixtures.js';

const dimensions = () => cell('PinX', 2) + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1);
const existing = (extra = '', geometry = rectangle, attrs = '') =>
	shape('1', dimensions() + geometry + '<Text>Existing</Text>' + extra, attrs);
const source = (shapes = existing(), trailing = '') =>
	fixture({ pages: [{ id: '0', contents: `<Shapes>${shapes}</Shapes>${trailing}` }] });
const move = (x = 5, y = 6): Extract<VisioEdit, { type: 'move-shape' }> => ({
	type: 'move-shape',
	pageId: '0',
	shapeId: '1',
	x,
	y,
});
const resize = (width = 4, height = 2): Extract<VisioEdit, { type: 'resize-shape' }> => ({
	type: 'resize-shape',
	pageId: '0',
	shapeId: '1',
	width,
	height,
});
const remove: VisioEdit = { type: 'delete-shape', pageId: '0', shapeId: '1' };
const create = (): Extract<VisioEdit, { type: 'create-rectangle' }> => ({
	type: 'create-rectangle',
	pageId: '0',
	shapeId: '9',
	x: 4,
	y: 5,
	width: 3,
	height: 2,
	text: 'Created & <safe>',
});
async function pageXml(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	return new DOMParser().parseFromString(
		await zip.file('visio/pages/page1.xml')!.async('string'),
		'text/xml',
	);
}
async function cache(bytes: Uint8Array, id: string, name: string) {
	const doc = await pageXml(bytes);
	const target = Array.from(doc.getElementsByTagName('Shape')).find(
		(s) => s.getAttribute('ID') === id,
	)!;
	return Array.from(target.getElementsByTagName('Cell')).find((c) => c.getAttribute('N') === name)!;
}
async function number(bytes: Uint8Array, id: string, name: string) {
	return Number((await cache(bytes, id, name)).getAttribute('V'));
}

describe('source-backed geometry edit contract', () => {
	it('creates, moves, resizes and deletes a rectangle across fresh round trips', async () => {
		let bytes = (await editVsdx(await source(), [create()])).bytes;
		let created = (await parseVsdx(bytes)).pages[0]!.shapes.find((s) => s.id === '9')!;
		expect(created.text.plainText).toBe('Created & <safe>');
		expect(created.geometry[0]?.path).toBeTruthy();
		expect([created.width, created.height]).toEqual([3, 2]);
		bytes = (await editVsdx(bytes, [{ ...move(), shapeId: '9', x: 7, y: 8 }])).bytes;
		bytes = (await editVsdx(bytes, [{ ...resize(), shapeId: '9' }])).bytes;
		created = (await parseVsdx(bytes)).pages[0]!.shapes.find((s) => s.id === '9')!;
		expect([created.width, created.height]).toEqual([4, 2]);
		expect(await number(bytes, '9', 'PinX')).toBe(7);
		expect(await number(bytes, '9', 'PinY')).toBe(8);
		bytes = (await editVsdx(bytes, [{ ...remove, shapeId: '9' }])).bytes;
		expect((await parseVsdx(bytes)).pages[0]!.shapes.map((s) => s.id)).toEqual(['1']);
	});
	it('moves an existing admitted local shape in drawing inches and preserves dimensions', async () => {
		const bytes = (await editVsdx(await source(), [move(-2, 8)])).bytes;
		expect(await number(bytes, '1', 'PinX')).toBe(-2);
		expect(await number(bytes, '1', 'PinY')).toBe(8);
		const parsed = (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
		expect([parsed.width, parsed.height]).toEqual([2, 1]);
		expect(parsed.transform.slice(4)).toEqual([-3, 7.5]);
	});
	it('scales relative geometry with dimensions while holding the pin and using default local pin', async () => {
		const bytes = (await editVsdx(await source(), [resize()])).bytes;
		const parsed = (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
		expect(parsed.transform.slice(4)).toEqual([0, 2]);
		expect(await number(bytes, '1', 'PinX')).toBe(2);
		expect(await number(bytes, '1', 'PinY')).toBe(3);
		expect(parsed.geometry[0]?.path).toContain('4');
	});
	it('recalculates absolute geometry and explicit local pin formulas without removing formulas', async () => {
		const absolute = section(
			'Geometry',
			row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) +
				row(2, 'LineTo', cell('X', 2, 'Width') + cell('Y', 1, 'Height')),
		);
		const bytes = (
			await editVsdx(
				await source(
					existing(cell('LocPinX', 1, 'Width*0.5') + cell('LocPinY', 0.5, 'Height*0.5'), absolute),
				),
				[resize()],
			)
		).bytes;
		expect(await number(bytes, '1', 'LocPinX')).toBe(2);
		expect(await number(bytes, '1', 'LocPinY')).toBe(1);
		const doc = await pageXml(bytes);
		const coordinates = Array.from(doc.getElementsByTagName('Cell')).filter((c) =>
			c.hasAttribute('F'),
		);
		expect(coordinates.map((c) => [c.getAttribute('F'), c.getAttribute('V')])).toEqual([
			['Width', '4'],
			['Height', '2'],
			['Width*0.5', '2'],
			['Height*0.5', '1'],
		]);
	});
	it('writes the affected static cross-shape dependency closure', async () => {
		const dependent = shape(
			'2',
			cell('PinX', 4, 'Sheet.1!Width*2') +
				cell('PinY', 3) +
				cell('Width', 1) +
				cell('Height', 1) +
				rectangle,
		);
		const bytes = (await editVsdx(await source(existing() + dependent), [resize()])).bytes;
		expect(await number(bytes, '2', 'PinX')).toBe(8);
		expect((await cache(bytes, '2', 'PinX')).getAttribute('F')).toBe('Sheet.1!Width*2');
	});
	it('recalculates a transitive IF dependency while preserving every formula', async () => {
		const extra =
			cell('LocPinX', 1, 'IF(Width&gt;3 IN,Width*0.25,Width*0.5)') + cell('TxtPinX', 1, 'LocPinX');
		const bytes = (await editVsdx(await source(existing(extra)), [resize(8, 2)])).bytes;
		expect(await number(bytes, '1', 'LocPinX')).toBe(2);
		expect(await number(bytes, '1', 'TxtPinX')).toBe(2);
		expect((await cache(bytes, '1', 'TxtPinX')).getAttribute('F')).toBe('LocPinX');
	});
	it('refuses to resize absolute cached coordinates with no provable scaling formula', async () => {
		const absolute = section(
			'Geometry',
			row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) + row(2, 'LineTo', cell('X', 2) + cell('Y', 1)),
		);
		await expect(editVsdx(await source(existing('', absolute)), [resize()])).rejects.toThrow();
	});
	it('preserves opaque parts, relationships, styles and independent formulas byte for byte', async () => {
		const independent = section('User', row(0, '', cell('Value', 17, 'ROUND(1,0)')));
		const original = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${existing(independent, rectangle, 'LineStyle="42" FillStyle="43" TextStyle="44"')}</Shapes>`,
				},
			],
			document:
				'<StyleSheets><StyleSheet ID="42"/><StyleSheet ID="43"/><StyleSheet ID="44"/></StyleSheets>',
			edit: (zip) => {
				zip.file('custom/opaque.bin', new Uint8Array([0, 255, 19, 3]));
			},
		});
		const result = await editVsdx(original, [move()]);
		expect(result.changedParts).toEqual(['visio/pages/page1.xml']);
		const before = await JSZip.loadAsync(original),
			after = await JSZip.loadAsync(result.bytes);
		const partNames = (zip: JSZip) =>
			Object.values(zip.files)
				.filter((entry) => !entry.dir)
				.map((entry) => entry.name)
				.sort();
		expect(partNames(after)).toEqual(partNames(before));
		for (const [path, entry] of Object.entries(before.files)) {
			if (!entry.dir && path !== 'visio/pages/page1.xml')
				expect(await after.file(path)!.async('uint8array')).toEqual(
					await entry.async('uint8array'),
				);
		}
		const doc = await pageXml(result.bytes),
			target = doc.getElementsByTagName('Shape')[0]!;
		expect(['LineStyle', 'FillStyle', 'TextStyle'].map((n) => target.getAttribute(n))).toEqual([
			'42',
			'43',
			'44',
		]);
		expect((await cache(result.bytes, '1', 'Value')).getAttribute('F')).toBe('ROUND(1,0)');
		expect(await number(result.bytes, '1', 'Value')).toBe(17);
	});
	it('rejects a late failed batch atomically and owns mutable commands and input before awaiting', async () => {
		const original = await source(),
			snapshot = original.slice();
		await expect(editVsdx(original, [move(), { ...remove, shapeId: 'missing' }])).rejects.toThrow();
		expect(original).toEqual(snapshot);
		expect((await editVsdx(original, [])).bytes).toEqual(snapshot);
		const mutable = { type: 'move-shape' as const, pageId: '0', shapeId: '1', x: 5, y: 6 };
		const pending = editVsdx(original, [mutable]);
		original.fill(0);
		mutable.x = 999;
		mutable.shapeId = 'missing';
		const saved = (await pending).bytes;
		expect(await number(saved, '1', 'PinX')).toBe(5);
	});
	it.each([
		['LockMoveX', move()],
		['LockMoveY', move()],
		['LockWidth', resize()],
		['LockHeight', resize()],
		['LockAspect', resize(4, 3)],
	] as const)('honors active protection %s', async (lock, command) => {
		await expect(editVsdx(await source(existing(cell(lock, 1))), [command])).rejects.toThrow();
	});
	it.each(['GUARD(2)', 'SETATREF(User.X)', 'GUARD(SETATREF(User.X))'])(
		'does not overwrite protected or redirected PinX %s',
		async (formula) => {
			const target = shape(
				'1',
				cell('PinX', 2, formula) +
					cell('PinY', 3) +
					cell('Width', 2) +
					cell('Height', 1) +
					rectangle,
			);
			await expect(editVsdx(await source(target), [move()])).rejects.toThrow();
		},
	);
	it.each(['UNKNOWN(Width)', 'INDIRECT("Width")', 'EVALCELL("Width")'])(
		'fails safely for affected or dynamic formula %s',
		async (formula) => {
			const extra = section(
				'User',
				row(0, '', cell('Value', 17, formula.replaceAll('"', '&quot;'))),
			);
			await expect(editVsdx(await source(existing(extra)), [resize()])).rejects.toThrow();
		},
	);
	it('rejects deletion with a static reference or a Connects edge', async () => {
		const dependent = shape('2', dimensions() + cell('LocPinX', 1, 'Sheet.1!Width') + rectangle);
		await expect(editVsdx(await source(existing() + dependent), [remove])).rejects.toThrow();
		await expect(
			editVsdx(
				await source(
					existing() + shape('2', dimensions() + rectangle),
					'<Connects><Connect FromSheet="2" ToSheet="1" FromCell="BeginX" ToCell="PinX"/></Connects>',
				),
				[remove],
			),
		).rejects.toThrow();
	});
	it.each(['Master="7"', 'MasterShape="4"', 'Type="Group"', 'Type="Foreign"'])(
		'refuses unsafe target semantics %s',
		async (attrs) => {
			await expect(
				editVsdx(await source(existing('', rectangle, attrs)), [resize()]),
			).rejects.toThrow();
		},
	);
	it('refuses 1D shapes and nested targets without pretending unglued lines are glued connectors', async () => {
		await expect(editVsdx(await source(existing(cell('OneD', 1))), [move()])).rejects.toThrow();
		const group = shape('10', `<Shapes>${existing()}</Shapes>`, 'Type="Group"');
		await expect(editVsdx(await source(group), [move()])).rejects.toThrow();
	});
	it('rejects duplicate creation and invalid nonfinite or nonpositive dimensions', async () => {
		await expect(editVsdx(await source(), [{ ...create(), shapeId: '1' }])).rejects.toThrow();
		for (const command of [move(NaN), move(Infinity), resize(0), resize(-1), resize(4, NaN)])
			await expect(editVsdx(await source(), [command])).rejects.toThrow();
	});
});
