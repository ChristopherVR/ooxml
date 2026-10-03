import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index.js';
import { editVsdx } from './edit.js';
import { parseVsdx } from './parser.js';
import { fixture, cell, shape, section, row, rectangle } from './test-fixtures.js';

const dimensions =
	cell('PinX', 2) +
	cell('PinY', 3) +
	cell('Width', 3) +
	cell('Height', 4) +
	cell('LocPinX', 1.5, 'Width*0.5') +
	cell('LocPinY', 2, 'Height*0.5');
const user = (name: string, value: number, formula = '') =>
	`<Row N="${name}">${cell('Value', value, formula)}</Row>`;
const geometry = section(
	'Geometry',
	row(1, 'MoveTo', cell('X', 0) + cell('Y', 0)) +
		row(
			2,
			'LineTo',
			cell('X', 1.5, 'IF(BITAND(User.Mask,1),Width*0,Width*0.5)') +
				cell('Y', 4, 'IF(AND(User.Enabled,NOT(User.Disabled)),Height,0)'),
		) +
		row(3, 'LineTo', cell('X', 5, 'SQRT(Width^2+Height^2)') + cell('Y', 0)),
);
const users = section(
	'User',
	user('Mask', 0, 'IF(Width/Height>1,1,0)') +
		user('Enabled', 1, 'OR(Width>0,Height>0)') +
		user('Disabled', 0) +
		user('Remainder', 0, 'MODULUS(Width/Height,1)'),
);
const source = (extra = '', extraUser = '') =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', dimensions + users.replace('</Section>', extraUser + '</Section>') + geometry + extra)}${shape('10', dimensions + rectangle + section('User', user('Follow', 1, 'BITXOR(Sheet.1!User.Mask,1)')))}</Shapes>`,
			},
			{
				id: '1',
				contents: `<Shapes>${shape('2', dimensions + rectangle + section('User', user('Independent', 9, 'ROUND(8.9,0)')))}</Shapes>`,
			},
		],
		edit: (zip) => {
			zip.file('custom/unknown.bin', new Uint8Array([17, 33, 0, 255]));
		},
	});
const resize = { type: 'resize-shape' as const, pageId: '0', shapeId: '1', width: 6, height: 2 };
async function caches(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	const doc = parseXml(await zip.file('visio/pages/page1.xml')!.async('string'));
	const formulas = new Map<string, string>();
	for (const node of Array.from(doc.documentElement.getElementsByTagName('*')))
		if (node.hasAttribute('F')) formulas.set(node.getAttribute('F')!, node.getAttribute('V')!);
	return { zip, formulas };
}
describe('numeric formula cache transactions', () => {
	it('recalculates transitive bit/logical geometry and squared-distance caches atomically', async () => {
		const original = await source(),
			snapshot = original.slice();
		const result = await editVsdx(original, [resize]);
		const before = await caches(original),
			after = await caches(result.bytes);
		expect(after.formulas.get('IF(Width/Height>1,1,0)')).toBe('1');
		expect(after.formulas.get('IF(BITAND(User.Mask,1),Width*0,Width*0.5)')).toBe('0');
		expect(after.formulas.get('IF(AND(User.Enabled,NOT(User.Disabled)),Height,0)')).toBe('2');
		expect(Number(after.formulas.get('SQRT(Width^2+Height^2)'))).toBeCloseTo(Math.sqrt(40), 12);
		expect(after.formulas.get('MODULUS(Width/Height,1)')).toBe('0');
		expect(after.formulas.get('BITXOR(Sheet.1!User.Mask,1)')).toBe('0');
		expect([...after.formulas.keys()].sort()).toEqual([...before.formulas.keys()].sort());
		for (const [name, entry] of Object.entries(before.zip.files))
			if (!entry.dir && !result.changedParts.includes(name))
				expect(await after.zip.file(name)!.async('uint8array'), name).toEqual(
					await entry.async('uint8array'),
				);
		expect(result.changedParts).toEqual(['visio/pages/page1.xml']);
		expect(original).toEqual(snapshot);
		const parsed = await parseVsdx(result.bytes);
		expect([parsed.pages[0]!.shapes[0]!.width, parsed.pages[0]!.shapes[0]!.height]).toEqual([6, 2]);
		const roundtrip = await editVsdx(result.bytes, [{ ...resize, width: 3, height: 4 }]);
		expect(
			(await caches(roundtrip.bytes)).formulas.get('IF(BITAND(User.Mask,1),Width*0,Width*0.5)'),
		).toBe('1.5');
		expect((await caches(roundtrip.bytes)).formulas.get('BITXOR(Sheet.1!User.Mask,1)')).toBe('1');
	});
	it('retains protections and mutation isolation for failures after numeric closure work', async () => {
		for (const [extra, extraUser] of [
			[cell('LockWidth', 1), ''],
			['', user('Bad', 0, 'BITAND(Width/Height*65536,1)')],
			['', user('WrongUnit', 0, 'Width*Height')],
			['', user('Missing', 0, 'AND(Width,Missing)')],
			['', user('CycleA', 0, 'IF(Width>0,User.CycleB,0)') + user('CycleB', 0, 'NOT(User.CycleA)')],
			['', user('Unsupported', 0, 'INDIRECT(Width)')],
		]) {
			const original = await source(extra, extraUser),
				snapshot = original.slice();
			const transaction = editVsdx(original, [resize]);
			if (extraUser?.includes('N="Bad"'))
				await expect(transaction).rejects.toThrow(/integer from 0 to 65535/);
			else if (extraUser?.includes('N="WrongUnit"'))
				await expect(transaction).rejects.toMatchObject({ code: 'EDIT_FORMULA_UNIT' });
			else await expect(transaction).rejects.toThrow();
			expect(original).toEqual(snapshot);
		}
		const protectedSource = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', dimensions.replace(cell('Width', 3), cell('Width', 3, 'GUARD(3)')) + rectangle)}</Shapes>`,
				},
			],
		});
		await expect(editVsdx(protectedSource, [resize])).rejects.toMatchObject({
			code: 'EDIT_PROTECTED_CELL',
		});
	});
});
