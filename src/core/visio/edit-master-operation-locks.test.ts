import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit.js';
import { parseVsdx } from './parser.js';
import { VisioPackageError } from './package-common.js';
import { cell, fixture, rectangle, shape } from './test-fixtures.js';
const locks = ['LockMoveX', 'LockMoveY', 'LockWidth', 'LockHeight', 'LockAspect', 'LockDelete'];
const independent = ['LockWidth', 'LockHeight', 'LockAspect', 'LockDelete'];
const pins = cell('PinX', 2) + cell('PinY', 3);
const dimensions =
	cell('Width', 3) +
	cell('Height', 4) +
	cell('LocPinX', 1.5, 'Width*0.5') +
	cell('LocPinY', 2, 'Height*0.5');
const move = { type: 'move-shape' as const, pageId: '0', shapeId: '1', x: 5, y: 6 };
async function source(
	lock: string,
	where: 'template' | 'style' | 'local' = 'template',
	replacement = cell(lock, 1, 'AND(1,1)'),
	localPins = pins,
	mastered = true,
) {
	const style = locks
		.map((name) => (where === 'style' && name === lock ? replacement : cell(name, 0)))
		.join('');
	return fixture({
		document: `<StyleSheets><StyleSheet ID="0">${style}</StyleSheet></StyleSheets>`,
		masters: [
			{
				id: '7',
				shapes: shape(
					'8',
					pins + dimensions + rectangle + (where === 'template' ? replacement : ''),
					'Type="Shape"',
				),
			},
		],
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', localPins + (mastered ? '' : dimensions + rectangle) + (where === 'local' ? replacement : ''), mastered ? 'Type="Shape" Master="7"' : 'Type="Shape"')}</Shapes>`,
			},
		],
	});
}
async function refuse(bytes: Uint8Array, commands: VisioEdit[] = [move]) {
	const snapshot = bytes.slice();
	await expect(editVsdx(bytes, commands)).rejects.toBeInstanceOf(VisioPackageError);
	expect(bytes).toEqual(snapshot);
}
describe('operation-specific proven master movement protection', () => {
	it('validates template-only matching ancestry and preserves unrelated master admission', async () => {
		for (const moved of ['1', '2']) {
			const parent = moved === '1' ? 1 : 0,
				child = 1;
			const bytes = await fixture({
				document: `<StyleSheets><StyleSheet ID="0">${locks.map((name) => cell(name, name === 'LockAspect' ? parent : 0)).join('')}</StyleSheet><StyleSheet ID="1" LineStyle="0" FillStyle="0" TextStyle="0">${cell('LockAspect', child, 'Inh')}</StyleSheet></StyleSheets>`,
				masters: [
					{
						id: '7',
						shapes: shape(
							'8',
							pins + dimensions + rectangle,
							'Type="Shape" LineStyle="1" FillStyle="1" TextStyle="1"',
						),
					},
				],
				pages: [
					{
						id: '0',
						contents: `<Shapes>${shape('1', pins, 'Type="Shape" Master="7"')}${shape('2', pins + dimensions + rectangle, 'Type="Shape"')}</Shapes>`,
					},
				],
			});
			const result = await editVsdx(bytes, [{ ...move, shapeId: moved }]);
			const before = await JSZip.loadAsync(bytes),
				after = await JSZip.loadAsync(result.bytes);
			expect(await after.file('visio/document.xml')!.async('uint8array')).toEqual(
				await before.file('visio/document.xml')!.async('uint8array'),
			);
			expect(await after.file('visio/masters/master1.xml')!.async('uint8array')).toEqual(
				await before.file('visio/masters/master1.xml')!.async('uint8array'),
			);
		}
	});
	it.each([
		'<Cell N="LockAspect" V="0" F="Inh"/>',
		'<Cell N="LockAspect" V="2" F="Inh"/>',
		'<Cell N="LockAspect" V="1" F="Inh" E="#VALUE!"/>',
		'<Cell N="LockAspect" V="1" F="Inh" U="DA"/>',
	])(
		'validates template-selected style ancestry even when instance selects default: %s',
		async (delegated) => {
			const bytes = await fixture({
				document: `<StyleSheets><StyleSheet ID="0">${locks.map((name) => cell(name, name === 'LockAspect' ? 1 : 0)).join('')}</StyleSheet><StyleSheet ID="1" LineStyle="0" FillStyle="0" TextStyle="0">${delegated}</StyleSheet></StyleSheets>`,
				masters: [
					{
						id: '7',
						shapes: shape(
							'8',
							pins + dimensions + rectangle,
							'Type="Shape" LineStyle="1" FillStyle="1" TextStyle="1"',
						),
					},
				],
				pages: [
					{ id: '0', contents: `<Shapes>${shape('1', pins, 'Type="Shape" Master="7"')}</Shapes>` },
				],
			});
			await refuse(bytes);
		},
	);
	it.each([
		[0, 1],
		[1, 0],
		[1, 1],
	])('checks delegated style lock cache %s against parent %s', async (child, parent) => {
		const bytes = await fixture({
			document: `<DocumentSheet LineStyle="1" FillStyle="1" TextStyle="1"/><StyleSheets><StyleSheet ID="0">${locks.map((name) => cell(name, name === 'LockAspect' ? parent : 0)).join('')}</StyleSheet><StyleSheet ID="1" LineStyle="0" FillStyle="0" TextStyle="0">${cell('LockAspect', child, 'Inh')}</StyleSheet></StyleSheets>`,
			masters: [{ id: '7', shapes: shape('8', pins + dimensions + rectangle, 'Type="Shape"') }],
			pages: [
				{ id: '0', contents: `<Shapes>${shape('1', pins, 'Type="Shape" Master="7"')}</Shapes>` },
			],
		});
		if (child !== parent) await refuse(bytes);
		else {
			const result = await editVsdx(bytes, [move]);
			const before = await JSZip.loadAsync(bytes),
				after = await JSZip.loadAsync(result.bytes);
			expect(await after.file('visio/document.xml')!.async('uint8array')).toEqual(
				await before.file('visio/document.xml')!.async('uint8array'),
			);
			expect((await parseVsdx(result.bytes)).pages[0]!.shapes[0]!.transform.slice(4)).toEqual([
				3.5, 4,
			]);
		}
	});
	it.each(
		independent.flatMap((lock) =>
			(['template', 'style', 'local'] as const).map((where) => [lock, where] as const),
		),
	)('preserves active %s in %s while moving only local pins', async (lock, where) => {
		const bytes = await source(lock, where),
			snapshot = bytes.slice();
		const result = await editVsdx(bytes, [move]);
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(result.bytes);
		for (const [name, part] of Object.entries(before.files))
			if (!part.dir && name !== 'visio/pages/page1.xml')
				expect(await after.file(name)!.async('uint8array')).toEqual(await part.async('uint8array'));
		if (where === 'local')
			expect(await after.file('visio/pages/page1.xml')!.async('string')).toContain(
				cell(lock, 1, 'AND(1,1)'),
			);
		const scene = (await parseVsdx(result.bytes)).pages[0]!.shapes[0]!;
		expect([scene.width, scene.height, ...scene.transform.slice(4)]).toEqual([3, 4, 3.5, 4]);
		expect(bytes).toEqual(snapshot);
	});
	it.each(
		['LockMoveX', 'LockMoveY'].flatMap((lock) =>
			(['template', 'style', 'local'] as const).map((where) => [lock, where] as const),
		),
	)('still refuses active %s in %s', async (lock, where) => {
		await refuse(await source(lock, where));
	});
	it.each(independent)('refuses stale active %s formula caches', async (lock) => {
		await refuse(await source(lock, 'template', cell(lock, 1, '0')));
		await refuse(await source(lock, 'style', cell(lock, 0, '1')));
	});
	it.each([
		'<Cell N="LockAspect" V="2"/>',
		'<Cell N="LockAspect" V="0.5"/>',
		'<Cell N="LockAspect" V="1" U="DL"/>',
		'<Cell N="LockAspect" V="1" E="#VALUE!"/>',
	])('refuses malformed boolean protection %s', async (replacement) => {
		await refuse(await source('LockAspect', 'template', replacement));
	});
	it('refuses unknown/dynamic and referenced active-lock formulas', async () => {
		for (const formula of ['UNKNOWN(1)', 'INDIRECT(&quot;User.Lock&quot;)', 'Sheet.1!User.Lock'])
			await refuse(await source('LockAspect', 'template', cell('LockAspect', 1, formula)));
	});
	it('preserves GUARD on a read-only aspect lock but refuses guarded PinX', async () => {
		const bytes = await source('LockAspect', 'template', cell('LockAspect', 1, 'GUARD(1)'));
		const result = await editVsdx(bytes, [move]);
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(result.bytes);
		expect(await after.file('visio/masters/master1.xml')!.async('uint8array')).toEqual(
			await before.file('visio/masters/master1.xml')!.async('uint8array'),
		);
		await refuse(
			await source(
				'LockAspect',
				'template',
				cell('LockAspect', 1),
				cell('PinX', 2, 'GUARD(2)') + cell('PinY', 3),
			),
		);
	});
	it('retains non-master style protection policy without a master move certificate', async () => {
		await refuse(await source('LockAspect', 'style', cell('LockAspect', 1), pins, false));
	});
	it('keeps mastered resize/delete batches atomic despite an admitted aspect-locked move', async () => {
		for (const edit of [
			{ type: 'resize-shape' as const, pageId: '0', shapeId: '1', width: 5, height: 5 },
			{ type: 'delete-shape' as const, pageId: '0', shapeId: '1' },
		])
			await refuse(await source('LockAspect'), [move, edit]);
	});
});
