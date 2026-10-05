import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit.js';
import { cell, fixture, rectangle, shape } from './test-fixtures.js';

const target = (extra = '') =>
	shape(
		'1',
		cell('PinX', 2) + cell('PinY', 3) + cell('Width', 2) + cell('Height', 1) + rectangle + extra,
	);
const move = { type: 'move-shape' as const, pageId: '0', shapeId: '1', x: 5, y: 6 };
const resize = { type: 'resize-shape' as const, pageId: '0', shapeId: '1', width: 4, height: 2 };

describe('independent geometry safety review', () => {
	it('refuses moves and resizes when glue caches cannot be proven', async () => {
		const line = shape(
			'2',
			cell('OneD', 1) + cell('BeginX', 2) + cell('BeginY', 3) + cell('EndX', 4) + cell('EndY', 3),
		);
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${target()}${line}</Shapes><Connects><Connect FromSheet="2" ToSheet="1" FromCell="BeginX" ToCell="PinX"/></Connects>`,
				},
			],
		});
		for (const command of [move, resize])
			await expect(editVsdx(bytes, [command])).rejects.toThrow();
	});
	it('refuses dependencies hidden in pages.xml PageSheet metadata', async () => {
		const bytes = await fixture({ pages: [{ id: '0', contents: `<Shapes>${target()}</Shapes>` }] });
		const zip = await JSZip.loadAsync(bytes);
		const path = 'visio/pages/pages.xml';
		zip.file(
			path,
			(await zip.file(path)!.async('string')).replace(
				'<Cell N="PageWidth" V="8.5"/>',
				'<Cell N="PageWidth" V="2" F="Sheet.1!Width"/>',
			),
		);
		await expect(
			editVsdx(await zip.generateAsync({ type: 'uint8array' }), [resize]),
		).rejects.toThrow();
	});
	it('refuses unknown dynamic dependencies in DocumentSheet', async () => {
		const bytes = await fixture({
			document:
				'<DocumentSheet><Section N="User"><Row N="target"><Cell N="Value" V="2" F="EVALCELL(&quot;Pages[Page 1]!Sheet.1!Width&quot;)"/></Row></Section></DocumentSheet>',
			pages: [{ id: '0', contents: `<Shapes>${target()}</Shapes>` }],
		});
		await expect(editVsdx(bytes, [resize])).rejects.toThrow();
	});
	it('preserves uninstantiated master definitions with separate formula scope', async () => {
		const bytes = await fixture({
			masters: [
				{
					id: '7',
					shapes: shape(
						'10',
						'<Cell N="PinX" V="2" F="EVALCELL(&quot;Pages[Page 1]!Sheet.1!PinX&quot;)"/>',
					),
				},
			],
			pages: [{ id: '0', contents: `<Shapes>${target()}</Shapes>` }],
		});
		const saved = await editVsdx(bytes, [move]);
		const before = await JSZip.loadAsync(bytes),
			after = await JSZip.loadAsync(saved.bytes);
		expect(await after.file('visio/masters/master1.xml')!.async('uint8array')).toEqual(
			await before.file('visio/masters/master1.xml')!.async('uint8array'),
		);
	});
	it('refuses dimensionally incompatible transform caches', async () => {
		for (const [name, unit, command] of [
			['PinX', 'DA', move],
			['Width', 'SEC', resize],
		] as const) {
			const content = target().replace(`N="${name}" V="2"`, `N="${name}" V="2" U="${unit}"`);
			await expect(
				editVsdx(await fixture({ pages: [{ id: '0', contents: `<Shapes>${content}</Shapes>` }] }), [
					command,
				]),
			).rejects.toThrow();
		}
	});
	it('does not allow closure recalculation to contradict requested move coordinates', async () => {
		const content = target()
			.replace('N="PinX" V="2"', 'N="PinX" V="2" F="PinY"')
			.replace('N="PinY" V="3"', 'N="PinY" V="2"');
		await expect(
			editVsdx(await fixture({ pages: [{ id: '0', contents: `<Shapes>${content}</Shapes>` }] }), [
				{ ...move, x: 2 },
			]).then(() => 'saved'),
		).rejects.toThrow();
	});
	it('does not allow closure recalculation to contradict requested dimensions', async () => {
		const content = target().replace('N="Width" V="2"', 'N="Width" V="2" F="Height*2"');
		await expect(
			editVsdx(await fixture({ pages: [{ id: '0', contents: `<Shapes>${content}</Shapes>` }] }), [
				{ ...resize, width: 2, height: 4 },
			]).then(() => 'saved'),
		).rejects.toThrow();
	});
	it('recalculates or refuses an affected implicit text-width dependency', async () => {
		const bytes = await fixture({
			pages: [
				{ id: '0', contents: `<Shapes>${target(cell('TxtPinX', 1, 'TxtWidth*0.5'))}</Shapes>` },
			],
		});
		let result;
		try {
			result = await editVsdx(bytes, [resize]);
		} catch {
			return;
		}
		const zip = await JSZip.loadAsync(result.bytes);
		const xml = await zip.file('visio/pages/page1.xml')!.async('string');
		expect(xml).toContain('N="TxtPinX" V="2"');
	});
	it('refuses non-page dependencies on a transitively affected shape', async () => {
		const dependent = shape(
			'2',
			cell('PinX', 5) +
				cell('PinY', 3) +
				cell('Width', 2, 'Sheet.1!Width') +
				cell('Height', 1) +
				rectangle,
		);
		const bytes = await fixture({
			pages: [{ id: '0', contents: `<Shapes>${target()}${dependent}</Shapes>` }],
		});
		const zip = await JSZip.loadAsync(bytes),
			path = 'visio/pages/pages.xml';
		zip.file(
			path,
			(await zip.file(path)!.async('string')).replace(
				'<Cell N="PageWidth" V="8.5"/>',
				'<Cell N="PageWidth" V="2" F="Sheet.2!Width"/>',
			),
		);
		await expect(
			editVsdx(await zip.generateAsync({ type: 'uint8array' }), [resize]).then(() => 'saved'),
		).rejects.toThrow();
	});
});
