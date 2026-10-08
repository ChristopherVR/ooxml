import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { snapshotFormatting } from './edit-formatting-commands';
import { snapshotEdits } from './ui/edit-commands';
import { shapeFormattingWrites } from './edit-formatting-paint';
import { copySnapshotScene } from './ui/snapshot-scene';
import { assertViewableDocument } from './ui/scene-validation';
import { visioShapeFormattingState } from './ui/shape-formatting';
import { attribute, children } from './sheet';
import { cell, fixture, shape, rectangle } from './test-fixtures';

const target = { type: 'format-shape' as const, pageId: '0', shapeId: '1' };
const source = (extra = '', document = '', attributes = '') =>
	fixture({
		document,
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', 3) + cell('Height', 1) + rectangle + extra + '<Unknown payload="preserved"/><Text>Paint</Text>', attributes)}</Shapes>`,
			},
		],
		edit: (zip) => zip.file('custom/opaque.bin', new Uint8Array([0, 255])),
	});
const read = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;

describe('source paint pattern and transparency edits', () => {
	it.each(Array.from({ length: 24 }, (_, value) => value))(
		'writes built-in line pattern %i without changing color, transparency or fill',
		async (linePattern) => {
			const bytes = await source(
				cell('LineColor', '#123456') +
					cell('LineColorTrans', 0.25) +
					cell('FillForegnd', '#abcdef'),
			);
			const saved = await editVsdx(bytes, [{ ...target, linePattern }]);
			expect((await read(saved.bytes)).style).toMatchObject({
				linePattern,
				lineColor: '#123456',
				lineOpacity: 0.75,
				fill: '#abcdef',
			});
			expect((await editVsdx(saved.bytes, [{ ...target, linePattern }])).bytes).toEqual(
				saved.bytes,
			);
		},
	);
	it.each(Array.from({ length: 25 }, (_, value) => value))(
		'writes classic non-gradient fill pattern %i',
		async (fillPattern) => {
			const saved = await editVsdx(
				await source(
					cell('FillGradientEnabled', 1) +
						cell('FillForegnd', '#ff0000') +
						cell('FillBkgnd', '#0000ff'),
				),
				[{ ...target, fillPattern }],
			);
			const actual = await read(saved.bytes);
			expect(actual.style.fillPatternIndex).toBe(fillPattern);
			expect(actual.style.fill).toBe(fillPattern === 0 ? 'none' : '#ff0000');
			expect(actual.style.fillGradient).toBeUndefined();
			expect(!!actual.style.fillPattern).toBe(fillPattern > 1);
		},
	);
	it.each([
		[0, 0],
		[0.24, 0],
		[0.25, 0.005],
		[0.75, 0.01],
		[12.25, 0.125],
		[12.75, 0.13],
		[99.75, 1],
		[100, 1],
	])('quantizes %s percent to native cache %s, with ties upward', async (percentage, cache) => {
		const saved = await editVsdx(await source(), [
			{ ...target, fillTransparency: percentage!, lineTransparency: percentage! },
		]);
		const actual = await read(saved.bytes);
		expect(actual.style).toMatchObject({
			fillForegroundOpacity: 1 - cache!,
			fillBackgroundOpacity: 1 - cache!,
			lineColorOpacity: 1 - cache!,
		});
		const root = await (await VisioPackage.open(saved.bytes)).readXml('visio/pages/page1.xml');
		const node = children(children(root, 'Shapes')[0], 'Shape')[0]!;
		for (const name of ['FillForegndTrans', 'FillBkgndTrans', 'LineColorTrans'])
			expect(
				attribute(
					children(node, 'Cell').find((item) => attribute(item, 'N') === name),
					'V',
				),
			).toBe(String(cache));
		expect(
			(
				await editVsdx(saved.bytes, [
					{ ...target, fillTransparency: percentage!, lineTransparency: percentage! },
				])
			).bytes,
		).toEqual(saved.bytes);
	});
	it('combines colors, hatches and transparency with one write per cell and retains source payloads', async () => {
		const command = {
			...target,
			fillColor: '#ff0000',
			fillPattern: 3,
			fillBackgroundColor: '#0000ff',
			fillTransparency: 25,
			lineColor: '#123456',
			lineTransparency: 50,
			linePattern: 2,
		};
		const writes = shapeFormattingWrites(command);
		expect(new Set(writes.map((write) => write.name)).size).toBe(writes.length);
		const bytes = await source(
			cell('FillGradientEnabled', 1) +
				'<Section N="FillGradient"><Row IX="0"><Cell N="GradientStopColor" V="#ffffff"/></Row></Section>',
		);
		const saved = await editVsdx(bytes, [command]);
		const actual = await read(saved.bytes);
		expect(actual.style).toMatchObject({
			fillPatternIndex: 3,
			fill: '#ff0000',
			fillBackgroundColor: '#0000ff',
			fillForegroundOpacity: 0.75,
			fillBackgroundOpacity: 0.75,
			fillOpacity: 1,
			lineOpacity: 0.5,
		});
		const before = await VisioPackage.open(bytes),
			after = await VisioPackage.open(saved.bytes);
		for (const path of before.paths())
			if (!saved.changedParts.includes(path))
				expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
		const xml = new TextDecoder().decode(await after.readBytes('visio/pages/page1.xml'));
		expect(xml).toContain('payload="preserved"');
		expect(xml).toContain('N="GradientStopColor" V="#ffffff"');
		expect(xml).toContain('<Text>Paint</Text>');
	});
	it('enables an outline without changing text, fill, local line color or gradient definitions', async () => {
		const saved = await editVsdx(
			await source(cell('LinePattern', 0) + cell('FillPattern', 0) + cell('LineColor', '#123456')),
			[{ ...target, linePattern: 1 }],
		);
		expect(await read(saved.bytes)).toMatchObject({
			text: { plainText: 'Paint' },
			style: { fill: 'none', linePattern: 1, lineColor: '#123456' },
		});
	});
	it('owns numeric/color patches before asynchronous package reads', async () => {
		const command = {
			...target,
			fillPattern: 3,
			fillTransparency: 25,
			fillBackgroundColor: '#0000ff',
		};
		const pending = editVsdx(await source(), [command]);
		command.fillPattern = 24;
		command.fillTransparency = 100;
		command.fillBackgroundColor = '#ff0000';
		expect((await read((await pending).bytes)).style).toMatchObject({
			fillPatternIndex: 3,
			fillBackgroundOpacity: 0.75,
			fillBackgroundColor: '#0000ff',
		});
	});
	it('snapshots new patch fields and strips arbitrary host properties', () => {
		const command = {
			...target,
			fillPattern: 24,
			fillTransparency: 12.25,
			fillBackgroundColor: '#ABCDEF',
			linePattern: 23,
			lineTransparency: 12.75,
			host: {},
		};
		const expected = {
			...target,
			fillPattern: 24,
			fillTransparency: 12.5,
			fillBackgroundColor: '#abcdef',
			linePattern: 23,
			lineTransparency: 13,
		};
		expect(snapshotFormatting(command)).toEqual(expected);
		expect(snapshotEdits([command])).toEqual([expected]);
	});
});

describe('optional source paint scene state', () => {
	it('preserves source metadata in snapshots and aggregates mixed foreground/background values', async () => {
		const saved = await editVsdx(await source(), [
			{
				...target,
				fillPattern: 3,
				fillBackgroundColor: '#0000ff',
				fillTransparency: 25,
				lineTransparency: 50,
			},
		]);
		const model = await parseVsdx(saved.bytes),
			actual = model.pages[0]!.shapes[0]!;
		const copy = copySnapshotScene(model);
		assertViewableDocument(copy);
		expect(copy.pages[0]!.shapes[0]!.style).toEqual(actual.style);
		expect(visioShapeFormattingState([actual])).toMatchObject({
			fillPatternIndex: 3,
			fillBackgroundColor: '#0000ff',
			fillTransparency: 25,
			lineTransparency: 50,
		});
		const different = structuredClone(actual);
		different.style.fillBackgroundOpacity = 0.5;
		expect(visioShapeFormattingState([actual, different])).toMatchObject({
			fillForegroundTransparency: 25,
			fillBackgroundTransparency: undefined,
			fillTransparency: undefined,
		});
	});
	it.each(['fillForegroundOpacity', 'fillBackgroundOpacity', 'lineColorOpacity'] as const)(
		'rejects malformed optional opacity %s',
		async (name) => {
			const model = await parseVsdx(await source());
			model.pages[0]!.shapes[0]!.style[name] = NaN;
			expect(() => assertViewableDocument(model)).toThrow();
		},
	);
	it('keeps old host scenes valid and does not infer hatch/gradient source opacity from render multipliers', async () => {
		const model = await parseVsdx(await source(cell('FillPattern', 3)));
		const actual = model.pages[0]!.shapes[0]!;
		delete actual.style.fillPatternIndex;
		delete actual.style.fillForegroundOpacity;
		delete actual.style.fillBackgroundOpacity;
		delete actual.style.fillBackgroundColor;
		delete actual.style.lineColorOpacity;
		assertViewableDocument(model);
		expect(visioShapeFormattingState([actual]).fillTransparency).toBeUndefined();
	});
});
