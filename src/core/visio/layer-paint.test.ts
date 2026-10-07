import { describe, expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import JSZip from 'jszip';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { layerPaintSheet } from './layer-paint';
import { readSheet } from './sheet';
import { parseXml } from '../xml/index';
import { cell, fixture, rectangle, row, section, shape, xml } from './test-fixtures';
import reference from './__fixtures__/layer-colors-native.json';
import type { VisioShape } from './model';

export async function layerPaintFixture(): Promise<Uint8Array> {
	return fixture({
		pages: reference.cases.map((item, index) => {
			const name = item.name;
			const multiple = name === 'multiple' || name === 'mixed';
			const alpha = name === 'source-alpha' || name === 'uncolored-alpha';
			const pattern =
				name === 'no-fill'
					? 0
					: name === 'hatch'
						? 2
						: name === 'gradient' || name === 'linear-partial'
							? 25
							: /^linear\d+$/.test(name)
								? Number(name.slice(6))
								: 1;
			const modern = name.startsWith('modern-gradient');
			return {
				id: String(index),
				width: 4,
				height: 3,
				pageCells: section(
					'Layer',
					row(
						0,
						'',
						cell('Color', ['none', 'uncolored-alpha'].includes(name) ? 255 : '#ff0000') +
							cell('ColorTrans', item.layerTransparency),
					) + (multiple ? row(1, '', cell('Color', name === 'mixed' ? 255 : '#0000ff')) : ''),
				),
				contents: `<Shapes>${shape(
					'1',
					cell('Width', 2) +
						cell('Height', 1) +
						cell('PinX', 2) +
						cell('PinY', 1.5) +
						cell('FillPattern', pattern) +
						cell('FillForegnd', '#00ff00') +
						cell('FillBkgnd', '#ffffff') +
						cell('LineColor', '#0000ff') +
						cell('LinePattern', name === 'no-line' ? 0 : 1) +
						cell('LineWeight', 0.03) +
						cell('LayerMember', multiple ? '0;1' : '0') +
						(name === 'source-alpha'
							? cell('FillForegndTrans', 0.3) + cell('LineColorTrans', 0.2)
							: '') +
						section(
							'Character',
							row(0, '', cell('Color', '#00ffff') + (alpha ? cell('ColorTrans', 0.5) : '')),
						) +
						(modern
							? cell('FillGradientEnabled', 1) +
								cell('FillGradientAngle', 0) +
								cell('FillGradientDir', 0) +
								cell('RotateGradientWithShape', 1) +
								cell('UseGroupGradient', 0) +
								section(
									'FillGradient',
									row(
										0,
										'',
										cell('GradientStopColor', '#00ff00') +
											cell('GradientStopPosition', 0) +
											cell('GradientStopColorTrans', 0.2),
									) +
										row(
											1,
											'',
											cell('GradientStopColor', '#0000ff') +
												cell('GradientStopPosition', 1) +
												cell('GradientStopColorTrans', 0.5),
										),
								)
							: '') +
						rectangle +
						'<Text>Layer color</Text>',
				)}</Shapes>`,
			};
		}),
	});
}

export function expectNativeLayerPaint(actual: VisioShape, index: number): void {
	const native = reference.cases[index]!.paint;
	if (native.stroke === 'none') expect(actual.style.linePattern).toBe(0);
	else expect(actual.style.lineColor).toBe(native.stroke);
	// Visio serializes vector opacity to two decimals, character opacity to six.
	expect(Math.abs(actual.style.lineOpacity - native.strokeOpacity)).toBeLessThan(0.0050001);
	expect(actual.text.color).toBe(native.text);
	expect(actual.text.opacity ?? 1).toBeCloseTo(native.textOpacity, 6);
	for (const run of actual.text.runs) {
		expect(run.color).toBe(native.text);
		expect(run.opacity ?? 1).toBeCloseTo(native.textOpacity, 6);
	}
	if (native.gradientStops.length) {
		const gradient = actual.style.fillGradient!;
		expect(gradient?.stops).toEqual(native.gradientStops);
		expect(actual.style.fillOpacity).toBe(1);
		if (gradient.type !== 'linear') throw new Error('Expected a native linear fill.');
		const direction = [gradient.end[0] - gradient.start[0], gradient.end[1] - gradient.start[1]];
		const expected =
			native.gradientAngle === 0
				? [2, 0]
				: native.gradientAngle === 180
					? [-2, 0]
					: native.gradientAngle === 90
						? [0, -1]
						: [0, 1];
		expect(direction).toEqual(expected);
	} else if (reference.cases[index]!.name !== 'hatch') {
		expect(actual.style.fill).toBe(native.fill);
		expect(actual.style.fillGradient).toBeUndefined();
		expect(Math.abs(actual.style.fillOpacity - native.fillOpacity)).toBeLessThan(0.0050001);
	}
}

describe('native single-layer paint context', () => {
	it('matches native colors, transparency, membership and orthogonal legacy gradients', async () => {
		const document = await parseVsdx(await layerPaintFixture());
		for (let index = 0; index < document.pages.length; index++)
			expectNativeLayerPaint(document.pages[index]!.shapes[0]!, index);
		expect(document.diagnostics.some((item) => item.code === 'unsupported-layer-color')).toBe(
			false,
		);
		expect(document.diagnostics.some((item) => item.code === 'unsupported-fill-pattern')).toBe(
			false,
		);
	});
	it('does not mutate source or inherited sheet colors, and supplies a missing character row', () => {
		const sheet = readSheet(
			parseXml(xml('Shape', cell('FillForegnd', '#00ff00') + section('Character', '')))
				.documentElement,
		);
		const layers = new Map([
			[
				'0',
				{
					id: '0',
					name: 'Red',
					visible: true,
					printable: true,
					locked: false,
					color: '#ff0000',
					colorOpacity: 0.6,
				},
			],
		]);
		const painted = layerPaintSheet(sheet, ['0'], layers);
		expect(sheet.cells.get('FillForegnd')!.value).toBe('#00ff00');
		expect(sheet.sections.get('Character:0')!.rows.size).toBe(0);
		expect(painted.sections.get('Character:0')!.rows.get('0')!.cells.get('Color')!.value).toBe(
			'#ff0000',
		);
		expect(layerPaintSheet(sheet, ['0', '1'], layers)).toBe(sheet);
	});
	it('keeps cached source colors intact through move, save and reparse', async () => {
		const source = await layerPaintFixture();
		const saved = await editVsdx(source, [
			{ type: 'move-shape', pageId: '0', shapeId: '1', x: 2.25, y: 1.5 },
		]);
		expectNativeLayerPaint((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!, 0);
		const original = (await JSZip.loadAsync(source)).file('visio/pages/page1.xml')!;
		const changed = (await JSZip.loadAsync(saved.bytes)).file('visio/pages/page1.xml')!;
		for (const value of ['#00ff00', '#0000ff', '#00ffff']) {
			expect(await original.async('string')).toContain(value);
			expect(await changed.async('string')).toContain(value);
		}
	});
	it('reports unverified colored group and text-background paint', async () => {
		const zip = await JSZip.loadAsync(await layerPaintFixture());
		const page = await zip.file('visio/pages/page1.xml')!.async('string');
		zip.file('visio/pages/page1.xml', page.replace('<Shape ID="1"', '<Shape ID="1" Type="Group"'));
		const document = await parseVsdx(await zip.generateAsync({ type: 'uint8array' }));
		expect(
			document.diagnostics.some((item) => item.code === 'unverified-layer-group-text-paint'),
		).toBe(true);
	});
});

const nativeDirectory = process.env.VISIO_NATIVE_LAYER_COLORS_DIR;
it.skipIf(!nativeDirectory)(
	'compares all genuine native pages and writes a core-edited acceptance drawing',
	async () => {
		const bytes = new Uint8Array(await readFile(join(nativeDirectory!, 'layer-colors.vsdx')));
		const document = await parseVsdx(bytes);
		expect(document.pages).toHaveLength(reference.cases.length);
		for (let index = 0; index < document.pages.length; index++)
			expectNativeLayerPaint(document.pages[index]!.shapes[0]!, index);
		const saved = await editVsdx(bytes, [
			{ type: 'move-shape', pageId: reference.cases[0]!.pageId, shapeId: '1', x: 2.25, y: 1.5 },
		]);
		await writeFile(join(nativeDirectory!, 'core-layer-colors.vsdx'), saved.bytes);
		const reopened = await parseVsdx(saved.bytes);
		for (let index = 0; index < document.pages.length; index++)
			expectNativeLayerPaint(reopened.pages[index]!.shapes[0]!, index);
	},
);
