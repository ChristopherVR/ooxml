import type { PptxSmartArtData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { setSmartArtInspectorColor, smartArtInspectorColors } from '../smartart-inspector-colors';
import { buildRibbonGallery } from './gallery-registry';
import { smartArtColorPatch } from './smartart-color-patch';
import { SMARTART_COLORS_GALLERY } from './smartart-colors-gallery';

describe('smartArt gallery color edits', () => {
	it('updates node lists and cached colors while preserving connector colors and geometry', () => {
		const data: PptxSmartArtData = {
			nodes: [{ id: 'n1', text: 'One' }],
			colorTransform: {
				fillColors: ['#AAAAAA'],
				lineColors: [],
				roleColors: {
					node1: { fill: ['#AAAAAA'], line: ['#000000'] },
					parChTrans1D2: { fill: ['#123456'], line: ['#654321'] },
				},
			},
			drawingShapes: [
				{
					id: 'reflow-list-n1',
					x: 12,
					y: 24,
					width: 70,
					height: 30,
					text: 'One',
					fillColor: '#AAAAAA',
				},
			],
		};
		const patch = smartArtColorPatch(data, ['#E97132']);
		expect(patch.colorTransform?.roleColors?.node1).toStrictEqual({
			fill: ['#E97132'],
			line: ['#000000'],
		});
		expect(patch.colorTransform?.roleColors?.parChTrans1D2).toStrictEqual(
			data.colorTransform?.roleColors?.parChTrans1D2,
		);
		expect(patch.drawingShapes?.[0]).toMatchObject({
			x: 12,
			y: 24,
			width: 70,
			height: 30,
			fillColor: '#E97132',
		});
		expect(patch.drawingDirty).toBeTruthy();
		expect(data.drawingShapes?.[0].fillColor).toBe('#AAAAAA');
	});

	it('recognizes a saved palette from its parsed colors', () => {
		const descriptor = buildRibbonGallery('smartArtColors', {
			element: {
				id: 'sa',
				type: 'smartArt',
				x: 0,
				y: 0,
				width: 100,
				height: 100,
				smartArtData: { nodes: [], colorTransform: { fillColors: ['#E97132'], lineColors: [] } },
			},
		});
		expect(
			descriptor.sections
				.flatMap((section) => section.items)
				.filter((item) => item.applied)
				.map((item) => item.id),
		).toStrictEqual(['accent2_2']);
	});

	it('recognizes a chosen palette while its definition is still loading', () => {
		const descriptor = buildRibbonGallery('smartArtColors', {
			element: {
				id: 'sa',
				type: 'smartArt',
				x: 0,
				y: 0,
				width: 100,
				height: 100,
				smartArtData: {
					nodes: [],
					colorScheme: 'monochromatic2',
					colorTransformDirty: true,
					colorTransform: {
						uniqueId: 'urn:pptx-viewer/colors/colorful1',
						fillColors: ['#E97132'],
						lineColors: [],
					},
				},
			},
		});
		expect(
			descriptor.sections.flatMap(({ items }) => items).find(({ applied }) => applied)?.id,
		).toBe('accent2_2');
	});

	it('shares the Accent 4 definition with the inspector before the XML is loaded', () => {
		const result = SMARTART_COLORS_GALLERY.apply('accent4_2', {
			element: {
				id: 'sa',
				type: 'smartArt',
				x: 0,
				y: 0,
				width: 100,
				height: 100,
				smartArtData: { nodes: [{ id: 'n1', text: 'One' }] },
			},
		});
		if (
			result?.kind !== 'element' ||
			!('smartArtData' in result.patch) ||
			!result.patch.smartArtData
		)
			throw new Error('Missing SmartArt patch');
		const data = result.patch.smartArtData;
		const t = (key: string, params?: Record<string, string | number>) =>
			`${key} ${params?.accent ?? ''}`;
		const colors = smartArtInspectorColors({ ...data, builtinColorsXml: undefined }, t);
		expect(colors.value).toBe('accent4_2');
		expect(colors.options.find(({ value }) => value === colors.value)?.label).toBe(
			'pptx.gallery.smartArtColors.coloredFill 4',
		);
		const next = setSmartArtInspectorColor(data, 'monochromatic1');
		expect(smartArtInspectorColors(next, t).value).toBe('monochromatic1');
		expect(next.colorTransform?.fillColors).not.toEqual(data.colorTransform?.fillColors);
	});
});
