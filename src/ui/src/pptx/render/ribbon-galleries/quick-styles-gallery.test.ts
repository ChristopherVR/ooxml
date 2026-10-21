import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { SHAPE_QUICK_STYLES } from '../shape-quick-styles';
import { applyRibbonGalleryItem, buildRibbonGallery } from './gallery-registry';
import { quickStyleShapeStyle } from './quick-styles-gallery';

const shape = (shapeStyle?: PptxElement['shapeStyle']): PptxElement =>
	({
		id: 'box',
		type: 'shape',
		x: 0,
		y: 0,
		width: 200,
		height: 100,
		shapeType: 'rect',
		...(shapeStyle ? { shapeStyle } : {}),
	}) as PptxElement;

describe('quick styles gallery', () => {
	it('offers every preset as a tile in rows of six, none marked', () => {
		const gallery = buildRibbonGallery('quickStyles', { element: shape() });
		expect(gallery.disabled).toBeFalsy();
		expect(gallery.sections).toHaveLength(1);
		const [section] = gallery.sections;
		expect(section!.columns).toBe(6);
		expect(section!.items.map((item) => item.labelKey)).toStrictEqual(
			SHAPE_QUICK_STYLES.map((preset) => preset.labelKey),
		);
		expect(section!.items[0]).toMatchObject({ id: 'blueFill', label: 'Blue Fill' });
		expect(section!.items.some((item) => item.applied)).toBeFalsy();
	});

	it('draws each preset with its fill, outline and shadow', () => {
		const items = buildRibbonGallery('quickStyles', { element: shape() }).sections[0]!.items;
		const tile = (id: string) => items.find((item) => item.id === id)!.previewSvg!;
		expect(tile('blueFill')).toMatch(/fill="#4472C4" stroke="#2F5597"/u);
		expect(tile('blueOutline')).toMatch(/fill="none" stroke="#4472C4" stroke-width="2"/u);
		expect(tile('blueShadow')).toMatch(/flood-color="#000000" flood-opacity="0.35"/u);
		expect(tile('blueFill')).not.toMatch(/<filter/u);
	});

	it('merges a pick over the shape style, and is unavailable without a shape style', () => {
		const element = shape({ fillColor: '#123456', shadowColor: '#000000', shadowOpacity: 0.5 });
		expect(applyRibbonGalleryItem('quickStyles', 'orangeOutline', { element })).toStrictEqual({
			kind: 'element',
			elementId: 'box',
			patch: {
				shapeStyle: {
					fillMode: 'none',
					fillColor: 'transparent',
					strokeColor: '#ED7D31',
					strokeWidth: 2,
					shadowColor: '#000000',
					shadowOpacity: 0.5,
				},
			},
		});
		expect(applyRibbonGalleryItem('quickStyles', 'nope', { element })).toBeNull();
		const chart = { id: 'c', type: 'chart', x: 0, y: 0, width: 1, height: 1 } as PptxElement;
		expect(buildRibbonGallery('quickStyles', { element: chart }).disabled).toBeTruthy();
		expect(buildRibbonGallery('quickStyles', { element: null }).disabled).toBeTruthy();
		expect(applyRibbonGalleryItem('quickStyles', 'blueFill', { element: chart })).toBeNull();
	});

	it('names the shape style of a tile for hosts that merge it themselves', () => {
		expect(quickStyleShapeStyle('greenFill')).toBe(SHAPE_QUICK_STYLES[4]!.style);
		expect(quickStyleShapeStyle('Green Fill')).toBeUndefined();
	});
});
