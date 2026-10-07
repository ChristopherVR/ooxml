import type { PptxElement } from 'ooxml-core/pptx';
import { beforeAll, describe, expect, it } from 'vitest';

import { preloadSmartArtBuiltinLayouts } from '../smartart-builtin-layouts';
import { applyRibbonGalleryItem, buildRibbonGallery } from './gallery-registry';

/** A SmartArt whose cached drawing is present, so a layout change has shapes to rebuild. */
const smartArt = {
	id: 'sa1',
	type: 'smartArt',
	x: 0,
	y: 0,
	width: 867,
	height: 533,
	smartArtData: {
		layout: 'basicBlockList',
		resolvedLayoutType: 'list',
		nodes: [
			{ id: 'n1', text: 'Node One' },
			{ id: 'n2', text: 'Node Two', parentId: 'n1' },
			{ id: 'n3', text: 'Node Three', parentId: 'n1' },
		],
		connections: [],
		drawingShapes: [{ id: 'old', shapeType: 'rect', x: 0, y: 0, width: 10, height: 10 }],
	},
} as unknown as PptxElement;
const ctx = { element: smartArt };

type Patch = { smartArtData: Record<string, unknown> & { drawingShapes?: { id: string }[] } };

describe('named SmartArt layouts in the Layouts gallery', () => {
	beforeAll(async () => {
		await preloadSmartArtBuiltinLayouts();
	});

	it("lists PowerPoint's built-in layouts under their categories once loaded", () => {
		const sections = buildRibbonGallery('smartArtLayouts', ctx).sections;
		const titles = sections.flatMap((section) => section.items.map((item) => item.label));
		expect(titles).toEqual(
			expect.arrayContaining(['Circle Relationship', 'Basic Timeline', 'Basic Bending Process']),
		);
		expect(sections.map((section) => section.id)).toEqual(
			expect.arrayContaining(['layouts', 'named-relationship', 'named-timeline']),
		);
	});

	it("applying one installs that layout's own definition and rebuilds the drawing", () => {
		const sections = buildRibbonGallery('smartArtLayouts', ctx).sections;
		const item = sections.flatMap((s) => s.items).find((i) => i.label === 'Circle Relationship');
		const result = applyRibbonGalleryItem('smartArtLayouts', item!.id, ctx);
		expect(result?.kind).toBe('element');
		const data = (result as { patch: Patch }).patch.smartArtData;
		expect(data.layoutDefinition).toMatchObject({ uniqueId: item!.id });
		expect(data.resolvedLayoutType).toBe('relationship');
		// the engine rebuilt the shapes: none of them is the stale placeholder
		expect(data.drawingShapes?.length).toBeGreaterThan(0);
		expect(data.drawingShapes?.some((shape) => shape.id === 'old')).toBe(false);
	});

	it('a category switch now applies that family’s real layout', () => {
		const result = applyRibbonGalleryItem('smartArtLayouts', 'timeline', ctx);
		const data = (result as { patch: Patch }).patch.smartArtData;
		expect(data.resolvedLayoutType).toBe('timeline');
		expect(data.layoutDefinition).toBeDefined();
	});

	it('an unknown layout id applies nothing', () => {
		expect(applyRibbonGalleryItem('smartArtLayouts', 'urn:nope', ctx)).toBeNull();
	});
});
