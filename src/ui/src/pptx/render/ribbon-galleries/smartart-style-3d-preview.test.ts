import type { PptxElement, PptxSmartArtData } from 'ooxml-core/pptx';
import { DIAGRAM_NATIVE_STYLE_ENTRIES } from 'ooxml-core/diagram';
import { loadBuiltinSmartArtStyleDefinitionXml } from 'ooxml-core/diagram/builtin';
import { describe, expect, it } from 'vitest';

import { preloadSmartArtBuiltinLayouts } from '../smartart-builtin-layouts';
import { resolveSmartArt3DStylePath } from '../smartart-3d-style-path';
import { applyRibbonGalleryItem } from './gallery-registry';
import { smartArtQuickStyle3dPreview } from './smartart-style-3d-preview';

const xml = (id: string) =>
	loadBuiltinSmartArtStyleDefinitionXml(
		'style',
		DIAGRAM_NATIVE_STYLE_ENTRIES.find((entry) => entry.id === id)!.uniqueId,
	)!;

function data(): PptxSmartArtData {
	return {
		nodes: [{ id: 'n1', text: 'One', styleRole: 'node1' }],
		quickStyle: { labels: [{ name: 'node1' }] },
		drawingShapes: [
			{ id: 'n1', x: 0, y: 0, width: 100, height: 60, shapeType: 'rect', styleLabel: 'node1' },
		],
	} as PptxSmartArtData;
}

describe('PowerPoint preview of a 3-D quick-style pick', () => {
	it("gives the cached shapes Polished's bevel without a scene camera", () => {
		const preview = smartArtQuickStyle3dPreview(data(), xml('3d1'));
		expect(preview.quickStyle?.scene3d).toBeUndefined();
		expect(preview.drawingShapes?.[0]?.shape3d?.bevelTopWidth).toBe(120900);
		expect(resolveSmartArt3DStylePath(preview.drawingShapes!)).not.toBe('flat');
	});

	it("puts a Scene style's camera on the diagram and returns to flat on a 2-D pick", () => {
		const scene = smartArtQuickStyle3dPreview(data(), xml('3d5'));
		expect(scene.quickStyle?.scene3d?.cameraPreset).toBe('isometricOffAxis2Left');
		expect(resolveSmartArt3DStylePath(scene.drawingShapes!)).toBe('scene');
		const back = smartArtQuickStyle3dPreview(
			{ ...data(), ...scene } as PptxSmartArtData,
			xml('simple1'),
		);
		expect(back.quickStyle?.scene3d).toBeUndefined();
		expect(resolveSmartArt3DStylePath(back.drawingShapes!)).toBe('flat');
	});

	it('applies the preview through the SmartArt Styles gallery', async () => {
		await preloadSmartArtBuiltinLayouts();
		const element = {
			id: 'sa1',
			type: 'smartArt',
			x: 0,
			y: 0,
			width: 300,
			height: 200,
			smartArtData: data(),
		} as PptxElement;
		const result = applyRibbonGalleryItem('smartArtStyles', '3d9', { element }) as {
			patch: { smartArtData: PptxSmartArtData };
		};
		const picked = result.patch.smartArtData;
		expect(picked.builtinQuickStyleXml).toContain('quickstyle/3d9');
		expect(picked.quickStyle?.scene3d?.cameraPreset).toBeDefined();
		expect(picked.drawingShapes?.[0]?.shape3d).toBeDefined();
	});
});
