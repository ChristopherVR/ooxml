import { describe, expect, it } from 'vitest';
import {
	BUILTIN_SMARTART_COLORS_XML,
	builtinLayoutIdForPreset,
	loadBuiltinSmartArtLayoutXml,
	loadBuiltinSmartArtStyleDefinitionXml,
} from './builtin';
import { previewDiagramLayout } from './layout-preview';
import { DIAGRAM_NATIVE_STYLE_ENTRIES } from './catalogue-native';
import { parseDiagramDrawing } from './drawing';
import { diagramQuickStyle3d, diagramShape3dView, parseDiagramShape3dXml } from './scene-3d';
import { projectDiagramScenePoint, resolveDiagramSceneCamera } from './scene-camera';

const style = (id: string) =>
	loadBuiltinSmartArtStyleDefinitionXml(
		'style',
		DIAGRAM_NATIVE_STYLE_ENTRIES.find((entry) => entry.id === id)!.uniqueId,
	)!;

describe('SmartArt 3D values', () => {
	it("reads a cached shape's bevel, extrusion and material", () => {
		const drawing = parseDiagramDrawing(
			`<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp modelId="{00000000-0000-0000-0000-000000000001}"><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d><a:sp3d extrusionH="38100" prstMaterial="plastic"><a:bevelT w="120900" h="88900"/></a:sp3d></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`,
		);
		const view = diagramShape3dView(drawing.shapes[0]!);
		expect(view?.scene).toMatchObject({
			cameraPreset: 'orthographicFront',
			lightRigType: 'threePt',
		});
		expect(view?.shape).toMatchObject({
			extrusionHeight: 38100,
			presetMaterial: 'plastic',
			bevelTopType: 'circle',
			bevelTopWidth: 120900,
		});
		expect(parseDiagramShape3dXml('<a:sp3d/>')).toBeUndefined();
	});

	it("keeps a Scene style's spatial camera and every label's 3D", () => {
		const scene = diagramQuickStyle3d(style('3d9'));
		expect(scene.scene?.cameraPreset).toBe(
			DIAGRAM_NATIVE_STYLE_ENTRIES.find((entry) => entry.id === '3d9')!.preview.camera,
		);
		expect(scene.labels.find((label) => label.name === 'node1')?.shape).toBeDefined();
		const polished = diagramQuickStyle3d(style('3d1'));
		expect(polished.scene).toBeUndefined();
		expect(polished.labels.find((label) => label.name === 'node1')?.shape?.bevelTopWidth).toBe(
			120900,
		);
		expect(diagramQuickStyle3d(style('simple1')).labels.every((label) => !label.shape)).toBe(true);
	});

	it('projects through the fitted scene cameras', () => {
		expect(resolveDiagramSceneCamera({ cameraPreset: 'orthographicFront' })).toBeUndefined();
		const left = resolveDiagramSceneCamera({ cameraPreset: 'perspectiveLeft' })!;
		const centre = { x: 300, y: 200 };
		const near = projectDiagramScenePoint(left, centre, { x: 100, y: 200, z: 0 });
		const far = projectDiagramScenePoint(left, centre, { x: 500, y: 200, z: 0 });
		// Turned to the left: one side recedes, the other comes nearer and grows.
		expect(Math.abs(near.x - centre.x)).not.toBeCloseTo(Math.abs(far.x - centre.x), 0);
		const iso = resolveDiagramSceneCamera({ cameraPreset: 'isometricOffAxis2Left' })!;
		expect(iso.projection).toBe('orthographic');
		// A parallel projection shears a horizontal edge: its right end rises or falls.
		const right = projectDiagramScenePoint(iso, centre, { x: 500, y: 200, z: 0 });
		expect(right.y).not.toBeCloseTo(200, 1);
		expect(right.x - centre.x).toBeLessThan(200);
	});

	it('carries the quick-style camera onto an interpreted drawing', () => {
		const layoutId = builtinLayoutIdForPreset('basicBlockList')!;
		const preview = previewDiagramLayout({
			source: {
				layoutXml: loadBuiltinSmartArtLayoutXml(layoutId),
				colorsXml: BUILTIN_SMARTART_COLORS_XML,
				quickStyleXml: style('3d5'),
			},
			layoutId,
			extentEmu: { width: 600 * 9525, height: 350 * 9525 },
			nodes: [
				{ id: 'a', text: 'Alpha' },
				{ id: 'b', text: 'Beta' },
			],
		});
		expect(preview.drawing?.scene3d?.cameraPreset).toBe('isometricOffAxis2Left');
		expect(preview.drawing?.shapes.some((shape) => diagramShape3dView(shape)?.shape)).toBe(true);
	});
});
