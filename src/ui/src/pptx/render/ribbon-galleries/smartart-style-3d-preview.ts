/**
 * The 3D of a picked built-in quick style, applied to the diagram in the editor so a 3-D pick
 * shows its look before the deck is saved and reopened. The picked definition's per-label 3D
 * (bevel, extrusion, material, light rig, text 3D) replaces the quick-style labels' 3D, its
 * spatial camera (the five Scene styles) becomes the diagram camera, and every drawing shape is
 * restyled from its label as PowerPoint writes its cache (`applySmartArtQuickStyle3d`). The lit
 * scene then draws it. Label colours of the extrusion and contour are not resolved here (the
 * renderer's defaults apply), and the colours and effects of the pick still come from the
 * nearest model intensity until the deck is reopened.
 *
 * @module render/ribbon-galleries/smartart-style-3d-preview
 */
import { diagramQuickStyle3d } from 'ooxml-core/diagram';
import type { DrawingText3DValues } from 'ooxml-core/drawingml';
import { applySmartArtQuickStyle3d } from 'ooxml-core/pptx';
import type {
	Pptx3DShape,
	PptxSmartArtData,
	PptxSmartArtDrawingShape,
	PptxSmartArtQuickStyle,
	PptxSmartArtQuickStyleLabel,
	Text3DStyle,
} from 'ooxml-core/pptx';

function shape3d(values: DrawingText3DValues): Pptx3DShape {
	const { z, ...rest } = values;
	return { ...rest, ...(z !== undefined ? { positionZ: z } : {}) };
}

function text3d(values: DrawingText3DValues): Text3DStyle {
	const { z: _z, contourWidth: _contour, ...rest } = values;
	return rest as Text3DStyle;
}

function without3d<T extends { scene3d?: unknown; shape3d?: unknown; text3d?: unknown }>(
	value: T,
): T {
	const { scene3d: _scene, shape3d: _shape, text3d: _text, ...rest } = value;
	return rest as T;
}

/** The quick style and drawing shapes a built-in quick-style definition gives the diagram. */
export function smartArtQuickStyle3dPreview(
	data: PptxSmartArtData,
	quickStyleXml: string,
): Pick<PptxSmartArtData, 'quickStyle' | 'drawingShapes'> {
	const parsed = diagramQuickStyle3d(quickStyleXml);
	const picked = new Map(parsed.labels.map((label) => [label.name, label]));
	const labels: PptxSmartArtQuickStyleLabel[] = (data.quickStyle?.labels ?? []).map(without3d);
	for (const name of picked.keys())
		if (!labels.some((label) => label.name === name)) labels.push({ name });
	const styled = labels.map((label) => {
		const source = picked.get(label.name);
		if (!source) return label;
		return {
			...label,
			...(source.scene ? { scene3d: source.scene } : {}),
			...(source.shape ? { shape3d: shape3d(source.shape) } : {}),
			...(source.text ? { text3d: text3d(source.text) } : {}),
		};
	});
	const quickStyle: PptxSmartArtQuickStyle = {
		...without3d({ ...data.quickStyle }),
		labels: styled,
		...(parsed.scene ? { scene3d: parsed.scene } : {}),
	};
	const shapes = data.drawingShapes;
	const drawingShapes: PptxSmartArtDrawingShape[] | undefined = shapes?.length
		? applySmartArtQuickStyle3d(shapes.map(without3d), { nodes: data.nodes, quickStyle })
		: shapes;
	return { quickStyle, ...(drawingShapes ? { drawingShapes } : {}) };
}
