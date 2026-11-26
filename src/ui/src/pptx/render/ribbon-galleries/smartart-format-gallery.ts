import type { PptxSmartArtDrawingShape, PptxSmartArtNode, ShapePptxElement } from 'ooxml-core/pptx';
import type { DiagramNodeStylePatch } from 'ooxml-core/diagram';
import { presetShapeSections } from './crop-to-shape-gallery';
import { applyRibbonGalleryItem } from './gallery-registry';
import type { RibbonGalleryContext, RibbonGalleryDescriptor } from './gallery-types';

type StyleGallery = 'shapeStyles' | 'wordArtStyles';

export function smartArtStyleContext(
	node: PptxSmartArtNode | undefined,
	drawing: PptxSmartArtDrawingShape | undefined,
	themeColorMap: RibbonGalleryContext['themeColorMap'],
): RibbonGalleryContext {
	const element: ShapePptxElement | null =
		node && drawing
			? {
					type: 'shape',
					id: node.id,
					x: drawing.x,
					y: drawing.y,
					width: drawing.width,
					height: drawing.height,
					shapeType: drawing.shapeType ?? 'rect',
					text: node.text,
					textStyle: { color: node.style?.fontColor ?? drawing.fontColor },
					textSegments: drawing.textSegments?.length
						? drawing.textSegments
						: [{ text: node.text, style: { color: node.style?.fontColor ?? drawing.fontColor } }],
					shapeStyle: {
						fillColor: node.style?.fillColor ?? drawing.fillColor,
						strokeColor: node.style?.lineColor ?? drawing.strokeColor,
						strokeWidth: drawing.strokeWidth,
						fillOpacity: drawing.fillOpacity,
					},
				}
			: null;
	return { element, ...(themeColorMap && { themeColorMap }) };
}

export function smartArtShapeGallery(shapeType: string | undefined): RibbonGalleryDescriptor {
	return {
		id: 'cropToShape',
		labelKey: 'pptx.mediaRibbon.changeShape',
		label: 'Change Shape',
		sections: presetShapeSections(shapeType),
		disabled: shapeType === undefined,
	};
}

export function smartArtStylePick(
	id: StyleGallery,
	itemId: string,
	context: RibbonGalleryContext,
): { style: DiagramNodeStylePatch; drawing: Partial<PptxSmartArtDrawingShape> } | undefined {
	const result = applyRibbonGalleryItem(id, itemId, context);
	if (result?.kind !== 'element') return;
	const patch = result.patch as Partial<ShapePptxElement>;
	if (id === 'wordArtStyles') {
		return {
			style: { fontColor: patch.textStyle?.color },
			drawing: {
				fontColor: patch.textStyle?.color,
				textSegments: patch.textSegments?.map((segment) => ({
					...segment,
					style: { ...segment.style, runPropertiesXml: undefined },
				})),
				text3d: patch.textStyle?.text3d,
			},
		};
	}
	const shape = patch.shapeStyle;
	if (!shape) return;
	const fillColor = shape.fillMode === 'none' ? undefined : shape.fillColor;
	const strokeColor = shape.strokeColor === 'transparent' ? 'none' : shape.strokeColor;
	const fontColor = patch.textSegments?.find((segment) => !segment.isParagraphBreak)?.style.color;
	return {
		style: {
			fillColor,
			lineColor: strokeColor === 'none' ? undefined : strokeColor,
			...(fontColor && { fontColor }),
		},
		drawing: {
			fillColor,
			fillNone: shape.fillMode === 'none',
			fillOpacity: shape.fillOpacity,
			fillGradientStops: shape.fillGradientStops,
			fillGradientType: shape.fillGradientType,
			fillGradientAngle: shape.fillGradientAngle,
			fillPatternPreset: undefined,
			fillBlipEmbedId: undefined,
			fillImageUrl: undefined,
			pictureFill: undefined,
			strokeColor,
			strokeWidth: shape.strokeWidth === undefined ? undefined : shape.strokeWidth * 0.75,
			...(fontColor && { fontColor, textSegments: patch.textSegments }),
		},
	};
}
