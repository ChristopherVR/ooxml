import { previewDiagramLayout } from '../../../diagram/layout-preview';
import { resolveDiagramDrawingShape } from '../../../diagram/drawing-resolve';
import { resolveDrawingColor } from '../../../drawingml/drawing-color';
import type { PptxSmartArtData, PptxSmartArtDrawingShape } from '../types';

/** Picture-bearing layouts use the same placeholder-aware engine as Word and Excel. */
export function reflowSmartArtPictures(
	data: PptxSmartArtData,
	width: number,
	height: number,
): PptxSmartArtDrawingShape[] | undefined {
	if (!data.pictureLayoutSource?.pictures?.length) return undefined;
	// A category switch deliberately drops the old definition. Never replay it as the new layout.
	if (data.layoutDirty && !data.layoutDefinition?.rawXmlText) return undefined;
	const source = {
		...data.pictureLayoutSource,
		...(data.layoutDefinition?.rawXmlText ? { layoutXml: data.layoutDefinition.rawXmlText } : {}),
	};
	const preview = previewDiagramLayout({
		extentEmu: { width: width * 9525, height: height * 9525 },
		nodes: data.nodes.map((node) => ({
			id: node.id,
			text: node.text,
			parentId: node.parentId,
			...(node.nodeType === 'asst' ? { nodeType: 'asst' as const } : {}),
		})),
		source,
		minorFont: data.themeMinorFont,
		scheme: data.pictureScheme,
	});
	if (preview.mode !== 'interpreted' || !preview.drawing) return undefined;
	return preview.drawing.shapes.map((shape) => {
		const view = resolveDiagramDrawingShape(shape, {
			resolveColor: (color) =>
				resolveDrawingColor(color, { scheme: (name) => data.pictureScheme?.[name] }),
		});
		if (shape.sourceStyleLabel) view.styleLabel = shape.sourceStyleLabel;
		if (shape.sourceNodeId) view.id = `sa-interp-${shape.sourceNodeId}`;
		if (view.pictureFill?.relId) {
			const part = source.pictureParts?.[view.pictureFill.relId];
			if (part && data.pictureUrls?.[part]) view.fillImageUrl = data.pictureUrls[part];
		}
		return view;
	});
}
