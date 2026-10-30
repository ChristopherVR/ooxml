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
		// A generated layout carries its caption on the flat `text` summary, not on a cached
		// `dsp:txBody`, so the resolved view has no text until it is copied over.
		if (view.text === undefined && shape.text?.text) {
			view.text = shape.text.text;
			const run = shape.text.paragraphs.flatMap((paragraph) => paragraph.runs)[0];
			if (run?.sizePt) view.fontSize = (run.sizePt * 96) / 72;
			if (run?.typeface) view.fontFamily = run.typeface;
			if (run?.bold !== undefined) view.fontWeight = run.bold ? 700 : 400;
			if (run?.italic !== undefined) view.fontStyle = run.italic ? 'italic' : 'normal';
			const color = run?.color
				? resolveDrawingColor(run.color, { scheme: (name) => data.pictureScheme?.[name] })
				: undefined;
			if (color) view.fontColor = color.hex;
			if (shape.text.anchor) view.textVerticalAnchor = shape.text.anchor;
		}
		if (shape.sourceStyleLabel) view.styleLabel = shape.sourceStyleLabel;
		if (shape.sourceNodeId) view.id = `sa-interp-${shape.sourceNodeId}`;
		if (view.pictureFill?.relId) {
			const part = source.pictureParts?.[view.pictureFill.relId];
			if (part && data.pictureUrls?.[part]) view.fillImageUrl = data.pictureUrls[part];
			if (part && data.pictureSizes?.[part]) {
				view.fillImageWidth = data.pictureSizes[part].width;
				view.fillImageHeight = data.pictureSizes[part].height;
			}
		}
		return view;
	});
}
