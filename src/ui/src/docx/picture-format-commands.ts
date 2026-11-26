// Picture Format commands: Wrap Text and Crop on the selected picture. The model changes come from
// the core (`setDrawingWrap`, `setPictureCrop`); this module maps them onto the image node's
// attributes in one undoable step.
import { closeHistory } from 'prosemirror-history';
import { NodeSelection, type EditorState } from 'prosemirror-state';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import {
	drawingWrapOf,
	setDrawingWrap,
	setPictureCrop,
	type DrawingWrap,
	type FloatOrigin,
	type InlineImage,
	type PicturePlacement,
} from 'ooxml-core/docx';
import type { CropInsets } from 'ooxml-core/geometry';
import { wordYjsPluginKey } from 'ooxml-core/docx/ui';
import { schema } from './schema';
import { patchPictureEffects, type PictureEffectPatch } from 'ooxml-core/drawingml';

const parse = <T>(value: unknown): T | undefined => {
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value) as T;
	} catch {
		return undefined;
	}
};

/** Whether `node` is a real picture (not a chart, SmartArt, shape, text box or placeholder). */
export function isPictureNode(node: ProseMirrorNode | null | undefined): node is ProseMirrorNode {
	return Boolean(
		node &&
		node.type === schema.nodes.image &&
		node.attrs.partName &&
		!node.attrs.unsupported &&
		!node.attrs.shape &&
		!node.attrs.chart &&
		!node.attrs.diagram &&
		!node.attrs.payload &&
		!node.attrs.textBoxEditable &&
		// A legacy VML picture (`w:pict`) has no DrawingML fill or wrap to write; one inserted in
		// this session has no relationship yet and is written as DrawingML.
		(node.attrs.pictureFill || !node.attrs.relId),
	);
}

/** The selected picture's position, if the selection is on one. */
export function selectedPicturePos(state: EditorState): number | undefined {
	const { selection } = state;
	return selection instanceof NodeSelection && isPictureNode(selection.node)
		? selection.from
		: undefined;
}

/** The model fields of a picture node that cropping and wrapping read and write. */
export function pictureOf(node: ProseMirrorNode): InlineImage {
	const placement = parse<PicturePlacement>(node.attrs.placement);
	return {
		relId: String(node.attrs.relId),
		partName: String(node.attrs.partName),
		contentType: String(node.attrs.contentType),
		widthPx: Number(node.attrs.widthPx),
		heightPx: Number(node.attrs.heightPx),
		// A picture inserted in this session has no source fill yet; it is a plain picture fill.
		pictureFill: parse<NonNullable<InlineImage['pictureFill']>>(node.attrs.pictureFill) ?? {
			kind: 'picture',
		},
		...(node.attrs.anchored ? { anchored: true } : {}),
		...(placement ? { placement } : {}),
	};
}

/** The Wrap Text choice that describes a picture node. */
export function pictureWrapOf(node: ProseMirrorNode): DrawingWrap {
	return drawingWrapOf(Boolean(node.attrs.anchored), parse(node.attrs.placement));
}

/** Writes `image`'s size, wrapping and fill onto the node at `pos` as one undoable step. */
function applyPicture(view: EditorView, pos: number, image: InlineImage): boolean {
	const node = view.state.doc.nodeAt(pos);
	if (!view.editable || !isPictureNode(node)) return false;
	const history = wordYjsPluginKey.getState(view.state);
	history?.stopCapturing();
	const tr = view.state.tr.setNodeMarkup(pos, undefined, {
		...node.attrs,
		widthPx: image.widthPx,
		heightPx: image.heightPx,
		anchored: Boolean(image.anchored),
		placement: image.placement ? JSON.stringify(image.placement) : null,
		pictureFill: image.pictureFill ? JSON.stringify(image.pictureFill) : null,
	});
	view.dispatch(closeHistory(tr.setSelection(NodeSelection.create(tr.doc, pos))));
	history?.stopCapturing();
	return true;
}

/** Where an inline picture sits in its paragraph, so it starts floating in the same place. */
export function floatOrigin(view: EditorView, pos: number): FloatOrigin {
	try {
		const picture = view.nodeDOM(pos);
		const $pos = view.state.doc.resolve(pos);
		const paragraph = $pos.depth > 0 ? view.nodeDOM($pos.before()) : null;
		if (!(picture instanceof HTMLElement) || !(paragraph instanceof HTMLElement)) return {};
		const box = paragraph.getBoundingClientRect();
		const scale = paragraph.offsetWidth ? box.width / paragraph.offsetWidth : 1;
		const offsetXPx = (picture.getBoundingClientRect().left - box.left) / (scale || 1);
		return Number.isFinite(offsetXPx) ? { offsetXPx: Math.max(0, offsetXPx) } : {};
	} catch {
		return {};
	}
}

/** Wrap Text on the picture at `pos` (the selected one by default). */
export function setPictureWrap(view: EditorView, wrap: DrawingWrap, at?: number): boolean {
	const pos = at ?? selectedPicturePos(view.state);
	const node = pos === undefined ? undefined : view.state.doc.nodeAt(pos);
	if (pos === undefined || !isPictureNode(node) || pictureWrapOf(node) === wrap) return false;
	const origin = node.attrs.anchored ? undefined : floatOrigin(view, pos);
	return applyPicture(view, pos, setDrawingWrap(pictureOf(node), wrap, origin));
}

/**
 * Crops the picture at `pos` to `crop` (fractions of the whole picture). `size`, when given, is
 * the shown size after cropping; otherwise the picture keeps its scale, as Word's Crop fields do.
 */
export function setPictureCropAt(
	view: EditorView,
	pos: number,
	crop: CropInsets,
	size?: { widthPx: number; heightPx: number },
): boolean {
	const node = view.state.doc.nodeAt(pos);
	if (!isPictureNode(node)) return false;
	const cropped = setPictureCrop(pictureOf(node), crop);
	return applyPicture(view, pos, size ? { ...cropped, ...size } : cropped);
}

export function setGraphicsFillAt(
	view: EditorView,
	pos: number,
	patch: PictureEffectPatch,
): boolean {
	const node = view.state.doc.nodeAt(pos);
	if (!isPictureNode(node) || !node.attrs.svgPartName) return false;
	const picture = pictureOf(node);
	const fill = picture.pictureFill ?? { kind: 'picture' };
	const effects = patchPictureEffects(fill.effects ?? [], patch);
	return applyPicture(view, pos, { ...picture, pictureFill: { ...fill, effects } });
}

/** Selects the picture a pointer event landed on, as Word does before showing its menu. */
export function selectPictureAt(view: EditorView, target: EventTarget | null): boolean {
	const dom = target instanceof Element ? target.closest('.dve-picture') : null;
	if (!dom || !view.dom.contains(dom)) return false;
	let pos: number;
	try {
		pos = view.posAtDOM(dom, 0);
	} catch {
		return false;
	}
	const at = [pos, pos - 1].find((candidate) => isPictureNode(view.state.doc.nodeAt(candidate)));
	if (at === undefined) return false;
	const { selection } = view.state;
	if (!(selection instanceof NodeSelection && selection.from === at))
		view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, at)));
	return true;
}
