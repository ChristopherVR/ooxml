// Drawing shapes in the editor: placing a drawn shape, and changing the selected one.
import { closeHistory } from 'prosemirror-history';
import { NodeSelection, type EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import {
	DEFAULT_DRAWN_SHAPE,
	isDirectionalPreset,
	newShape,
	type PicturePlacement,
	type ShapeSpec,
} from 'ooxml-core/docx';
import type { DrawingFill } from 'ooxml-core/drawingml';
import { schema } from './schema';

export interface SelectedShape {
	pos: number;
	spec: ShapeSpec;
	widthPx: number;
	heightPx: number;
	anchored: boolean;
	placement: PicturePlacement | undefined;
}

const MIN_SIZE = 4;
const parse = <T>(value: unknown): T | undefined => {
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value) as T;
	} catch {
		return undefined;
	}
};

/** The shape the selection is on (an image atom that carries a shape), if any. */
export function selectedShape(state: EditorState): SelectedShape | undefined {
	const { selection } = state;
	if (!(selection instanceof NodeSelection) || selection.node.type !== schema.nodes.image)
		return undefined;
	const spec = parse<ShapeSpec>(selection.node.attrs.shape);
	if (!spec) return undefined;
	const placement = parse<PicturePlacement>(selection.node.attrs.placement);
	return {
		pos: selection.from,
		spec,
		widthPx: Number(selection.node.attrs.widthPx),
		heightPx: Number(selection.node.attrs.heightPx),
		anchored: Boolean(selection.node.attrs.anchored),
		placement,
	};
}

const solid = (hex: string): DrawingFill => ({
	kind: 'solid',
	color: { kind: 'srgb', value: hex.replace('#', '').toUpperCase(), transforms: [] },
});

interface Change {
	spec?: ShapeSpec;
	widthPx?: number;
	heightPx?: number;
	anchored?: boolean;
	placement?: PicturePlacement | null;
}

function apply(view: EditorView, selected: SelectedShape, change: Change): boolean {
	const node = view.state.doc.nodeAt(selected.pos);
	if (!view.editable || !node || !selected.spec.editable) return false;
	const attrs = {
		...node.attrs,
		...(change.spec ? { shape: JSON.stringify(change.spec) } : {}),
		...(change.widthPx !== undefined
			? { widthPx: Math.max(MIN_SIZE, Math.round(change.widthPx)) }
			: {}),
		...(change.heightPx !== undefined
			? { heightPx: Math.max(MIN_SIZE, Math.round(change.heightPx)) }
			: {}),
		...(change.anchored !== undefined ? { anchored: change.anchored } : {}),
		...(change.placement !== undefined
			? { placement: change.placement ? JSON.stringify(change.placement) : null }
			: {}),
	};
	const tr = view.state.tr.setNodeMarkup(selected.pos, undefined, attrs);
	view.dispatch(closeHistory(tr.setSelection(NodeSelection.create(tr.doc, selected.pos))));
	return true;
}

/** Shape Fill: a `#rrggbb` colour, `none`, or `default` for the theme style's fill. */
export function setShapeFill(view: EditorView, value: string): boolean {
	const selected = selectedShape(view.state);
	if (!selected) return false;
	const { fill: _old, ...rest } = selected.spec;
	const spec: ShapeSpec =
		value === 'default'
			? rest
			: { ...rest, fill: value === 'none' ? { kind: 'none' } : solid(value) };
	return apply(view, selected, { spec });
}

/** Shape Outline: a `#rrggbb` colour, `none`, or `default`. Keeps the weight and dash. */
export function setShapeOutline(view: EditorView, value: string): boolean {
	const selected = selectedShape(view.state);
	if (!selected) return false;
	const { line: old, ...rest } = selected.spec;
	if (value === 'default' && !old?.widthEmu && !old?.dash)
		return apply(view, selected, { spec: rest });
	const fill: DrawingFill | undefined =
		value === 'default' ? undefined : value === 'none' ? { kind: 'none' } : solid(value);
	return apply(view, selected, {
		spec: { ...rest, line: { ...old, ...(fill ? { fill } : {}) } },
	});
}

/** Shape Outline > Weight, in points. */
export function setShapeWeight(view: EditorView, points: number): boolean {
	const selected = selectedShape(view.state);
	if (!selected || !(points > 0)) return false;
	const line = { ...selected.spec.line, widthEmu: Math.round(points * 12700) };
	return apply(view, selected, { spec: { ...selected.spec, line } });
}

export function setShapeText(view: EditorView, lines: string[]): boolean {
	const selected = selectedShape(view.state);
	if (!selected?.spec.textEditable) return false;
	const { text: _old, ...rest } = selected.spec;
	return apply(view, selected, {
		spec: lines.some(Boolean) ? { ...rest, text: lines } : rest,
	});
}

export function setShapeSize(view: EditorView, widthPx?: number, heightPx?: number): boolean {
	const selected = selectedShape(view.state);
	return selected
		? apply(view, selected, {
				...(widthPx !== undefined ? { widthPx } : {}),
				...(heightPx !== undefined ? { heightPx } : {}),
			})
		: false;
}

export type ShapeWrap =
	| 'inline'
	| 'square'
	| 'tight'
	| 'through'
	| 'topAndBottom'
	| 'behind'
	| 'front';

/** Wrap Text: the wrapping of the shape, or In Line with Text. */
export function setShapeWrap(view: EditorView, wrap: ShapeWrap): boolean {
	const selected = selectedShape(view.state);
	if (!selected) return false;
	if (wrap === 'inline') return apply(view, selected, { anchored: false, placement: null });
	const base: PicturePlacement = selected.placement ?? {
		wrap: 'none',
		relativeFrom: 'column',
		offsetXPx: 0,
		relativeFromV: 'paragraph',
		offsetYPx: 0,
	};
	const { behindText: _behind, ...kept } = base;
	const placement: PicturePlacement =
		wrap === 'behind'
			? { ...kept, wrap: 'none', behindText: true }
			: wrap === 'front'
				? { ...kept, wrap: 'none' }
				: { ...kept, wrap };
	return apply(view, selected, { anchored: true, placement });
}

/** The Wrap Text choice that describes the selected shape. */
export function shapeWrapOf(shape: SelectedShape): ShapeWrap {
	if (!shape.anchored) return 'inline';
	const placement = shape.placement;
	if (!placement || placement.wrap === 'none') return placement?.behindText ? 'behind' : 'front';
	return placement.wrap;
}

export interface DrawnRect {
	/** Offset from the left edge of the text column and the top of the paragraph, in CSS pixels. */
	x: number;
	y: number;
	widthPx: number;
	heightPx: number;
	flipH?: boolean;
	flipV?: boolean;
}

/**
 * Places a drawn shape as a floating shape in front of text, anchored to the paragraph holding
 * `pos`, and selects it. Returns whether a shape was inserted.
 */
export function insertDrawnShape(
	view: EditorView,
	preset: string,
	pos: number,
	rect: Partial<DrawnRect> & Pick<DrawnRect, 'x' | 'y'>,
): boolean {
	if (!view.editable) return false;
	const $pos = view.state.doc.resolve(Math.min(pos, view.state.doc.content.size));
	let depth = $pos.depth;
	while (depth > 0 && $pos.node(depth).type.name !== 'paragraph') depth--;
	if (depth === 0) return false;
	const insertAt = $pos.start(depth);
	const directional = isDirectionalPreset(preset);
	const widthPx = Math.max(MIN_SIZE, Math.round(rect.widthPx ?? DEFAULT_DRAWN_SHAPE.widthPx));
	const heightPx = Math.max(
		directional ? 1 : MIN_SIZE,
		Math.round(rect.heightPx ?? (directional ? 0 : DEFAULT_DRAWN_SHAPE.heightPx)),
	);
	const spec: ShapeSpec = {
		...newShape(preset),
		...(rect.flipH ? { flipH: true } : {}),
		...(rect.flipV ? { flipV: true } : {}),
	};
	const placement: PicturePlacement = {
		wrap: 'none',
		relativeFrom: 'column',
		offsetXPx: Math.round(rect.x),
		relativeFromV: 'paragraph',
		offsetYPx: Math.round(rect.y),
	};
	const node = schema.nodes.image!.create({
		unsupported: 'Shape',
		widthPx,
		heightPx,
		anchored: true,
		placement: JSON.stringify(placement),
		shape: JSON.stringify(spec),
	});
	const tr = view.state.tr.insert(insertAt, node);
	view.dispatch(
		closeHistory(tr.setSelection(NodeSelection.create(tr.doc, insertAt))).scrollIntoView(),
	);
	return true;
}
