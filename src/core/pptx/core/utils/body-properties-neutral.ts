/**
 * The pptx object-tree adapter of the shared DrawingML body property codec
 * (`drawingml/text-body-codec.ts`, docs/agnostic-core-plan.md step 3).
 *
 * The `a:bodyPr` attribute and autofit rules are the codec's, the same ones the
 * DOM reader and writer run. This module reads an object-tree node through the
 * codec, converts the neutral values into the pptx {@link TextStyle} fields
 * (EMU to px, the PowerPoint autofit modes) and back, and assigns the codec's
 * entries onto an existing `a:bodyPr` node (a merge: attributes already there
 * keep their place). Text warp, 3D and linked text boxes stay with their own
 * readers.
 *
 * @module body-properties-neutral
 */
import type {
	DrawingAutofit,
	DrawingBodyProperties,
	DrawingTextWarp,
} from '../../../drawingml/text-model';
import {
	drawingAutofitChild,
	drawingBodyAttributeEntries,
	drawingTextWarpEntries,
	readDrawingBodyCodecProperties,
} from '../../../drawingml/text-body-codec';
import type { TextStyle, XmlObject } from '../types';
import { objectAttributes, objectTextElementView } from './paragraph-properties-neutral';

const EMU_PER_PX = 9525;

/** The codec families of an `a:bodyPr` object node; empty when absent. */
export function objectBodyProperties(bodyPr: unknown): DrawingBodyProperties {
	const view = objectTextElementView(bodyPr);
	return view ? readDrawingBodyCodecProperties(view) : {};
}

/** The `a:bodyPr` flag and rotation fields of {@link TextStyle}. */
export type BodyFlagFields = Pick<
	TextStyle,
	| 'compatibleLineSpacing'
	| 'forceAntiAlias'
	| 'upright'
	| 'fromWordArt'
	| 'spaceFirstLastParagraph'
	| 'anchorCenter'
	| 'rtlColumns'
	| 'textBodyRotation'
>;

/** The flag and rotation fields of neutral body properties. */
export function bodyFlagsFromDrawing(body: DrawingBodyProperties): BodyFlagFields {
	const out: BodyFlagFields = {};
	if (body.compatibleLineSpacing !== undefined) {
		out.compatibleLineSpacing = body.compatibleLineSpacing;
	}
	if (body.forceAntiAlias !== undefined) out.forceAntiAlias = body.forceAntiAlias;
	if (body.upright !== undefined) out.upright = body.upright;
	if (body.fromWordArt !== undefined) out.fromWordArt = body.fromWordArt;
	if (body.spaceFirstLastParagraph !== undefined) {
		out.spaceFirstLastParagraph = body.spaceFirstLastParagraph;
	}
	if (body.anchorCenter !== undefined) out.anchorCenter = body.anchorCenter;
	if (body.rightToLeftColumns !== undefined) out.rtlColumns = body.rightToLeftColumns;
	if (body.rotation !== undefined) out.textBodyRotation = body.rotation;
	return out;
}

/** The flag and rotation fields of a {@link TextStyle} on the neutral model. */
export function bodyFlagsToDrawing(style: BodyFlagFields): DrawingBodyProperties {
	const out: DrawingBodyProperties = {};
	if (style.compatibleLineSpacing !== undefined) {
		out.compatibleLineSpacing = style.compatibleLineSpacing;
	}
	if (style.forceAntiAlias !== undefined) out.forceAntiAlias = style.forceAntiAlias;
	if (style.upright !== undefined) out.upright = style.upright;
	if (style.fromWordArt !== undefined) out.fromWordArt = style.fromWordArt;
	if (style.spaceFirstLastParagraph !== undefined) {
		out.spaceFirstLastParagraph = style.spaceFirstLastParagraph;
	}
	if (style.anchorCenter !== undefined) out.anchorCenter = style.anchorCenter;
	if (style.rtlColumns !== undefined) out.rightToLeftColumns = style.rtlColumns;
	if (style.textBodyRotation !== undefined && Number.isFinite(style.textBodyRotation)) {
		out.rotation = style.textBodyRotation;
	}
	return out;
}

/** The inset fields of neutral body properties, in px. */
export function bodyInsetsFromDrawing(
	body: DrawingBodyProperties,
): Pick<TextStyle, 'bodyInsetLeft' | 'bodyInsetTop' | 'bodyInsetRight' | 'bodyInsetBottom'> {
	const out: Pick<
		TextStyle,
		'bodyInsetLeft' | 'bodyInsetTop' | 'bodyInsetRight' | 'bodyInsetBottom'
	> = {};
	if (body.insetLeftEmu !== undefined) out.bodyInsetLeft = body.insetLeftEmu / EMU_PER_PX;
	if (body.insetTopEmu !== undefined) out.bodyInsetTop = body.insetTopEmu / EMU_PER_PX;
	if (body.insetRightEmu !== undefined) out.bodyInsetRight = body.insetRightEmu / EMU_PER_PX;
	if (body.insetBottomEmu !== undefined) out.bodyInsetBottom = body.insetBottomEmu / EMU_PER_PX;
	return out;
}

/** The PowerPoint autofit fields of a neutral autofit choice. */
export function autofitFromDrawing(
	autofit: DrawingAutofit | undefined,
): Pick<TextStyle, 'autoFit' | 'autoFitMode' | 'autoFitFontScale' | 'autoFitLineSpacingReduction'> {
	if (autofit?.kind === 'shape') return { autoFit: true, autoFitMode: 'shrink' };
	if (autofit?.kind === 'none') return { autoFit: false, autoFitMode: 'none' };
	if (autofit?.kind !== 'normal') return {};
	return {
		autoFit: true,
		autoFitMode: 'normal',
		...(autofit.fontScale !== undefined && autofit.fontScale > 0
			? { autoFitFontScale: autofit.fontScale }
			: {}),
		...(autofit.lineSpacingReduction !== undefined && autofit.lineSpacingReduction > 0
			? { autoFitLineSpacingReduction: autofit.lineSpacingReduction }
			: {}),
	};
}

/** Assigns the codec's attribute entries of `value` onto an existing `a:bodyPr` node. */
export function assignBodyAttributes(bodyPr: XmlObject, value: DrawingBodyProperties): void {
	for (const [name, text] of drawingBodyAttributeEntries(value)) bodyPr[`@_${name}`] = text;
}

/**
 * The object-tree `a:prstTxWarp` node of a warp. `a:avLst` is always written (empty when the warp
 * keeps its default adjustments), as PowerPoint writes it; one guide is an object, several an array.
 */
export function objectTextWarp(warp: DrawingTextWarp): XmlObject {
	const entries = drawingTextWarpEntries({ ...warp, guides: warp.guides ?? [] });
	const node = objectAttributes(entries.attributes);
	const guides = (entries.guides ?? []).map(objectAttributes);
	node['a:avLst'] = guides.length > 0 ? { 'a:gd': guides.length === 1 ? guides[0] : guides } : {};
	return node;
}

/** The object-tree node of an autofit choice, keyed by its element name. */
export function objectAutofit(autofit: DrawingAutofit): { key: string; node: XmlObject } {
	const child = drawingAutofitChild(autofit);
	return { key: `a:${child.local}`, node: objectAttributes(child.attributes) };
}
