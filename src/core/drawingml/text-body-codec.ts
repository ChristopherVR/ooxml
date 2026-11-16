// The DrawingML body property codec (CT_TextBodyProperties, ECMA-376 Part 1, 21.1.2.1.1) for
// `a:bodyPr`: the attributes and the autofit choice (`a:noAutofit`, `a:spAutoFit`,
// `a:normAutofit`). Like the paragraph and run codecs it reads through a `DrawingTextElementView`,
// so the DOM reader (`read-text-body.ts`) and PowerPoint's object-tree reader share one set of
// rules, and it writes attribute entries in schema order. Text warp and 3D (`a:prstTxWarp`,
// `a:scene3d`, `a:sp3d`, `a:flatTx`) are read by their own codecs.
import { parseOoxmlPercent } from '../color/color-primitives';
import type { DrawingTextElementView } from './text-element-view';
import type { DrawingAutofit, DrawingBodyProperties } from './text-model';
import {
	drawingBooleanValue,
	drawingIntegerValue,
	drawingStringValue,
	type DrawingAttributeEntry,
	type DrawingChildEntry,
} from './text-paragraph-codec';

const BODY_STRINGS = [
	['verticalOverflow', 'vertOverflow'],
	['horizontalOverflow', 'horzOverflow'],
	['vertical', 'vert'],
	['wrap', 'wrap'],
	['anchor', 'anchor'],
] as const;

const BODY_INTEGERS = [
	['insetLeftEmu', 'lIns'],
	['insetTopEmu', 'tIns'],
	['insetRightEmu', 'rIns'],
	['insetBottomEmu', 'bIns'],
	['columnCount', 'numCol'],
	['columnSpacingEmu', 'spcCol'],
] as const;

const BODY_BOOLEANS = [
	['spaceFirstLastParagraph', 'spcFirstLastPara'],
	['rightToLeftColumns', 'rtlCol'],
	['fromWordArt', 'fromWordArt'],
	['anchorCenter', 'anchorCtr'],
	['forceAntiAlias', 'forceAA'],
	['upright', 'upright'],
	['compatibleLineSpacing', 'compatLnSpc'],
] as const;

/** Reads the attributes of `a:bodyPr` into `out` (`rot` in degrees). */
export function readDrawingBodyAttributes(
	view: DrawingTextElementView,
	out: DrawingBodyProperties,
): void {
	const rotation = drawingIntegerValue(view.attribute('rot'));
	if (rotation !== undefined) out.rotation = rotation / 60000;
	for (const [key, name] of BODY_STRINGS) {
		const value = drawingStringValue(view.attribute(name));
		if (value) out[key] = value;
	}
	for (const [key, name] of BODY_INTEGERS) {
		const value = drawingIntegerValue(view.attribute(name));
		if (value !== undefined) out[key] = value;
	}
	for (const [key, name] of BODY_BOOLEANS) {
		const value = drawingBooleanValue(view.attribute(name));
		if (value !== undefined) out[key] = value;
	}
}

/** The autofit children, in the order a reader that finds several prefers them. */
export const DRAWING_AUTOFIT_CHILDREN = ['spAutoFit', 'normAutofit', 'noAutofit'] as const;

/** Reads one autofit child (`noAutofit`, `spAutoFit`, `normAutofit`) by its local name. */
export function readDrawingAutofit(
	local: string,
	view: DrawingTextElementView,
): DrawingAutofit | undefined {
	if (local === 'noAutofit') return { kind: 'none' };
	if (local === 'spAutoFit') return { kind: 'shape' };
	if (local !== 'normAutofit') return undefined;
	const fontScaleText = view.attribute('fontScale');
	const reductionText = view.attribute('lnSpcReduction');
	const fontScale = parseOoxmlPercent(fontScaleText ?? '');
	const reduction = parseOoxmlPercent(reductionText ?? '');
	return {
		kind: 'normal',
		...(fontScaleText !== undefined && fontScale !== undefined ? { fontScale } : {}),
		...(reductionText !== undefined && reduction !== undefined
			? { lineSpacingReduction: reduction }
			: {}),
	};
}

/**
 * Reads the codec families of `a:bodyPr`: attributes and the autofit choice. Product readers that
 * hold their own tree use this; the DOM reader walks the children in document order and also reads
 * the 3D children and keeps the others.
 */
export function readDrawingBodyCodecProperties(
	view: DrawingTextElementView,
): DrawingBodyProperties {
	const out: DrawingBodyProperties = {};
	readDrawingBodyAttributes(view, out);
	for (const local of DRAWING_AUTOFIT_CHILDREN) {
		const child = view.child(local);
		if (!child) continue;
		const autofit = readDrawingAutofit(local, child);
		if (autofit) out.autofit = autofit;
		break;
	}
	return out;
}

const flag = (value: boolean) => (value ? '1' : '0');
const scaled = (value: number | undefined, scale: number) =>
	value !== undefined && Number.isFinite(value) ? String(Math.round(value * scale)) : undefined;

/** The attributes of `a:bodyPr`, in CT_TextBodyProperties order. */
export function drawingBodyAttributeEntries(value: DrawingBodyProperties): DrawingAttributeEntry[] {
	const out: DrawingAttributeEntry[] = [];
	const add = (name: string, text: string | undefined) => {
		if (text !== undefined) out.push([name, text]);
	};
	const bool = (field: boolean | undefined) => (field === undefined ? undefined : flag(field));
	add('rot', scaled(value.rotation, 60000));
	add('spcFirstLastPara', bool(value.spaceFirstLastParagraph));
	add('vertOverflow', value.verticalOverflow);
	add('horzOverflow', value.horizontalOverflow);
	add('vert', value.vertical);
	add('wrap', value.wrap);
	add('lIns', scaled(value.insetLeftEmu, 1));
	add('tIns', scaled(value.insetTopEmu, 1));
	add('rIns', scaled(value.insetRightEmu, 1));
	add('bIns', scaled(value.insetBottomEmu, 1));
	add('numCol', scaled(value.columnCount, 1));
	add('spcCol', scaled(value.columnSpacingEmu, 1));
	add('rtlCol', bool(value.rightToLeftColumns));
	add('fromWordArt', bool(value.fromWordArt));
	add('anchor', value.anchor);
	add('anchorCtr', bool(value.anchorCenter));
	add('forceAA', bool(value.forceAntiAlias));
	add('upright', bool(value.upright));
	add('compatLnSpc', bool(value.compatibleLineSpacing));
	return out;
}

/** The autofit child of `a:bodyPr`. */
export function drawingAutofitChild(autofit: DrawingAutofit): DrawingChildEntry {
	if (autofit.kind === 'none') return { local: 'noAutofit', attributes: [] };
	if (autofit.kind === 'shape') return { local: 'spAutoFit', attributes: [] };
	const attributes: DrawingAttributeEntry[] = [];
	const fontScale = scaled(autofit.fontScale, 100000);
	if (fontScale !== undefined) attributes.push(['fontScale', fontScale]);
	const reduction = scaled(autofit.lineSpacingReduction, 100000);
	if (reduction !== undefined) attributes.push(['lnSpcReduction', reduction]);
	return { local: 'normAutofit', attributes };
}
