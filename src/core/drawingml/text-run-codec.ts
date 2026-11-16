// The DrawingML run property codec (CT_TextCharacterProperties, ECMA-376 Part 1, 21.1.2.3.2) for
// `a:rPr`, `a:defRPr` and `a:endParaRPr`: the attributes, the font slots (`a:latin`, `a:ea`,
// `a:cs`, `a:sym`), the run `a:rtl` and the underline follow-text markers (`a:uLnTx`,
// `a:uFillTx`). Like the paragraph codec it reads through a `DrawingTextElementView`, so the DOM
// reader (`read-run-properties.ts`) and PowerPoint's object-tree reader share one set of rules, and
// it writes attribute entries in schema order. Fills, outlines, effects, highlight and hyperlinks
// are read by their own codecs. Values keep the file's units (points, fractions).
import { parseOoxmlPercent } from '../color/color-primitives';
import type { DrawingTextElementView } from './text-element-view';
import type { DrawingRunProperties } from './text-model';
import {
	drawingBooleanValue,
	drawingIntegerValue,
	drawingStringValue,
	readDrawingFont,
	type DrawingAttributeEntry,
} from './text-paragraph-codec';

const BOOLEAN_ATTRIBUTES = [
	['bold', 'b'],
	['italic', 'i'],
	['kumimoji', 'kumimoji'],
	['normalizeHeight', 'normalizeH'],
	['noProof', 'noProof'],
	['dirty', 'dirty'],
	['spellingError', 'err'],
	['smartTagClean', 'smtClean'],
] as const;

const STRING_ATTRIBUTES = [
	['underline', 'u'],
	['strike', 'strike'],
	['caps', 'cap'],
	['language', 'lang'],
	['altLanguage', 'altLang'],
	['bookmark', 'bmk'],
] as const;

/** The font slots of a run, by local name. */
export const DRAWING_RUN_FONT_SLOTS = [
	['latin', 'latin'],
	['ea', 'eastAsian'],
	['cs', 'complexScript'],
	['sym', 'symbol'],
] as const;

/** Reads the attributes of `a:rPr` / `a:defRPr` / `a:endParaRPr` into `out`. */
export function readDrawingRunAttributes(
	view: DrawingTextElementView,
	out: DrawingRunProperties,
): void {
	const size = drawingIntegerValue(view.attribute('sz'));
	if (size !== undefined) out.sizePt = size / 100;
	for (const [key, name] of BOOLEAN_ATTRIBUTES) {
		const value = drawingBooleanValue(view.attribute(name));
		if (value !== undefined) out[key] = value;
	}
	for (const [key, name] of STRING_ATTRIBUTES) {
		const value = drawingStringValue(view.attribute(name));
		if (value) out[key] = value;
	}
	const baseline = view.attribute('baseline');
	const baselineFraction = baseline === undefined ? undefined : parseOoxmlPercent(baseline);
	if (baselineFraction !== undefined) out.baseline = baselineFraction;
	const spacing = drawingIntegerValue(view.attribute('spc'));
	if (spacing !== undefined) out.spacingPt = spacing / 100;
	const kern = drawingIntegerValue(view.attribute('kern'));
	if (kern !== undefined) out.kernPt = kern / 100;
	const smartTagId = drawingIntegerValue(view.attribute('smtId'));
	if (smartTagId !== undefined) out.smartTagId = smartTagId;
}

/** Reads a run `a:rtl` (CT_Boolean, `val` defaulting to true). */
export function readDrawingRunRightToLeft(view: DrawingTextElementView): boolean {
	return drawingBooleanValue(view.attribute('val')) ?? true;
}

/**
 * Reads the codec families of run properties: attributes, font slots, `a:rtl` and the underline
 * follow-text markers. Product readers that hold their own tree use this; the DOM reader walks the
 * children in document order and also reads the fills, outline, highlight and hyperlinks.
 */
export function readDrawingRunCodecProperties(view: DrawingTextElementView): DrawingRunProperties {
	const out: DrawingRunProperties = {};
	readDrawingRunAttributes(view, out);
	for (const [local, key] of DRAWING_RUN_FONT_SLOTS) {
		const font = readDrawingFont(view.child(local));
		if (font) out[key] = font;
	}
	if (view.child('uLnTx')) out.underlineLineFollowsText = true;
	if (view.child('uFillTx')) out.underlineFillFollowsText = true;
	const rtl = view.child('rtl');
	if (rtl) out.rightToLeft = readDrawingRunRightToLeft(rtl);
	return out;
}

const flag = (value: boolean) => (value ? '1' : '0');
const scaled = (value: number | undefined, scale: number) =>
	value !== undefined && Number.isFinite(value) ? String(Math.round(value * scale)) : undefined;

/** The attributes of run properties, in CT_TextCharacterProperties order. */
export function drawingRunAttributeEntries(value: DrawingRunProperties): DrawingAttributeEntry[] {
	const out: DrawingAttributeEntry[] = [];
	const add = (name: string, text: string | undefined) => {
		if (text !== undefined) out.push([name, text]);
	};
	const bool = (field: boolean | undefined) => (field === undefined ? undefined : flag(field));
	add('kumimoji', bool(value.kumimoji));
	add('lang', value.language);
	add('altLang', value.altLanguage);
	add('sz', scaled(value.sizePt, 100));
	add('b', bool(value.bold));
	add('i', bool(value.italic));
	add('u', value.underline);
	add('strike', value.strike);
	add('kern', scaled(value.kernPt, 100));
	add('cap', value.caps);
	add('spc', scaled(value.spacingPt, 100));
	add('normalizeH', bool(value.normalizeHeight));
	add('baseline', scaled(value.baseline, 100000));
	add('noProof', bool(value.noProof));
	add('dirty', bool(value.dirty));
	add('err', bool(value.spellingError));
	add('smtClean', bool(value.smartTagClean));
	add('smtId', value.smartTagId === undefined ? undefined : String(value.smartTagId));
	add('bmk', value.bookmark);
	return out;
}
