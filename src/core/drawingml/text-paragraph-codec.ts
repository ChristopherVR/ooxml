// The DrawingML paragraph property codec (CT_TextParagraphProperties, ECMA-376 Part 1, 21.1.2.2.7)
// for the families every product reads the same way: the paragraph attributes, line and paragraph
// spacing (`a:lnSpc`, `a:spcBef`, `a:spcAft`), the bullet group except its colour and picture
// (`a:buSz*`, `a:buFont*`, `a:buNone`, `a:buChar`, `a:buAutoNum`), tab stops (`a:tabLst`) and
// `CT_TextFont` slots. It reads through a `DrawingTextElementView`, so the DOM reader
// (`read-paragraph-properties.ts`) and PowerPoint's object-tree adapter share one set of rules, and
// it writes attribute entries in schema order that the DOM writer and the object-tree writer both
// apply. Values keep the file's units.
import { parseOoxmlPercent } from '../color/color-primitives';
import type { DrawingTextElementView } from './text-element-view';
import type {
	DrawingBullet,
	DrawingFont,
	DrawingParagraphProperties,
	DrawingTabStop,
} from './text-model';
import type { DrawingTextSpacing } from './types';

/** One attribute to write: local name and lexical value. */
export type DrawingAttributeEntry = readonly [name: string, value: string];

/** One child element to write: its local name and attributes. */
export interface DrawingChildEntry {
	local: string;
	attributes: DrawingAttributeEntry[];
}

/** An integer value (base 10); `undefined` when absent or not a number. */
export function drawingIntegerValue(value: string | undefined): number | undefined {
	const parsed = Number.parseInt(value ?? '', 10);
	return Number.isFinite(parsed) ? parsed : undefined;
}

/** `1`, `true` -> true; `0`, `false` -> false; otherwise `undefined`. */
export function drawingBooleanValue(value: string | undefined): boolean | undefined {
	const token = value?.trim().toLowerCase();
	if (token === '1' || token === 'true') return true;
	if (token === '0' || token === 'false') return false;
	return undefined;
}

/** A trimmed, non-empty value. */
export function drawingStringValue(value: string | undefined): string | undefined {
	const trimmed = value?.trim();
	return trimmed ? trimmed : undefined;
}

const EMU_ATTRIBUTES = [
	['marginLeftEmu', 'marL'],
	['marginRightEmu', 'marR'],
	['indentEmu', 'indent'],
	['defaultTabSizeEmu', 'defTabSz'],
] as const;

const BOOLEAN_ATTRIBUTES = [
	['rightToLeft', 'rtl'],
	['eastAsianLineBreak', 'eaLnBrk'],
	['latinLineBreak', 'latinLnBrk'],
	['hangingPunctuation', 'hangingPunct'],
] as const;

/** Reads the attributes of `a:pPr` / `a:lvlNpPr` into `out` (`lvl` clamped to 0-8). */
export function readDrawingParagraphAttributes(
	view: DrawingTextElementView,
	out: DrawingParagraphProperties,
): void {
	const level = drawingIntegerValue(view.attribute('lvl'));
	if (level !== undefined) out.level = Math.min(Math.max(level, 0), 8);
	const align = drawingStringValue(view.attribute('algn'));
	if (align) out.align = align;
	for (const [key, name] of EMU_ATTRIBUTES) {
		const value = drawingIntegerValue(view.attribute(name));
		if (value !== undefined) out[key] = value;
	}
	for (const [key, name] of BOOLEAN_ATTRIBUTES) {
		const value = drawingBooleanValue(view.attribute(name));
		if (value !== undefined) out[key] = value;
	}
	const fontAlign = drawingStringValue(view.attribute('fontAlgn'));
	if (fontAlign) out.fontAlign = fontAlign;
}

const flag = (value: boolean) => (value ? '1' : '0');
const emu = (value: number) => String(Math.round(value));

/** The attributes of paragraph properties, in CT_TextParagraphProperties order. */
export function drawingParagraphAttributeEntries(
	value: DrawingParagraphProperties,
): DrawingAttributeEntry[] {
	const out: DrawingAttributeEntry[] = [];
	const add = (name: string, text: string | undefined) => {
		if (text !== undefined) out.push([name, text]);
	};
	const number = (field: number | undefined) =>
		field !== undefined && Number.isFinite(field) ? emu(field) : undefined;
	const bool = (field: boolean | undefined) => (field === undefined ? undefined : flag(field));
	add('marL', number(value.marginLeftEmu));
	add('marR', number(value.marginRightEmu));
	add('lvl', number(value.level));
	add('indent', number(value.indentEmu));
	add('algn', value.align);
	add('defTabSz', number(value.defaultTabSizeEmu));
	add('rtl', bool(value.rightToLeft));
	add('eaLnBrk', bool(value.eastAsianLineBreak));
	add('fontAlgn', value.fontAlign);
	add('latinLnBrk', bool(value.latinLineBreak));
	add('hangingPunct', bool(value.hangingPunctuation));
	return out;
}

/** Reads `a:lnSpc`, `a:spcBef` or `a:spcAft`: points (`a:spcPts`, which wins) or a fraction. */
export function readDrawingTextSpacing(
	view: DrawingTextElementView | undefined,
): DrawingTextSpacing | undefined {
	const points = drawingIntegerValue(view?.child('spcPts')?.attribute('val'));
	if (points !== undefined) return { unit: 'points', value: points / 100 };
	const percent = parseOoxmlPercent(view?.child('spcPct')?.attribute('val'));
	if (percent !== undefined) return { unit: 'percent', value: percent };
	return undefined;
}

/** The child of a spacing element: `a:spcPts` (hundredths of a point) or `a:spcPct`. */
export function drawingTextSpacingChild(spacing: DrawingTextSpacing): DrawingChildEntry {
	const points = spacing.unit === 'points';
	return {
		local: points ? 'spcPts' : 'spcPct',
		attributes: [['val', String(Math.round(spacing.value * (points ? 100 : 100000)))]],
	};
}

/** Reads one `a:tab`. */
export function readDrawingTabStop(view: DrawingTextElementView): DrawingTabStop {
	const stop: DrawingTabStop = {};
	const position = drawingIntegerValue(view.attribute('pos'));
	if (position !== undefined) stop.positionEmu = position;
	const align = drawingStringValue(view.attribute('algn'));
	if (align) stop.align = align;
	const leader = drawingStringValue(view.attribute('leader'));
	if (leader) stop.leader = leader;
	return stop;
}

/** The attributes of one `a:tab`. */
export function drawingTabStopEntries(stop: DrawingTabStop): DrawingAttributeEntry[] {
	const out: DrawingAttributeEntry[] = [];
	if (stop.positionEmu !== undefined && Number.isFinite(stop.positionEmu))
		out.push(['pos', emu(stop.positionEmu)]);
	if (stop.align) out.push(['algn', stop.align]);
	if (stop.leader) out.push(['leader', stop.leader]);
	return out;
}

/** Reads a `CT_TextFont` slot (`a:latin`, `a:buFont`...); `undefined` without a typeface. */
export function readDrawingFont(view: DrawingTextElementView | undefined): DrawingFont | undefined {
	const typeface = drawingStringValue(view?.attribute('typeface'));
	if (!view || typeface === undefined) return undefined;
	const font: DrawingFont = { typeface };
	const panose = drawingStringValue(view.attribute('panose'));
	if (panose) font.panose = panose;
	const pitchFamily = drawingIntegerValue(view.attribute('pitchFamily'));
	if (pitchFamily !== undefined) font.pitchFamily = pitchFamily;
	const charset = drawingIntegerValue(view.attribute('charset'));
	if (charset !== undefined) font.charset = charset;
	return font;
}

/** The attributes of a `CT_TextFont` slot. */
export function drawingFontEntries(font: DrawingFont): DrawingAttributeEntry[] {
	const out: DrawingAttributeEntry[] = [['typeface', font.typeface]];
	if (font.panose !== undefined) out.push(['panose', font.panose]);
	if (font.pitchFamily !== undefined) out.push(['pitchFamily', String(font.pitchFamily)]);
	if (font.charset !== undefined) out.push(['charset', String(font.charset)]);
	return out;
}

/**
 * Reads one bullet-group child other than `a:buClr` and `a:buBlip` (which the caller reads with its
 * colour and picture readers) into `out`; false when `local` is not one or it carries no value.
 */
export function readDrawingBulletChild(
	local: string,
	view: DrawingTextElementView,
	out: DrawingParagraphProperties,
): boolean {
	switch (local) {
		case 'buClrTx':
			out.bulletColor = 'text';
			return true;
		case 'buSzTx':
			out.bulletSize = 'text';
			return true;
		case 'buSzPct': {
			const value = parseOoxmlPercent(view.attribute('val'));
			if (value !== undefined) out.bulletSize = { unit: 'percent', value };
			return value !== undefined;
		}
		case 'buSzPts': {
			const value = drawingIntegerValue(view.attribute('val'));
			if (value !== undefined) out.bulletSize = { unit: 'points', value: value / 100 };
			return value !== undefined;
		}
		case 'buFontTx':
			out.bulletFont = 'text';
			return true;
		case 'buFont': {
			const font = readDrawingFont(view);
			if (font) out.bulletFont = font;
			return font !== undefined;
		}
		case 'buNone':
			out.bullet = { kind: 'none' };
			return true;
		case 'buChar': {
			const char = view.attribute('char');
			if (char) out.bullet = { kind: 'char', char };
			return Boolean(char);
		}
		case 'buAutoNum': {
			const scheme = drawingStringValue(view.attribute('type'));
			if (!scheme) return false;
			const startAt = drawingIntegerValue(view.attribute('startAt'));
			out.bullet = { kind: 'autoNumber', scheme, ...(startAt !== undefined ? { startAt } : {}) };
			return true;
		}
		default:
			return false;
	}
}

/** The `a:buSz*` child of a bullet size. */
export function drawingBulletSizeChild(
	size: NonNullable<DrawingParagraphProperties['bulletSize']>,
): DrawingChildEntry {
	if (size === 'text') return { local: 'buSzTx', attributes: [] };
	const percent = size.unit === 'percent';
	return {
		local: percent ? 'buSzPct' : 'buSzPts',
		attributes: [['val', String(Math.round(size.value * (percent ? 100000 : 100)))]],
	};
}

/** The bullet type child (`a:buNone`, `a:buChar`, `a:buAutoNum`); a picture is the caller's. */
export function drawingBulletChild(bullet: DrawingBullet): DrawingChildEntry | undefined {
	switch (bullet.kind) {
		case 'none':
			return { local: 'buNone', attributes: [] };
		case 'char':
			return { local: 'buChar', attributes: [['char', bullet.char]] };
		case 'autoNumber': {
			const attributes: DrawingAttributeEntry[] = [['type', bullet.scheme]];
			if (bullet.startAt !== undefined) attributes.push(['startAt', String(bullet.startAt)]);
			return { local: 'buAutoNum', attributes };
		}
		default:
			return undefined;
	}
}

/** The bullet-group children this codec reads, in schema order. */
export const CODEC_BULLET_CHILDREN = [
	'buClrTx',
	'buSzTx',
	'buSzPct',
	'buSzPts',
	'buFontTx',
	'buFont',
	'buNone',
	'buAutoNum',
	'buChar',
] as const;

const SPACING_CHILDREN = [
	['lineSpacing', 'lnSpc'],
	['spaceBefore', 'spcBef'],
	['spaceAfter', 'spcAft'],
] as const;

/**
 * Reads the codec families of a paragraph properties element: attributes, spacing, the bullet
 * group without `a:buClr` and `a:buBlip`, and tab stops. Product readers that hold their own tree
 * (PowerPoint's object tree) use this; the DOM reader walks the children in document order and
 * also keeps what the codec does not cover. Children are read in schema order, so when a file
 * repeats a choice (two bullet types) the later one in the schema wins, as it does over the DOM.
 */
export function readDrawingParagraphCodecProperties(
	view: DrawingTextElementView,
): DrawingParagraphProperties {
	const out: DrawingParagraphProperties = {};
	readDrawingParagraphAttributes(view, out);
	for (const [key, local] of SPACING_CHILDREN) {
		const spacing = readDrawingTextSpacing(view.child(local));
		if (spacing) out[key] = spacing;
	}
	for (const local of CODEC_BULLET_CHILDREN) {
		const child = view.child(local);
		if (child) readDrawingBulletChild(local, child, out);
	}
	const tabList = view.child('tabLst');
	if (tabList) out.tabStops = tabList.children('tab').map(readDrawingTabStop);
	return out;
}
