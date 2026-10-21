// Formats compared by what they look like, not by style index: two files (or one file before and
// after a save) can number the same formats differently, and spell a default either way (`bold:
// false` or no `bold`, `rgb: 'FF0000'` or `'FFFF0000'`). Each format is reduced to canonical keys
// of its effective attributes, then interned so a cell comparison is one integer test.
import type {
	Alignment,
	Border,
	BorderEdge,
	CellStyle,
	Color,
	Fill,
	Font,
	Protection,
	Workbook,
} from '../model';
import { stableKey } from '../stable-key';
import { cellFont } from '../layout/style-view';
import { styleAt } from '../styles';
import type { FormatAttribute } from './types';

/** The attributes of a cell format, in report order (`richText` belongs to the cell). */
const ATTRIBUTES = [
	'numberFormat',
	'font',
	'fill',
	'border',
	'alignment',
	'protection',
] as const satisfies readonly FormatAttribute[];

function color(value: Color | undefined): Color | undefined {
	if (!value) return undefined;
	const out: Color = {};
	if (value.rgb) {
		const rgb = value.rgb.toUpperCase();
		out.rgb = rgb.length === 6 ? `FF${rgb}` : rgb;
	}
	if (value.theme !== undefined) out.theme = value.theme;
	if (value.indexed !== undefined) out.indexed = value.indexed;
	if (value.auto) out.auto = true;
	if (value.tint) out.tint = value.tint;
	return out;
}

/** A font without its defaults; `family` is round-trip metadata and does not change the look. */
export function canonicalFont(value: Font | undefined): Font {
	const out: Font = {};
	if (!value) return out;
	if (value.name !== undefined) out.name = value.name;
	if (value.size !== undefined) out.size = value.size;
	if (value.bold) out.bold = true;
	if (value.italic) out.italic = true;
	if (value.underline) out.underline = value.underline;
	if (value.strike) out.strike = true;
	const c = color(value.color);
	if (c) out.color = c;
	if (value.vertAlign) out.vertAlign = value.vertAlign;
	if (value.scheme) out.scheme = value.scheme;
	return out;
}

function fill(value: Fill): Fill {
	if (value.type === 'gradient')
		return {
			...value,
			stops: value.stops.map((stop) => ({
				position: stop.position,
				color: color(stop.color) ?? {},
			})),
		};
	if (value.pattern === 'none') return { type: 'pattern', pattern: 'none' };
	const out: Fill = { type: 'pattern', pattern: value.pattern };
	const fg = color(value.fgColor);
	if (fg) out.fgColor = fg;
	// A solid fill paints only its foreground colour.
	const bg = value.pattern === 'solid' ? undefined : color(value.bgColor);
	if (bg) out.bgColor = bg;
	return out;
}

/** An edge with no line draws nothing, as the reader treats `style="none"`. */
function edge(value: BorderEdge | undefined): BorderEdge | undefined {
	if (!value || (value.style as string) === 'none') return undefined;
	const c = color(value.color);
	return c ? { style: value.style, color: c } : { style: value.style };
}

function border(value: Border): Border {
	const out: Border = {};
	for (const side of ['left', 'right', 'top', 'bottom', 'diagonal'] as const) {
		const e = edge(value[side]);
		if (e) out[side] = e;
	}
	if (value.diagonalUp) out.diagonalUp = true;
	if (value.diagonalDown) out.diagonalDown = true;
	return out;
}

function alignment(value: Alignment | undefined): Alignment {
	const out: Alignment = {};
	if (!value) return out;
	if (value.horizontal && value.horizontal !== 'general') out.horizontal = value.horizontal;
	if (value.vertical && value.vertical !== 'bottom') out.vertical = value.vertical;
	if (value.wrapText) out.wrapText = true;
	if (value.shrinkToFit) out.shrinkToFit = true;
	if (value.indent) out.indent = value.indent;
	if (value.textRotation) out.textRotation = value.textRotation;
	if (value.readingOrder) out.readingOrder = value.readingOrder;
	return out;
}

/** Cells are locked and not hidden unless the format says otherwise. */
function protection(value: Protection | undefined): Protection {
	const out: Protection = {};
	if (value?.locked === false) out.locked = false;
	if (value?.hidden) out.hidden = true;
	return out;
}

/** Canonical keys of a format's attributes, in {@link ATTRIBUTES} order. */
export function formatAttributeKeys(style: CellStyle): string[] {
	return [
		JSON.stringify(style.numFmt),
		stableKey(canonicalFont(style.font)),
		stableKey(fill(style.fill)),
		stableKey(border(style.border)),
		stableKey(alignment(style.alignment)),
		stableKey(protection(style.protection)),
	];
}

/**
 * A copy of a plain value with every colour in it canonical (RGB with alpha, upper case, no zero
 * tint) and nothing else changed: for differential formats, whose `false` flags switch a setting
 * off and so are not defaults.
 */
export function withCanonicalColors<T>(value: T): T {
	if (Array.isArray(value)) return value.map(withCanonicalColors) as T;
	if (!value || typeof value !== 'object') return value;
	const entries = Object.entries(value).map(([key, entry]) => [key, withCanonicalColors(entry)]);
	const out = Object.fromEntries(entries) as Record<string, unknown>;
	if (typeof out.rgb === 'string') {
		const rgb = out.rgb.toUpperCase();
		out.rgb = rgb.length === 6 ? `FF${rgb}` : rgb;
	}
	if (out.tint === 0) delete out.tint;
	return out as T;
}

/**
 * The format at `styleId` with the font it paints with: a font without a name or size takes the
 * default style's (`cellFont`, as the grid does), so changing style 0 shows in the cells using it.
 */
export function paintedStyle(workbook: Workbook, styleId: number | undefined): CellStyle {
	const style = styleAt(workbook, styleId);
	return { ...style, font: cellFont(style.font, workbook.styles[0]?.font ?? {}) };
}

/** Interns the formats of the workbooks being compared into shared canonical ids. */
export class FormatIndex {
	private readonly ids = new Map<string, number>();
	private readonly keys: string[][] = [];
	private readonly styleIds = new Map<Workbook, number[]>();
	private readonly diffs = new Map<number, FormatAttribute[]>();

	/** The canonical id of the format at `styleId` of `workbook`. */
	idOf(workbook: Workbook, styleId: number | undefined): number {
		let ids = this.styleIds.get(workbook);
		if (!ids) {
			ids = [];
			this.styleIds.set(workbook, ids);
		}
		const index = styleId ?? 0;
		return (ids[index] ??= this.intern(formatAttributeKeys(paintedStyle(workbook, index))));
	}

	/** The canonical id of a format with these attribute keys, added when it is new. */
	private intern(keys: string[]): number {
		const key = keys.join('\n');
		let id = this.ids.get(key);
		if (id === undefined) {
			id = this.keys.length;
			this.keys.push(keys);
			this.ids.set(key, id);
		}
		return id;
	}

	/**
	 * The attributes that differ between two canonical formats (empty when they are equal). A new
	 * array each time, since results hand it out.
	 */
	attributesBetween(a: number, b: number): FormatAttribute[] {
		if (a === b) return [];
		const pair = a * 0x100000 + b;
		let out = this.diffs.get(pair);
		if (!out) {
			const left = this.keys[a] ?? [];
			const right = this.keys[b] ?? [];
			out = ATTRIBUTES.filter((_, i) => left[i] !== right[i]);
			this.diffs.set(pair, out);
		}
		return [...out];
	}
}
