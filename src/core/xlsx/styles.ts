import type { Alignment, Border, CellStyle, Fill, Font, Protection, Workbook } from './model.js';

/** A stable key for a style, so equal formats share one `cellXfs` entry. */
export function styleKey(style: CellStyle): string {
	return JSON.stringify(sortKeys(style));
}

function sortKeys(value: unknown): unknown {
	if (Array.isArray(value)) return value.map(sortKeys);
	if (value && typeof value === 'object') {
		const out: Record<string, unknown> = {};
		for (const key of Object.keys(value).sort()) {
			const entry = (value as Record<string, unknown>)[key];
			if (entry !== undefined) out[key] = sortKeys(entry);
		}
		return out;
	}
	return value;
}

/** The style at `styleId`, falling back to the default when the id is out of range. */
export function styleAt(workbook: Workbook, styleId: number | undefined): CellStyle {
	const style = workbook.styles[styleId ?? 0] ?? workbook.styles[0];
	if (!style) throw new Error('Workbook has no default cell style');
	return style;
}

/** The id of a style equal to `style`, adding it to the workbook when it is new. */
export function internStyle(workbook: Workbook, style: CellStyle): number {
	const key = styleKey(style);
	const cache = styleCache(workbook);
	const known = cache.get(key);
	if (known !== undefined && workbook.styles[known] && styleKey(workbook.styles[known]) === key)
		return known;
	const index = workbook.styles.findIndex((candidate) => styleKey(candidate) === key);
	if (index >= 0) {
		cache.set(key, index);
		return index;
	}
	workbook.styles.push(style);
	cache.set(key, workbook.styles.length - 1);
	return workbook.styles.length - 1;
}

const caches = new WeakMap<Workbook, Map<string, number>>();
function styleCache(workbook: Workbook): Map<string, number> {
	let cache = caches.get(workbook);
	if (!cache) {
		cache = new Map();
		caches.set(workbook, cache);
	}
	return cache;
}

/** A partial change to a cell format, as the ribbon and format dialog produce. */
export interface StylePatch {
	font?: Partial<Font>;
	fill?: Fill;
	border?: Partial<Border>;
	numFmt?: string;
	alignment?: Partial<Alignment>;
	protection?: Partial<Protection>;
	/**
	 * The named cell style (`Good`, `Heading 1`) the format derives from; `undefined` (when the
	 * key is present) detaches it. Formatting is not changed by this key alone: use
	 * `applyCellStyle` to apply a named style's formatting.
	 */
	cellStyleName?: string | undefined;
}

/** Removes keys set to `undefined` so patched styles keep a canonical shape. */
function compact<T extends object>(value: T): T {
	const out: Record<string, unknown> = {};
	for (const [key, entry] of Object.entries(value)) if (entry !== undefined) out[key] = entry;
	return out as T;
}

/** A new style: `base` with `patch` applied. A key patched to `undefined` is cleared. */
export function patchStyle(base: CellStyle, patch: StylePatch): CellStyle {
	const next: CellStyle = {
		font: patch.font ? compact({ ...base.font, ...patch.font }) : base.font,
		fill: patch.fill ?? base.fill,
		border: patch.border ? compact({ ...base.border, ...patch.border }) : base.border,
		numFmt: patch.numFmt ?? base.numFmt,
	};
	const alignment = patch.alignment
		? compact({ ...base.alignment, ...patch.alignment })
		: base.alignment;
	if (alignment && Object.keys(alignment).length) next.alignment = alignment;
	const protection = patch.protection
		? compact({ ...base.protection, ...patch.protection })
		: base.protection;
	if (protection && Object.keys(protection).length) next.protection = protection;
	const styleName = 'cellStyleName' in patch ? patch.cellStyleName : base.cellStyleName;
	if (styleName) next.cellStyleName = styleName;
	return next;
}

/** Applies a patch to the style at `styleId` and returns the id of the result. */
export function applyStylePatch(
	workbook: Workbook,
	styleId: number | undefined,
	patch: StylePatch,
): number {
	return internStyle(workbook, patchStyle(styleAt(workbook, styleId), patch));
}
