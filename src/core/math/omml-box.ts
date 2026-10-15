/**
 * Box model for drawing Office Math without a browser: every construct becomes a box with a width,
 * an ascent and a descent from the baseline plus the primitives (text, lines, paths) it paints.
 * Metrics are an estimate of Cambria Math, close enough to size an equation on a printed page and
 * draw it as SVG; they never depend on the DOM, so layout and tests are deterministic.
 */
import type { OmmlNode } from './omml-node';

export type Prim =
	| {
			kind: 'text';
			x: number;
			y: number;
			text: string;
			size: number;
			italic: boolean;
			bold: boolean;
			/** Width the run is fitted to, so glyph spacing follows the estimate. */
			fit: number;
			/** Vertical stretch for large delimiters. */
			stretch?: number;
	  }
	| { kind: 'line'; x1: number; y1: number; x2: number; y2: number; width: number }
	| { kind: 'path'; d: string; width: number }
	| { kind: 'rect'; x: number; y: number; w: number; h: number; width: number };

export interface Box {
	width: number;
	/** Extent above the baseline (positive). */
	ascent: number;
	/** Extent below the baseline (positive). */
	descent: number;
	prims: Prim[];
}

/** Style carried down the tree: font size and how deep in scripts we are. */
export interface MathStyle {
	size: number;
	/** 0 text, 1 script, 2+ script of script. */
	level: number;
	plain?: boolean;
	bold?: boolean;
	italic?: boolean;
}

export const AXIS = 0.25;
export const emptyBox = (): Box => ({ width: 0, ascent: 0, descent: 0, prims: [] });

export function scriptStyle(style: MathStyle): MathStyle {
	return {
		...style,
		level: style.level + 1,
		size: style.level >= 2 ? style.size : style.size * 0.7,
	};
}

/** Moves a box's primitives by `(dx, dy)`; y grows downward from the baseline. */
export function shifted(prims: Prim[], dx: number, dy: number): Prim[] {
	return prims.map((prim): Prim => {
		switch (prim.kind) {
			case 'text':
				return { ...prim, x: prim.x + dx, y: prim.y + dy };
			case 'line':
				return { ...prim, x1: prim.x1 + dx, x2: prim.x2 + dx, y1: prim.y1 + dy, y2: prim.y2 + dy };
			case 'rect':
				return { ...prim, x: prim.x + dx, y: prim.y + dy };
			default:
				return { ...prim, d: translatePath(prim.d, dx, dy) };
		}
	});
}

/** Paths are written with absolute `M`/`L` pairs only. */
function translatePath(d: string, dx: number, dy: number): string {
	return d.replace(
		/([ML])\s*(-?[\d.]+)[ ,]+(-?[\d.]+)/g,
		(_m, cmd: string, x: string, y: string) =>
			`${cmd}${round(Number(x) + dx)} ${round(Number(y) + dy)}`,
	);
}
export const round = (value: number): number => Math.round(value * 100) / 100;

/** Places boxes left to right on a common baseline, with `gap` between them. */
export function row(boxes: Box[], gap = 0): Box {
	const result = emptyBox();
	let x = 0;
	for (const box of boxes) {
		if (result.prims.length || x > 0) x += gap;
		result.prims.push(...shifted(box.prims, x, 0));
		x += box.width;
		result.ascent = Math.max(result.ascent, box.ascent);
		result.descent = Math.max(result.descent, box.descent);
	}
	result.width = x;
	return result;
}

/** Places `box` with its baseline moved by `dy` (positive lowers) inside a parent box. */
export function placed(box: Box, dx: number, dy: number): Box {
	return {
		width: box.width,
		ascent: box.ascent - dy,
		descent: box.descent + dy,
		prims: shifted(box.prims, dx, dy),
	};
}

/** Union of boxes that share a baseline, each at its own x offset. */
export function overlay(parts: Array<{ box: Box; x: number; y: number }>): Box {
	const result = emptyBox();
	for (const { box, x, y } of parts) {
		result.prims.push(...shifted(box.prims, x, y));
		result.width = Math.max(result.width, x + box.width);
		result.ascent = Math.max(result.ascent, box.ascent - y);
		result.descent = Math.max(result.descent, box.descent + y);
	}
	return result;
}

const RELATIONS = new Set('=≠<>≤≥≈≡≪≫∈∉⊂⊃⊆⊇→←↔⇒⇐⇔∝≅∼≃'.split(''));
const BINARY = new Set('+−-×÷·±∓∪∩∧∨⊕⊗∘∗'.split(''));
export const isRelation = (ch: string) => RELATIONS.has(ch);
export const isBinary = (ch: string) => BINARY.has(ch);

/** Estimated advance of one character, in em. */
export function glyphEm(ch: string, bold: boolean): number {
	let em: number;
	if (/[iIjl|!.,;:'`]/.test(ch)) em = ch === 'I' ? 0.36 : 0.3;
	else if (/[ftr]/.test(ch)) em = 0.38;
	else if (/[mwMW@]/.test(ch)) em = ch === 'm' ? 0.82 : 0.86;
	else if (/[A-Z]/.test(ch)) em = 0.66;
	else if (/[a-z]/.test(ch)) em = 0.52;
	else if (/[0-9]/.test(ch)) em = 0.5;
	else if (/[()[\]{}]/.test(ch)) em = 0.34;
	else if (ch === ' ') em = 0.25;
	else if (isRelation(ch) || isBinary(ch)) em = 0.7;
	else if (/[Ͱ-Ͽ]/.test(ch)) em = 0.58;
	else if (/[⺀-鿿＀-￯]/.test(ch)) em = 1;
	else em = 0.6;
	return bold ? em * 1.08 : em;
}

export interface LayoutContext {
	layout(node: OmmlNode, style: MathStyle): Box;
	script(style: MathStyle): MathStyle;
}

export const isOn = (value: string) => value === '1' || value === 'on' || value === 'true';

/** A run of upright text as a box, optionally stretched vertically (large delimiters). */
export function glyph(text: string, size: number, stretch = 1): Box {
	const width = Array.from(text).reduce((sum, ch) => sum + glyphEm(ch, false), 0) * size;
	const box = emptyBox();
	box.width = width;
	box.ascent = size * 0.78 * stretch;
	box.descent = size * 0.22 * stretch;
	box.prims.push({
		kind: 'text',
		x: 0,
		y: 0,
		text,
		size,
		italic: false,
		bold: false,
		fit: width,
		...(stretch > 1 ? { stretch } : {}),
	});
	return box;
}

/** Base with a subscript and/or superscript set to its right. */
export function scripts(
	base: Box,
	sub: Box | undefined,
	sup: Box | undefined,
	style: MathStyle,
): Box {
	const raise = Math.max(style.size * 0.4, base.ascent - style.size * 0.3);
	const drop = Math.max(style.size * 0.2, base.descent + style.size * 0.12);
	const parts = [{ box: base, x: 0, y: 0 }];
	let width = base.width;
	if (sup) {
		const y = -Math.max(raise, sup.descent + style.size * 0.25);
		parts.push({ box: sup, x: base.width + 1, y });
		width = Math.max(width, base.width + 1 + sup.width);
	}
	if (sub) {
		const y = Math.max(drop, sub.ascent - style.size * 0.1);
		parts.push({ box: sub, x: base.width + 1, y });
		width = Math.max(width, base.width + 1 + sub.width);
	}
	const box = overlay(parts);
	box.width = width;
	return box;
}
