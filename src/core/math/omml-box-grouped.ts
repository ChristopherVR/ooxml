/** N-ary operators, delimiters, matrices and accents/limits: constructs that wrap or stack content. */
import {
	AXIS,
	emptyBox,
	glyph,
	isOn,
	overlay,
	row,
	scripts,
	shifted,
	type Box,
	type LayoutContext,
	type MathStyle,
} from './omml-box';
import { child, ensureArray, val } from './omml-mathml-helpers';
import type { OmmlNode } from './omml-node';

const NARY = new Set(['∑', '∏', '∐', '⋀', '⋁', '⋂', '⋃']);

export function nary(node: OmmlNode, style: MathStyle, ctx: LayoutContext): Box {
	const pr = child(node, 'm:naryPr');
	const chr = val(child(pr, 'm:chr')) || '∫';
	const big = NARY.has(chr) ? 1.4 : 1.6;
	const op = glyph(chr, style.size * big);
	// Display-style operators reserve the tall extent Word gives them (measured against Word).
	op.ascent = style.size * (big > 1.5 ? 1.17 : 1.15);
	op.descent = style.size * (big > 1.5 ? 0.55 : 0.6);
	const limits = (val(child(pr, 'm:limLoc')) || (NARY.has(chr) ? 'undOvr' : 'subSup')) === 'undOvr';
	const sub = isOn(val(child(pr, 'm:subHide')))
		? undefined
		: ctx.layout(child(node, 'm:sub'), ctx.script(style));
	const sup = isOn(val(child(pr, 'm:supHide')))
		? undefined
		: ctx.layout(child(node, 'm:sup'), ctx.script(style));
	const hasSub = Boolean(sub?.width);
	const hasSup = Boolean(sup?.width);
	let head: Box;
	if (limits && (hasSub || hasSup)) {
		const width = Math.max(op.width, sub?.width ?? 0, sup?.width ?? 0);
		const parts = [{ box: op, x: (width - op.width) / 2, y: 0 }];
		if (hasSup)
			parts.push({
				box: sup!,
				x: (width - sup!.width) / 2,
				y: -(op.ascent + sup!.descent + style.size * 0.075),
			});
		if (hasSub)
			parts.push({
				box: sub!,
				x: (width - sub!.width) / 2,
				y: op.descent + sub!.ascent + style.size * 0.075,
			});
		head = overlay(parts);
		head.width = width;
	} else head = scripts(op, hasSub ? sub : undefined, hasSup ? sup : undefined, style);
	const body = ctx.layout(child(node, 'm:e'), style);
	return row([head, body], style.size * 0.12);
}

const MATCHING_OPEN: Record<string, string> = {
	'(': '(',
	'[': '[',
	'{': '{',
	'|': '|',
	'‖': '‖',
	'⟨': '⟨',
	'⌈': '⌈',
	'⌊': '⌊',
};

export function delimiter(node: OmmlNode, style: MathStyle, ctx: LayoutContext): Box {
	const pr = child(node, 'm:dPr');
	const begin = child(pr, 'm:begChr')['@_val'] === undefined ? '(' : val(child(pr, 'm:begChr'));
	const end = child(pr, 'm:endChr')['@_val'] === undefined ? ')' : val(child(pr, 'm:endChr'));
	const separator = val(child(pr, 'm:sepChr')) || '|';
	const cells = ensureArray(node['m:e']).map((cell) => ctx.layout(cell, style));
	const parts: Box[] = [];
	cells.forEach((cell, index) => {
		if (index > 0) parts.push(glyph(separator, style.size));
		parts.push(cell);
	});
	const content = row(parts, style.size * 0.05);
	const tall = (content.ascent + content.descent) / style.size;
	const stretch = Math.max(1, tall / 1.0);
	const open = begin ? glyph(MATCHING_OPEN[begin] ?? begin, style.size, stretch) : emptyBox();
	const close = end ? glyph(end, style.size, stretch) : emptyBox();
	return row([open, content, close], style.size * 0.04);
}

export function matrix(node: OmmlNode, style: MathStyle, ctx: LayoutContext): Box {
	const rows = ensureArray(node['m:mr']).map((mr) =>
		ensureArray(mr['m:e']).map((cell) => ctx.layout(cell, style)),
	);
	const columns = Math.max(0, ...rows.map((cells) => cells.length));
	const widths = Array.from({ length: columns }, (_, c) =>
		Math.max(0, ...rows.map((cells) => cells[c]?.width ?? 0)),
	);
	const gapX = style.size * 0.8;
	const gapY = 0;
	const result = emptyBox();
	let y = 0;
	const heights = rows.map((cells) => ({
		a: Math.max(style.size * 0.78, ...cells.map((cell) => cell.ascent)),
		d: Math.max(style.size * 0.22, ...cells.map((cell) => cell.descent)),
	}));
	const total =
		heights.reduce((sum, h) => sum + h.a + h.d, 0) + gapY * Math.max(0, rows.length - 1);
	y = -(total / 2) + style.size * AXIS;
	rows.forEach((cells, r) => {
		const h = heights[r]!;
		y += h.a;
		let x = 0;
		cells.forEach((cell, c) => {
			result.prims.push(...shifted(cell.prims, x + (widths[c]! - cell.width) / 2, y));
			x += widths[c]! + gapX;
		});
		y += h.d + gapY;
	});
	result.width = widths.reduce((sum, w) => sum + w, 0) + gapX * Math.max(0, columns - 1);
	result.ascent = total / 2 + style.size * AXIS;
	result.descent = total / 2 - style.size * AXIS;
	return result;
}

export function above(base: Box, mark: Box, style: MathStyle): Box {
	const width = Math.max(base.width, mark.width);
	const box = overlay([
		{ box: base, x: (width - base.width) / 2, y: 0 },
		{
			box: mark,
			x: (width - mark.width) / 2,
			y: -(base.ascent + mark.descent + style.size * 0.04),
		},
	]);
	box.width = width;
	return box;
}
export function below(base: Box, mark: Box, style: MathStyle): Box {
	const width = Math.max(base.width, mark.width);
	const box = overlay([
		{ box: base, x: (width - base.width) / 2, y: 0 },
		{ box: mark, x: (width - mark.width) / 2, y: base.descent + mark.ascent + style.size * 0.04 },
	]);
	box.width = width;
	return box;
}
