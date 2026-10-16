/**
 * Axis lines with the line formats a plain SVG line cannot draw: gradient
 * lines, compound (double, thick-thin, triple) lines and begin / end arrows.
 *
 * The axis builders tag each axis line with its axis id (`SvgLine.axisLine`);
 * this pass draws the decorated line into an effect image placed above the
 * gridlines and below the marks, hides the plain line (kept for hit-testing),
 * and strips the tags so the bindings never see them.
 *
 * @module render/chart-format/chart-format-render-lines
 */
import type { PptxChartData, PptxChartElementLine, PptxChartLineEndType } from 'ooxml-core/pptx';

import type { ChartViewModel, SvgLine, SvgPrimitive } from '../chart-view-model-types';
import { compoundStrokes, linePaintOf, ptToPx } from './chart-format-effects-svg';

const r2 = (v: number) => Math.round(v * 100) / 100;

/** Whether a line format needs more than a plain SVG line. */
export function needsLineDecoration(line: PptxChartElementLine | undefined): boolean {
	return Boolean(
		line &&
		(line.kind === 'gradient' ||
			(line.compound && line.compound !== 'sng') ||
			(line.beginArrow && line.beginArrow !== 'none') ||
			(line.endArrow && line.endArrow !== 'none')),
	);
}

/** An arrow head at `tip`, pointing along `(ux, uy)`, for a line `width` px wide. */
function arrowHead(
	type: PptxChartLineEndType,
	tip: { x: number; y: number },
	ux: number,
	uy: number,
	width: number,
	paint: string,
): string {
	// PowerPoint's medium arrow: three line widths long and wide (never tiny on hairlines).
	const length = Math.max(6, width * 3);
	const half = Math.max(3, width * 1.5);
	const px = -uy;
	const py = ux;
	const at = (along: number, across: number) =>
		`${r2(tip.x + ux * along + px * across)},${r2(tip.y + uy * along + py * across)}`;
	switch (type) {
		case 'triangle':
			return `<polygon points="${at(0, 0)} ${at(-length, half)} ${at(-length, -half)}" fill="${paint}"/>`;
		case 'stealth':
			return `<polygon points="${at(0, 0)} ${at(-length, half)} ${at(-length * 0.6, 0)} ${at(-length, -half)}" fill="${paint}"/>`;
		case 'arrow':
			return `<polyline points="${at(-length, half)} ${at(0, 0)} ${at(-length, -half)}" fill="none" stroke="${paint}" stroke-width="${r2(width)}" stroke-linejoin="miter"/>`;
		case 'diamond':
			return `<polygon points="${at(half, 0)} ${at(0, half)} ${at(-half, 0)} ${at(0, -half)}" fill="${paint}"/>`;
		case 'oval':
			return `<circle cx="${r2(tip.x)}" cy="${r2(tip.y)}" r="${r2(half)}" fill="${paint}"/>`;
		default:
			return '';
	}
}

/** How far a line is shortened at an end so a filled head's tip lands on the end. */
function inset(type: PptxChartLineEndType | undefined, width: number): number {
	return type === 'triangle' || type === 'stealth' ? Math.max(6, width * 3) * 0.7 : 0;
}

/**
 * The decorated line `from` -> `to` (begin -> end) as markup: gradient or
 * solid paint, compound strokes and arrow heads.
 */
export function decoratedLineMarkup(
	from: { x: number; y: number },
	to: { x: number; y: number },
	line: PptxChartElementLine,
	fallbackColor: string,
	fallbackWidth: number,
): string {
	const length = Math.hypot(to.x - from.x, to.y - from.y);
	if (length === 0) return '';
	const ux = (to.x - from.x) / length;
	const uy = (to.y - from.y) / length;
	const width = line.width !== undefined ? ptToPx(line.width) : fallbackWidth;
	const box = {
		x: Math.min(from.x, to.x),
		y: Math.min(from.y, to.y),
		w: Math.max(Math.abs(to.x - from.x), 0.01),
		h: Math.max(Math.abs(to.y - from.y), 0.01),
	};
	const stroke = linePaintOf(line, box) ?? { defs: '', paint: fallbackColor, opacity: 1 };
	const head = inset(line.beginArrow, width);
	const tail = inset(line.endArrow, width);
	const a = { x: from.x + ux * head, y: from.y + uy * head };
	const b = { x: to.x - ux * tail, y: to.y - uy * tail };
	const strokes = compoundStrokes(line.compound)
		.map(([share, offset]) => {
			const ox = -uy * offset * width;
			const oy = ux * offset * width;
			return `<line x1="${r2(a.x + ox)}" y1="${r2(a.y + oy)}" x2="${r2(b.x + ox)}" y2="${r2(b.y + oy)}" stroke="${stroke.paint}" stroke-width="${r2(width * share)}"/>`;
		})
		.join('');
	const heads =
		(line.beginArrow ? arrowHead(line.beginArrow, from, -ux, -uy, width, stroke.paint) : '') +
		(line.endArrow ? arrowHead(line.endArrow, to, ux, uy, width, stroke.paint) : '');
	return `${stroke.defs}<g opacity="${r2(stroke.opacity)}">${strokes}${heads}</g>`;
}

const isTagged = (prim: SvgPrimitive | SvgLine): prim is SvgLine & { axisLine: number } =>
	prim.kind === 'line' && typeof (prim as SvgLine).axisLine === 'number';

function untag<T extends SvgPrimitive | SvgLine>(prim: T, hide: boolean): T {
	if (!isTagged(prim)) return prim;
	const { axisLine: _tag, ...rest } = prim;
	return (hide ? { ...rest, opacity: 0 } : rest) as T;
}

/**
 * Draw the decorated axis lines of `data` (markup returned, original lines
 * hidden) and strip every axis-line tag from `vm`. Returns the markup to
 * paint above the gridlines, or `''`.
 */
export function decorateAxisLines(
	vm: ChartViewModel,
	data: PptxChartData | undefined,
): { vm: ChartViewModel; markup: string } {
	const decorated = new Map<number, PptxChartElementLine>();
	(data?.axes ?? []).forEach((axis, index) => {
		const format = data?.elementFormats?.[`axis:${index}`];
		if (axis.axisId !== undefined && format && !format.reset && needsLineDecoration(format.line)) {
			decorated.set(axis.axisId, format.line!);
		}
	});
	const lists = [vm.gridlines, vm.secondaryGridlines ?? [], vm.primitives];
	if (!lists.some((list) => list.some(isTagged))) {
		return { vm, markup: '' };
	}
	let markup = '';
	for (const list of lists) {
		for (const prim of list) {
			if (!isTagged(prim)) continue;
			const line = decorated.get(prim.axisLine);
			if (!line) continue;
			const axis = data?.axes?.find((candidate) => candidate.axisId === prim.axisLine);
			const vertical = prim.x1 === prim.x2;
			const reversed = axis?.orientation === 'maxMin';
			// Begin is the axis minimum: the bottom of a vertical axis, the left of a horizontal one.
			const low = vertical
				? { x: prim.x1, y: Math.max(prim.y1, prim.y2) }
				: { x: Math.min(prim.x1, prim.x2), y: prim.y1 };
			const high = vertical
				? { x: prim.x1, y: Math.min(prim.y1, prim.y2) }
				: { x: Math.max(prim.x1, prim.x2), y: prim.y1 };
			const [from, to] = reversed ? [high, low] : [low, high];
			markup += decoratedLineMarkup(from, to, line, prim.stroke, prim.strokeWidth);
		}
	}
	const hidden = (prim: SvgPrimitive | SvgLine) => isTagged(prim) && decorated.has(prim.axisLine);
	const next: ChartViewModel = {
		...vm,
		gridlines: vm.gridlines.map((line) => untag(line, hidden(line))),
		primitives: vm.primitives.map((prim) => untag(prim, hidden(prim))),
		...(vm.secondaryGridlines
			? { secondaryGridlines: vm.secondaryGridlines.map((line) => untag(line, hidden(line))) }
			: {}),
	};
	return { vm: next, markup };
}
