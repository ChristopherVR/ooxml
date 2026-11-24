/**
 * Ruled tables from a page's rules: thin lines and rectangles become horizontal and vertical
 * edges, edges that touch form a table, their positions cluster into the grid, and a missing rule
 * between two grid units merges them into one cell. Filled rectangles inside a cell shade it.
 */
import type { PdfPageRule } from '../../pdf/index';

export interface GridBorder {
	/** Points. */
	width: number;
	/** `#rrggbb`. */
	color: string;
}

export interface GridCell {
	row: number;
	col: number;
	rowSpan: number;
	colSpan: number;
	fill?: string;
	borders: { top?: GridBorder; left?: GridBorder; bottom?: GridBorder; right?: GridBorder };
}

export interface RuledGrid {
	/** Column edges, left to right. */
	xs: number[];
	/** Row edges, top to bottom. */
	ys: number[];
	cells: GridCell[];
}

interface Edge {
	/** y of a horizontal edge, x of a vertical one. */
	pos: number;
	from: number;
	to: number;
	width: number;
	color: string;
	rules: PdfPageRule[];
}

/** Edges thicker than this are shapes, not rules. */
const THICK = 6;
const TOLERANCE = 2.5;
const MAX_PAIRS = 4_000_000;

function classify(rules: readonly PdfPageRule[]) {
	const horizontal: Edge[] = [];
	const vertical: Edge[] = [];
	const fills: PdfPageRule[] = [];
	const junctions: PdfPageRule[] = [];
	const edge = (pos: number, from: number, to: number, width: number, rule: PdfPageRule): Edge => ({
		pos,
		from,
		to,
		width,
		color: rule.color,
		rules: [rule],
	});
	for (const rule of rules) {
		const w = rule.x1 - rule.x0;
		const h = rule.y1 - rule.y0;
		const stroke = rule.stroked ? Math.max(0.25, rule.lineWidth) : 0;
		if (rule.kind === 'line') {
			if (!rule.stroked) continue;
			if (h < 0.5 && w >= 2) horizontal.push(edge(rule.y0, rule.x0, rule.x1, stroke, rule));
			else if (w < 0.5 && h >= 2) vertical.push(edge(rule.x0, rule.y0, rule.y1, stroke, rule));
			continue;
		}
		if (rule.filled && h <= THICK && w > 2 * h && w >= 2)
			horizontal.push(edge((rule.y0 + rule.y1) / 2, rule.x0, rule.x1, Math.max(h, stroke), rule));
		else if (rule.filled && w <= THICK && h > 2 * w && h >= 2)
			vertical.push(edge((rule.x0 + rule.x1) / 2, rule.y0, rule.y1, Math.max(w, stroke), rule));
		else if (rule.filled && w <= THICK && h <= THICK) junctions.push(rule);
		else if (rule.stroked && w > THICK && h > THICK) {
			// A stroked rectangle is four edges (its fill, if any, is a shading candidate).
			horizontal.push(edge(rule.y0, rule.x0, rule.x1, stroke, rule));
			horizontal.push(edge(rule.y1, rule.x0, rule.x1, stroke, rule));
			vertical.push(edge(rule.x0, rule.y0, rule.y1, stroke, rule));
			vertical.push(edge(rule.x1, rule.y0, rule.y1, stroke, rule));
			if (rule.filled) fills.push(rule);
		} else if (rule.filled && w > THICK && h > THICK) fills.push(rule);
	}
	return { horizontal: merge(horizontal), vertical: merge(vertical), fills, junctions };
}

/** Joins collinear pieces of one colour that touch (producers draw a rule cell by cell). */
function merge(edges: Edge[]): Edge[] {
	edges.sort((a, b) => a.pos - b.pos || a.from - b.from);
	const out: Edge[] = [];
	for (const edge of edges) {
		let same: Edge | undefined;
		for (let at = out.length - 1; at >= 0 && !same; at--) {
			const last = out[at]!;
			if (last.pos < edge.pos - 1.5) break;
			if (
				Math.abs(last.pos - edge.pos) <= 1.5 &&
				edge.from <= last.to + 2 &&
				last.color === edge.color &&
				Math.abs(last.width - edge.width) < 0.3
			)
				same = last;
		}
		if (same) {
			same.to = Math.max(same.to, edge.to);
			same.rules.push(...edge.rules);
		} else out.push({ ...edge, rules: [...edge.rules] });
	}
	return out;
}

function cluster(values: number[]): number[] {
	const sorted = [...values].sort((a, b) => a - b);
	const groups: number[][] = [];
	for (const value of sorted) {
		const last = groups[groups.length - 1];
		if (last && value - last[last.length - 1]! <= TOLERANCE) last.push(value);
		else groups.push([value]);
	}
	return groups.map((group) => group.reduce((sum, value) => sum + value, 0) / group.length);
}

const touches = (h: Edge, v: Edge) =>
	v.pos >= h.from - TOLERANCE &&
	v.pos <= h.to + TOLERANCE &&
	h.pos >= v.from - TOLERANCE &&
	h.pos <= v.to + TOLERANCE;

function components(horizontal: Edge[], vertical: Edge[]): { h: Edge[]; v: Edge[] }[] {
	const parent = [...horizontal, ...vertical].map((_, index) => index);
	const find = (index: number): number => {
		while (parent[index] !== index) index = parent[index] = parent[parent[index]!]!;
		return index;
	};
	horizontal.forEach((h, i) =>
		vertical.forEach((v, j) => {
			if (touches(h, v)) parent[find(i)] = find(horizontal.length + j);
		}),
	);
	const groups = new Map<number, { h: Edge[]; v: Edge[] }>();
	horizontal.forEach((h, i) => {
		const group = groups.get(find(i)) ?? { h: [], v: [] };
		group.h.push(h);
		groups.set(find(i), group);
	});
	vertical.forEach((v, j) => {
		const group = groups.get(find(horizontal.length + j)) ?? { h: [], v: [] };
		group.v.push(v);
		groups.set(find(horizontal.length + j), group);
	});
	return [...groups.values()];
}

const covering = (edges: Edge[], pos: number, at: number) =>
	edges.find(
		(edge) => Math.abs(edge.pos - pos) <= TOLERANCE && edge.from - 1 <= at && edge.to + 1 >= at,
	);

function buildGrid(h: Edge[], v: Edge[], fills: PdfPageRule[]): RuledGrid | undefined {
	const xs = cluster([...v.map((edge) => edge.pos), ...h.flatMap((edge) => [edge.from, edge.to])]);
	const ys = cluster([...h.map((edge) => edge.pos), ...v.flatMap((edge) => [edge.from, edge.to])]);
	const cols = xs.length - 1;
	const rows = ys.length - 1;
	// More units than any real table has: a drawing (and too costly to resolve cell by cell).
	if (cols < 2 || rows < 1 || rows * cols > 5000) return undefined;
	const parent = Array.from({ length: rows * cols }, (_, index) => index);
	const find = (index: number): number => {
		while (parent[index] !== index) index = parent[index] = parent[parent[index]!]!;
		return index;
	};
	const midY = (r: number) => (ys[r]! + ys[r + 1]!) / 2;
	const midX = (c: number) => (xs[c]! + xs[c + 1]!) / 2;
	for (let r = 0; r < rows; r++)
		for (let c = 0; c < cols; c++) {
			if (c + 1 < cols && !covering(v, xs[c + 1]!, midY(r)))
				parent[find(r * cols + c)] = find(r * cols + c + 1);
			if (r + 1 < rows && !covering(h, ys[r + 1]!, midX(c)))
				parent[find(r * cols + c)] = find((r + 1) * cols + c);
		}
	const regions = new Map<number, number[]>();
	for (let index = 0; index < rows * cols; index++) {
		const list = regions.get(find(index)) ?? [];
		list.push(index);
		regions.set(find(index), list);
	}
	const cells: GridCell[] = [];
	for (const units of regions.values()) {
		const rs = units.map((unit) => Math.floor(unit / cols));
		const cs = units.map((unit) => unit % cols);
		const [r0, r1, c0, c1] = [Math.min(...rs), Math.max(...rs), Math.min(...cs), Math.max(...cs)];
		// A region that is not a rectangle cannot be one Word cell: keep its units apart.
		const spans =
			(r1 - r0 + 1) * (c1 - c0 + 1) === units.length
				? [{ r0, r1, c0, c1 }]
				: units.map((unit) => {
						const r = Math.floor(unit / cols);
						const c = unit % cols;
						return { r0: r, r1: r, c0: c, c1: c };
					});
		for (const span of spans) cells.push(cellOf(span, xs, ys, h, v, fills));
	}
	cells.sort((a, b) => a.row - b.row || a.col - b.col);
	return { xs, ys, cells };
}

function cellOf(
	span: { r0: number; r1: number; c0: number; c1: number },
	xs: number[],
	ys: number[],
	h: Edge[],
	v: Edge[],
	fills: PdfPageRule[],
): GridCell {
	const [x0, x1, y0, y1] = [xs[span.c0]!, xs[span.c1 + 1]!, ys[span.r0]!, ys[span.r1 + 1]!];
	const border = (edge: Edge | undefined): GridBorder | undefined =>
		edge ? { width: edge.width, color: edge.color } : undefined;
	const cell: GridCell = {
		row: span.r0,
		col: span.c0,
		rowSpan: span.r1 - span.r0 + 1,
		colSpan: span.c1 - span.c0 + 1,
		borders: {},
	};
	const top = border(covering(h, y0, (x0 + x1) / 2));
	const bottom = border(covering(h, y1, (x0 + x1) / 2));
	const left = border(covering(v, x0, (y0 + y1) / 2));
	const right = border(covering(v, x1, (y0 + y1) / 2));
	if (top) cell.borders.top = top;
	if (bottom) cell.borders.bottom = bottom;
	if (left) cell.borders.left = left;
	if (right) cell.borders.right = right;
	const area = (x1 - x0) * (y1 - y0);
	let best = 0;
	for (const fill of fills) {
		const overlap =
			Math.max(0, Math.min(x1, fill.x1) - Math.max(x0, fill.x0)) *
			Math.max(0, Math.min(y1, fill.y1) - Math.max(y0, fill.y0));
		if (overlap > area * 0.5 && overlap > best && fill.color.toLowerCase() !== '#ffffff') {
			best = overlap;
			cell.fill = fill.color;
		}
	}
	return cell;
}

/** The page's ruled tables, and the rules they do not use (underlines, shapes). */
export function findGrids(rules: readonly PdfPageRule[]): {
	grids: RuledGrid[];
	rest: PdfPageRule[];
} {
	const { horizontal, vertical, fills, junctions } = classify(rules);
	const grids: RuledGrid[] = [];
	// A page of hatching (thousands of crossing strokes) is a drawing, and too costly to pair up.
	if (horizontal.length * vertical.length > MAX_PAIRS) return { grids, rest: [...rules] };
	const used = new Set<PdfPageRule>();
	for (const group of components(horizontal, vertical)) {
		if (group.h.length < 2 || group.v.length < 2) continue;
		const grid = buildGrid(group.h, group.v, fills);
		if (!grid) continue;
		grids.push(grid);
		for (const edge of [...group.h, ...group.v]) for (const rule of edge.rules) used.add(rule);
		const [x0, x1] = [grid.xs[0]! - 1, grid.xs[grid.xs.length - 1]! + 1];
		const [y0, y1] = [grid.ys[0]! - 1, grid.ys[grid.ys.length - 1]! + 1];
		for (const rule of [...fills, ...junctions])
			if (rule.x0 >= x0 && rule.x1 <= x1 && rule.y0 >= y0 && rule.y1 <= y1) used.add(rule);
	}
	grids.sort((a, b) => a.ys[0]! - b.ys[0]!);
	return { grids, rest: rules.filter((rule) => !used.has(rule)) };
}
