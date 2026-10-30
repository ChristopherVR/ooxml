// Trace Precedents / Dependents arrows drawn over the grid: blue lines with a dot at the source and
// an arrowhead at the target, red when the source holds an error, and a dashed black arrow with a
// sheet icon when the other end is on another sheet. The arrows come from the core's `traceNext`
// (kept in the editor's audit state); this file only positions and paints them.
import type { TraceArrow } from 'ooxml-core/xlsx';
import { auditState, onAuditChange } from 'ooxml-core/xlsx/ui';
import { h } from './dom';
import type { GridView } from './grid-view';

export const AUDIT_CSS = `
.xga-layer{position:absolute;inset:0;pointer-events:none;z-index:3;overflow:hidden}
.xga-layer svg{position:absolute;inset:0;width:100%;height:100%;overflow:hidden}
`;

const SVG = 'http://www.w3.org/2000/svg';
const BLUE = '#1f4e9e';
const RED = '#c00000';
const BLACK = '#201f1e';
const ICON = 22;

type Box = { x: number; y: number; w: number; h: number };

export function installAuditArrows(view: GridView): () => void {
	const doc = view.doc;
	const layer = h(doc, 'div', 'xga-layer', { 'data-part': 'audit-arrows' });
	const svg = doc.createElementNS(SVG, 'svg');
	layer.append(svg);
	view.view.append(layer);

	const node = (name: string, attrs: Record<string, string | number>): SVGElement => {
		const n = doc.createElementNS(SVG, name);
		for (const [k, v] of Object.entries(attrs)) n.setAttribute(k, String(v));
		return n;
	};

	// A line from (x1,y1) ending at (x2,y2) with an arrowhead there.
	const arrow = (
		x1: number,
		y1: number,
		x2: number,
		y2: number,
		color: string,
		dashed: boolean,
	) => {
		const out: SVGElement[] = [];
		const line = node('line', { x1, y1, x2, y2, stroke: color, 'stroke-width': 1.5 });
		if (dashed) line.setAttribute('stroke-dasharray', '4 3');
		out.push(line);
		const angle = Math.atan2(y2 - y1, x2 - x1);
		const spread = 0.45;
		const size = 8;
		out.push(
			node('polygon', {
				points: `${x2},${y2} ${x2 - size * Math.cos(angle - spread)},${y2 - size * Math.sin(angle - spread)} ${x2 - size * Math.cos(angle + spread)},${y2 - size * Math.sin(angle + spread)}`,
				fill: color,
			}),
		);
		return out;
	};

	const sheetIcon = (x: number, y: number): SVGElement => {
		const g = node('g', {});
		g.append(
			node('rect', {
				x,
				y,
				width: ICON,
				height: ICON * 0.8,
				fill: '#fff',
				stroke: BLACK,
				'stroke-width': 1,
			}),
			node('path', {
				d: `M${x} ${y + 5}h${ICON}M${x} ${y + 10}h${ICON}M${x + 7} ${y}v${ICON * 0.8}M${x + 14} ${y}v${ICON * 0.8}`,
				stroke: BLACK,
				'stroke-width': 0.8,
				fill: 'none',
			}),
		);
		return g;
	};

	const visible = (b: Box): boolean => {
		const g = view.geometry;
		return b.x + b.w > g.headerW && b.y + b.h > g.headerH && b.x < g.width && b.y < g.height;
	};

	const draw = (a: TraceArrow, out: SVGElement[]): void => {
		const g = view.geometry;
		const sheet = view.sheetIndex();
		const color = a.error ? RED : BLUE;
		const toBox = g.rangeBox({ start: a.to, end: a.to });
		if (a.offSheet) {
			if (a.kind === 'precedents' && a.to.sheet === sheet) {
				if (!visible(toBox)) return;
				const cy = toBox.y + toBox.h / 2;
				const ix = Math.max(g.headerW + 2, toBox.x - 70);
				out.push(
					sheetIcon(ix, cy - ICON * 0.4),
					...arrow(ix + ICON, cy, toBox.x + 2, cy, BLACK, true),
				);
			} else if (a.kind === 'dependents' && a.from.sheet === sheet) {
				const fromBox = g.rangeBox(a.from.range);
				if (!visible(fromBox)) return;
				const cy = fromBox.y + fromBox.h / 2;
				const ix = fromBox.x + fromBox.w + 50;
				out.push(
					node('circle', { cx: fromBox.x + fromBox.w - 4, cy, r: 3, fill: BLACK }),
					...arrow(fromBox.x + fromBox.w - 4, cy, ix, cy, BLACK, true),
					sheetIcon(ix, cy - ICON * 0.4),
				);
			}
			return;
		}
		if (a.from.sheet !== sheet || a.to.sheet !== sheet) return;
		const fromBox = g.rangeBox(a.from.range);
		if (!visible(fromBox) && !visible(toBox)) return;
		const first = g.rangeBox({ start: a.from.range.start, end: a.from.range.start });
		const multi =
			a.from.range.start.row !== a.from.range.end.row ||
			a.from.range.start.col !== a.from.range.end.col;
		if (multi)
			out.push(
				node('rect', {
					x: fromBox.x,
					y: fromBox.y,
					width: fromBox.w,
					height: fromBox.h,
					fill: 'none',
					stroke: color,
					'stroke-width': 1.5,
				}),
			);
		const sx = first.x + Math.min(first.w, 12) / 2 + 2;
		const sy = first.y + first.h / 2;
		const tx = toBox.x + toBox.w / 2;
		const ty = toBox.y + toBox.h / 2;
		// Stop the head at the target's border so it stays visible.
		const dx = tx - sx;
		const dy = ty - sy;
		const len = Math.hypot(dx, dy) || 1;
		const trim = Math.min(toBox.w, toBox.h) / 2;
		const ex = len > trim ? tx - (dx / len) * trim : tx;
		const ey = len > trim ? ty - (dy / len) * trim : ty;
		out.push(
			node('circle', { cx: sx, cy: sy, r: 3, fill: color }),
			...arrow(sx, sy, ex, ey, color, false),
		);
	};

	const hook = (): void => {
		const arrows = auditState(view.ctx).arrows;
		if (!arrows.length) {
			if (svg.firstChild) svg.replaceChildren();
			return;
		}
		const out: SVGElement[] = [];
		for (const a of arrows) draw(a, out);
		svg.replaceChildren(...out);
	};
	view.hooks.add(hook);
	const off = onAuditChange(view.ctx, () => view.schedule());
	return () => {
		view.hooks.delete(hook);
		off();
		layer.remove();
	};
}
