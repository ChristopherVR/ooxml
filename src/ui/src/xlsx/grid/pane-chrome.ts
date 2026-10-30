// Paints the row / column headers (frozen and scrolling parts), the select-all corner and the
// freeze-pane lines for one frame.
import type { GridMetrics, VisibleCells } from 'ooxml-core/xlsx';
import type { Selection } from 'ooxml-core/xlsx/ui';
import { place } from './dom';
import type { GridGeometry } from 'ooxml-core/xlsx/ui';
import type { HeaderLayer } from './headers';
import { colSelected, rowSelected } from 'ooxml-core/xlsx/ui';

/** Thickness of a split bar in pixels. */
export const SPLIT_BAR = 5;

export interface HeaderPane {
	box: HTMLDivElement;
	layer: HeaderLayer;
}

export interface ChromeNodes {
	colHeaders: { frozen: HeaderPane; scroll: HeaderPane };
	rowHeaders: { frozen: HeaderPane; scroll: HeaderPane };
	corner: HTMLDivElement;
	freezeH: HTMLDivElement;
	freezeV: HTMLDivElement;
}

type PaneSpec = [
	HeaderPane,
	number,
	number,
	number,
	number,
	readonly number[],
	number,
	'row' | 'col',
];

export function paintChrome(
	nodes: ChromeNodes,
	g: GridGeometry,
	metrics: GridMetrics,
	vis: VisibleCells,
	selection: Selection,
): void {
	const show = g.headerW > 0;
	const font = Math.max(6, Math.round(11 * (metrics.zoom / 100)));
	const colState = (c: number) => colSelected(selection, c);
	const rowState = (r: number) => rowSelected(selection, r);
	const fw = Math.min(g.frozenW, g.cellsW);
	const fh = Math.min(g.frozenH, g.cellsH);
	const { colHeaders, rowHeaders } = nodes;
	const panes: PaneSpec[] = [
		[colHeaders.frozen, g.headerW, 0, fw, g.headerH, vis.frozenCols, g.fixedOx, 'col'],
		[colHeaders.scroll, g.headerW + fw, 0, g.cellsW - fw, g.headerH, vis.cols, g.mainOx, 'col'],
		[rowHeaders.frozen, 0, g.headerH, g.headerW, fh, vis.frozenRows, g.fixedOy, 'row'],
		[rowHeaders.scroll, 0, g.headerH + fh, g.headerW, g.cellsH - fh, vis.rows, g.mainOy, 'row'],
	];
	for (const [pane, x, y, w, hgt, indices, offset, axis] of panes) {
		pane.box.hidden = !show || w <= 0 || hgt <= 0;
		if (pane.box.hidden) continue;
		place(pane.box, x, y, w, hgt);
		pane.layer.element.style.transform =
			axis === 'col' ? `translate3d(${-offset}px,0,0)` : `translate3d(0,${-offset}px,0)`;
		if (axis === 'col') pane.layer.element.style.height = `${hgt}px`;
		else pane.layer.element.style.width = `${w}px`;
		pane.layer.render(metrics, indices, axis === 'col' ? colState : rowState, font);
	}
	nodes.corner.hidden = !show;
	if (show) place(nodes.corner, 0, 0, g.headerW, g.headerH);
	const rowsFixed = g.split ? g.split.y > 0 : g.fRows > 0;
	const colsFixed = g.split ? g.split.x > 0 : g.fCols > 0;
	for (const bar of [nodes.freezeH, nodes.freezeV]) bar.classList.toggle('xg-split-bar', !!g.split);
	nodes.freezeH.hidden = !rowsFixed;
	nodes.freezeV.hidden = !colsFixed;
	// A split bar is wider than a freeze line so it can be grabbed.
	const t = g.split ? SPLIT_BAR : 1;
	const o = g.split ? Math.floor(SPLIT_BAR / 2) : 1;
	if (rowsFixed) place(nodes.freezeH, 0, g.headerH + g.frozenH - o, g.width, t);
	if (colsFixed) place(nodes.freezeV, g.headerW + g.frozenW - o, 0, t, g.height);
}
