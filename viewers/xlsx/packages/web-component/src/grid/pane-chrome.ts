// Paints the row / column headers (frozen and scrolling parts), the select-all corner and the
// freeze-pane lines for one frame.
import type { GridMetrics, VisibleCells } from '@christophervr/xlsx-core';
import type { Selection } from '../context.js';
import { place } from './dom.js';
import type { GridGeometry } from './geometry.js';
import type { HeaderLayer } from './headers.js';
import { colSelected, rowSelected } from './selection-ops.js';

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
		[colHeaders.frozen, g.headerW, 0, fw, g.headerH, vis.frozenCols, 0, 'col'],
		[
			colHeaders.scroll,
			g.headerW + fw,
			0,
			g.cellsW - fw,
			g.headerH,
			vis.cols,
			g.frozenW + g.scrollLeft,
			'col',
		],
		[rowHeaders.frozen, 0, g.headerH, g.headerW, fh, vis.frozenRows, 0, 'row'],
		[
			rowHeaders.scroll,
			0,
			g.headerH + fh,
			g.headerW,
			g.cellsH - fh,
			vis.rows,
			g.frozenH + g.scrollTop,
			'row',
		],
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
	nodes.freezeH.hidden = g.fRows === 0;
	nodes.freezeV.hidden = g.fCols === 0;
	if (g.fRows) place(nodes.freezeH, 0, g.headerH + g.frozenH - 1, g.width, 1);
	if (g.fCols) place(nodes.freezeV, g.headerW + g.frozenW - 1, 0, 1, g.height);
}
