// Collaborators' selections laid out for painting: one coloured outline per peer on the shown
// sheet, in the same plane pixels as the local selection, with a name tag. DOM-free; the Excel
// grid only places the boxes this returns (as Word draws remote carets with a name label).
import { formatCursorLabel } from '../../collab/identity';
import type { CellRange } from '../address';
import type { GridMetrics } from './metrics';

/** A collaborator's selection resolved to a local sheet (`ooxml-core/xlsx/collab` produces it). */
export interface RemoteRange {
	clientId: number;
	userName: string;
	userColor: string;
	/** Index of the sheet in the local workbook. */
	sheet: number;
	range: CellRange;
}

/** One remote outline in plane pixels (the zoomed cell area, origin at A1). */
export interface RemoteSelectionBox {
	clientId: number;
	x: number;
	y: number;
	w: number;
	h: number;
	color: string;
	/** The collaborator's name, shortened for the tag. */
	label: string;
	/** Full name, for the tooltip and the accessible label. */
	name: string;
	/** The tag sits above the outline, or inside its top edge when the outline starts on row 1. */
	tag: 'above' | 'inside';
}

export interface RemoteSelectionLayoutOptions {
	/** The sheet on screen; selections on other sheets are left out. */
	sheet: number;
	/** Last row / column worth painting (whole-column and whole-row ranges are clipped to it). */
	maxRow: number;
	maxCol: number;
	/** Height of the name tag in pixels; an outline closer to the top puts its tag inside. */
	tagHeight?: number;
	/** Longest tag text before it is shortened with an ellipsis. Default 20. */
	maxLabelChars?: number;
}

/** Outlines of the collaborators' selections on the shown sheet, ordered by client id. */
export function layoutRemoteSelections(
	metrics: GridMetrics,
	peers: readonly RemoteRange[],
	options: RemoteSelectionLayoutOptions,
): RemoteSelectionBox[] {
	const tagHeight = options.tagHeight ?? 16;
	const boxes: RemoteSelectionBox[] = [];
	for (const peer of peers) {
		if (peer.sheet !== options.sheet) continue;
		const { start, end } = peer.range;
		const top = Math.min(start.row, end.row);
		const left = Math.min(start.col, end.col);
		if (top > options.maxRow || left > options.maxCol) continue;
		const bottom = Math.min(Math.max(start.row, end.row), options.maxRow);
		const right = Math.min(Math.max(start.col, end.col), options.maxCol);
		const x = metrics.colLeft(left);
		const y = metrics.rowTop(top);
		const w = metrics.colLeft(right + 1) - x;
		const h = metrics.rowTop(bottom + 1) - y;
		// Hidden rows or columns only: nothing on screen to outline.
		if (w <= 0 || h <= 0) continue;
		boxes.push({
			clientId: peer.clientId,
			x,
			y,
			w,
			h,
			color: peer.userColor,
			label: formatCursorLabel(peer.userName, options.maxLabelChars),
			name: peer.userName,
			tag: y < tagHeight ? 'inside' : 'above',
		});
	}
	return boxes.sort((a, b) => a.clientId - b.clientId);
}
