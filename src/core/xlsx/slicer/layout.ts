import { SLICER_DEFAULT_ROW_HEIGHT } from '../model-slicers';
import { anchorToPixelBox, pixelBoxToAnchor } from '../layout/anchors';
import { createGridMetrics } from '../layout/metrics';
import type { DrawingAnchor, Slicer, Worksheet } from '../model';
import type { SlicerItem, SlicerItemState } from './items';

const EMU_PER_PIXEL = 9525;
/** Geometry constants of the painted slicer (CSS pixels at 100% zoom). */
export const SLICER_METRICS = {
	border: 1,
	padding: 5,
	gap: 3,
	headerHeight: 26,
	minButtonWidth: 24,
	scrollbar: 0,
} as const;

export interface SlicerButtonBox {
	key: string;
	label: string;
	state: SlicerItemState;
	index: number;
	x: number;
	y: number;
	w: number;
	h: number;
}

export interface SlicerGeometry {
	showHeader: boolean;
	header: { x: number; y: number; w: number; h: number };
	/** The two header buttons: Multi-Select and Clear Filter, right-aligned. */
	multiSelect: { x: number; y: number; w: number; h: number };
	clearFilter: { x: number; y: number; w: number; h: number };
	/** Area the buttons scroll in. */
	body: { x: number; y: number; w: number; h: number };
	columns: number;
	rowHeight: number;
	buttons: SlicerButtonBox[];
	/** Height of all rows; larger than `body.h` means vertical scrolling. */
	contentHeight: number;
}

export const emuToPixels = (emu: number): number => emu / EMU_PER_PIXEL;
export const pixelsToEmu = (px: number): number => Math.round(px * EMU_PER_PIXEL);

/** Buttons fill the grid column by column's worth: row-major, `columns` per row. */
export function slicerGeometry(
	slicer: Pick<Slicer, 'columnCount' | 'rowHeight' | 'showCaption'>,
	items: readonly SlicerItem[],
	width: number,
	height: number,
): SlicerGeometry {
	const m = SLICER_METRICS;
	const showHeader = slicer.showCaption !== false;
	const headerHeight = showHeader ? m.headerHeight : 0;
	const columns = Math.max(1, Math.min(20000, Math.floor(slicer.columnCount ?? 1)));
	const rowHeight = Math.max(8, emuToPixels(slicer.rowHeight ?? SLICER_DEFAULT_ROW_HEIGHT));
	const inner = { x: m.border + m.padding, w: Math.max(0, width - 2 * (m.border + m.padding)) };
	const body = {
		x: inner.x,
		y: m.border + headerHeight + (showHeader ? m.padding / 2 : m.padding),
		w: inner.w,
		h: Math.max(
			0,
			height - m.border - headerHeight - m.padding - (showHeader ? m.padding / 2 : m.padding),
		),
	};
	const buttonWidth = Math.max(
		m.minButtonWidth > inner.w ? inner.w : m.minButtonWidth,
		(inner.w - (columns - 1) * m.gap) / columns,
	);
	const rows = Math.ceil(items.length / columns);
	const buttons: SlicerButtonBox[] = items.map((item, index) => ({
		key: item.key,
		label: item.label,
		state: item.state,
		index,
		x: inner.x + (index % columns) * (buttonWidth + m.gap),
		y: body.y + Math.floor(index / columns) * (rowHeight + m.gap),
		w: buttonWidth,
		h: rowHeight,
	}));
	const icon = 20;
	return {
		showHeader,
		header: { x: m.border, y: m.border, w: Math.max(0, width - 2 * m.border), h: headerHeight },
		multiSelect: {
			x: width - m.border - m.padding - 2 * icon - 2,
			y: m.border + (headerHeight - icon) / 2,
			w: icon,
			h: icon,
		},
		clearFilter: {
			x: width - m.border - m.padding - icon,
			y: m.border + (headerHeight - icon) / 2,
			w: icon,
			h: icon,
		},
		body,
		columns,
		rowHeight,
		buttons,
		contentHeight: rows ? rows * (rowHeight + m.gap) - m.gap : 0,
	};
}

/** The button under a point of the slicer (relative to its top-left), if any. */
export function slicerHit(
	geometry: SlicerGeometry,
	x: number,
	y: number,
	scroll = 0,
): SlicerButtonBox | undefined {
	const body = geometry.body;
	if (x < body.x || x > body.x + body.w || y < body.y || y > body.y + body.h) return undefined;
	return geometry.buttons.find(
		(button) =>
			x >= button.x &&
			x <= button.x + button.w &&
			y + scroll >= button.y &&
			y + scroll <= button.y + button.h,
	);
}

/** The anchor of a frame resized to a width and/or height in pixels at 100% zoom (top-left kept). */
export function resizedFrameAnchor(
	sheet: Worksheet,
	anchor: DrawingAnchor,
	size: { width?: number | undefined; height?: number | undefined },
): DrawingAnchor {
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	const box = anchorToPixelBox(sheet, anchor, metrics);
	return pixelBoxToAnchor(
		sheet,
		{ ...box, w: Math.max(40, size.width ?? box.w), h: Math.max(40, size.height ?? box.h) },
		metrics,
		anchor,
	);
}
