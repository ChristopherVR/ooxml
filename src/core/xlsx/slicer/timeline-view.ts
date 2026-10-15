import type { PivotTable, Timeline, TimelineCache, Workbook, Worksheet } from '../model';
import {
	periodsInSelection,
	selectionLabel,
	timelinePeriods,
	type TimelinePeriod,
} from './timeline-periods';

export interface TimelineTarget {
	sheet: Worksheet;
	sheetIndex: number;
	pivot: PivotTable;
}

/** The PivotTable a timeline cache filters. */
export function timelineTarget(
	workbook: Workbook,
	cache: TimelineCache,
): TimelineTarget | undefined {
	const first = cache.pivots[0];
	if (!first) return undefined;
	for (const [sheetIndex, sheet] of workbook.sheets.entries()) {
		const pivot =
			sheet.sheetId === first.tabId
				? sheet.pivotTables?.find((entry) => entry.name === first.name)
				: undefined;
		if (pivot) return { sheet, sheetIndex, pivot };
	}
	return undefined;
}

export interface TimelineView {
	periods: TimelinePeriod[];
	/** First and last selected period (undefined: no range, everything is shown). */
	selected: [first: number, last: number] | undefined;
	label: string;
}

export function timelineView(
	cache: TimelineCache,
	timeline: Pick<Timeline, 'level'>,
	locale = 'en',
): TimelineView {
	const periods = timelinePeriods(timeline.level, cache.bounds.start, cache.bounds.end, locale);
	return {
		periods,
		selected: periodsInSelection(periods, cache.selection),
		label: selectionLabel(periods, cache.selection, timeline.level, locale),
	};
}

/** Button widths per level, in pixels. */
export const TIMELINE_PERIOD_WIDTH = [84, 64, 52, 26] as const;
export const TIMELINE_METRICS = {
	border: 1,
	headerHeight: 28,
	groupHeight: 20,
	cellHeight: 28,
	scrollbarHeight: 14,
	padding: 6,
} as const;

export interface TimelineGeometry {
	showHeader: boolean;
	header: { x: number; y: number; w: number; h: number };
	body: { x: number; y: number; w: number; h: number };
	groups: { label: string; x: number; w: number }[];
	cells: {
		index: number;
		label: string;
		x: number;
		y: number;
		w: number;
		h: number;
		selected: boolean;
	}[];
	contentWidth: number;
	/** Horizontal scroll that brings `scrollPosition` into view. */
	scrollLeft: number;
}

export function timelineGeometry(
	timeline: Pick<Timeline, 'level' | 'showHeader' | 'scrollPosition'>,
	view: TimelineView,
	width: number,
	height: number,
): TimelineGeometry {
	const m = TIMELINE_METRICS;
	const showHeader = timeline.showHeader !== false;
	const headerHeight = showHeader ? m.headerHeight : 0;
	const cellWidth = TIMELINE_PERIOD_WIDTH[Math.min(3, Math.max(0, timeline.level))] ?? 52;
	const top = m.border + headerHeight + m.padding;
	const cells = view.periods.map((period, i) => ({
		index: i,
		label: period.label,
		x: i * cellWidth,
		y: m.groupHeight,
		w: cellWidth,
		h: m.cellHeight,
		selected: !!view.selected && i >= view.selected[0] && i <= view.selected[1],
	}));
	const groups: TimelineGeometry['groups'] = [];
	view.periods.forEach((period, i) => {
		const last = groups.at(-1);
		if (last && last.label === period.group) last.w += cellWidth;
		else groups.push({ label: period.group, x: i * cellWidth, w: cellWidth });
	});
	const scrollTo = timeline.scrollPosition?.slice(0, 10);
	const at = scrollTo
		? view.periods.findIndex((period) => period.first <= scrollTo && period.last >= scrollTo)
		: -1;
	return {
		showHeader,
		header: { x: m.border, y: m.border, w: Math.max(0, width - 2 * m.border), h: headerHeight },
		body: {
			x: m.border + m.padding,
			y: top,
			w: Math.max(0, width - 2 * (m.border + m.padding)),
			h: Math.max(0, height - top - m.border),
		},
		groups,
		cells,
		contentWidth: view.periods.length * cellWidth,
		scrollLeft: Math.max(0, at) * cellWidth,
	};
}
