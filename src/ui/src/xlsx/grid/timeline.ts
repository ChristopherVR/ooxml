// Timelines over the grid. The core resolves the periods, the selected range and the geometry
// (`timelineView`, `timelineGeometry`); this module paints them and routes a click or drag over
// the period buttons to `session.setTimelineSelection`, the level list to `updateTimeline`.
import {
	TIMELINE_METRICS,
	TIMELINE_PERIOD_WIDTH,
	slicerStyleColors,
	timelineGeometry,
	timelineView,
	type SlicerFrameObject,
	type Workbook,
} from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { h, svgNode } from './dom';

const ICON_CLEAR =
	'<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path fill="currentColor" d="M2 2h12l-4.5 5.5V13l-3-1.5V7.5z"/><path stroke="#c42b1c" stroke-width="1.6" d="M9.5 9.5l5 5m0-5l-5 5"/></svg>';

const LEVELS: readonly [string, string][] = [
	['0', 'Years'],
	['1', 'Quarters'],
	['2', 'Months'],
	['3', 'Days'],
];

/** The timeline, its cache and the view behind a frame, or undefined when unresolved. */
export function timelineOf(
	workbook: Workbook | undefined,
	frame: SlicerFrameObject,
	locale = 'en',
) {
	const timeline = workbook?.slicers?.timelines.find((entry) => entry.name === frame.name);
	const cache = workbook?.slicers?.timelineCaches.find((entry) => entry.name === timeline?.cache);
	return workbook && timeline && cache
		? { timeline, cache, view: timelineView(cache, timeline, locale) }
		: undefined;
}

/** Paints one timeline (header, level list, period buttons and the selected range). */
export function paintTimeline(
	node: HTMLElement,
	frame: SlicerFrameObject,
	width: number,
	height: number,
	ctx: EditorContext,
	workbook: Workbook | undefined,
): void {
	const doc = node.ownerDocument;
	node.classList.add('xg-timeline');
	const found = timelineOf(workbook, frame, ctx.locale());
	if (!workbook || !found) {
		node.classList.add('xg-obj-unsupported');
		node.textContent = ctx.t('{name} (not shown)', { name: 'timeline' });
		return;
	}
	const { timeline, cache, view } = found;
	const colors = slicerStyleColors(timeline.style, workbook.theme);
	const geometry = timelineGeometry(timeline, view, width, height);
	node.style.setProperty('--sl-bg', colors.background);
	node.style.setProperty('--sl-border', colors.border);
	node.style.setProperty('--sl-head-line', colors.headerBorder);
	node.style.setProperty('--sl-sel-bg', colors.selectedFill);
	node.style.setProperty('--sl-sel-line', colors.selectedBorder);
	node.style.setProperty('--sl-un-bg', colors.unselectedFill);
	node.style.setProperty('--sl-un-fg', colors.unselectedText);
	node.style.setProperty('--sl-un-line', colors.unselectedBorder);
	node.style.setProperty('--sl-hover', colors.hoverFill);
	const caption = timeline.caption ?? timeline.name;
	node.setAttribute('role', 'group');
	node.setAttribute('aria-label', caption);
	if (geometry.showHeader) {
		const head = h(doc, 'div', 'xg-tl-head');
		head.style.height = `${geometry.header.h}px`;
		const title = h(doc, 'span', 'xg-tl-caption');
		title.textContent = caption;
		head.append(title);
		if (timeline.showSelectionLabel !== false) {
			const label = h(doc, 'span', 'xg-tl-selection', { 'aria-live': 'polite' });
			label.textContent = view.label || ctx.t('All Periods');
			head.append(label);
		}
		if (timeline.showTimeLevel !== false) {
			const level = h(doc, 'select', 'xg-tl-level', {
				'aria-label': ctx.t('Time Level'),
				'data-timeline-action': 'level',
			});
			for (const [value, label] of LEVELS) {
				const option = h(doc, 'option');
				option.value = value;
				option.textContent = ctx.t(label);
				level.append(option);
			}
			level.value = String(timeline.level);
			head.append(level);
		}
		const clear = h(doc, 'button', 'xg-sl-tool', {
			type: 'button',
			'data-timeline-action': 'clear',
			'aria-label': ctx.t('Clear Filter'),
			title: ctx.t('Clear Filter'),
		});
		clear.append(svgNode(doc, ICON_CLEAR));
		clear.toggleAttribute('disabled', !cache.selection);
		head.append(clear);
		node.append(head);
	}
	const body = h(doc, 'div', 'xg-tl-body');
	body.style.left = `${geometry.body.x}px`;
	body.style.top = `${geometry.body.y}px`;
	body.style.width = `${geometry.body.w}px`;
	body.style.height = `${geometry.body.h}px`;
	const track = h(doc, 'div', 'xg-tl-track', {
		'data-cell-width': String(
			TIMELINE_PERIOD_WIDTH[Math.min(3, Math.max(0, timeline.level))] ?? 52,
		),
	});
	track.style.width = `${geometry.contentWidth}px`;
	for (const group of geometry.groups) {
		const el = h(doc, 'div', 'xg-tl-group');
		el.style.left = `${group.x}px`;
		el.style.width = `${group.w}px`;
		el.style.height = `${TIMELINE_METRICS.groupHeight}px`;
		const text = h(doc, 'span');
		text.textContent = group.label;
		el.append(text);
		track.append(el);
	}
	for (const cell of geometry.cells) {
		const el = h(doc, 'button', `xg-tl-cell${cell.selected ? ' xg-tl-on' : ''}`, {
			type: 'button',
			'data-period': String(cell.index),
			'aria-pressed': String(cell.selected),
		});
		el.style.left = `${cell.x}px`;
		el.style.top = `${cell.y}px`;
		el.style.width = `${cell.w}px`;
		el.style.height = `${cell.h}px`;
		el.textContent = cell.label;
		const period = view.periods[cell.index];
		if (period)
			el.title = period.first === period.last ? period.first : `${period.first} - ${period.last}`;
		track.append(el);
	}
	body.append(track);
	node.append(body);
	body.scrollLeft = geometry.scrollLeft;
}

/** The period index under a pointer, from the track's position and cell width. */
export function periodAt(track: HTMLElement, clientX: number): number {
	const width = Number(track.dataset.cellWidth) || 52;
	return Math.floor((clientX - track.getBoundingClientRect().left) / width);
}
