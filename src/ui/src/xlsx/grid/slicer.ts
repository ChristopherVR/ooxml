// Slicers over the grid. The core resolves the items, their states and the button geometry
// (`slicerItems`, `slicerGeometry`) and the style colours; this module only paints them and routes
// clicks to `session.setSlicerSelection`. Ctrl+click, Shift+click and the Multi-Select button map
// onto the core's `selectionAfterClick`; Clear Filter selects every item again.
import {
	selectionAfterClick,
	slicerGeometry,
	slicerItems,
	slicerStyleColors,
	type SlicerClick,
	type SlicerFrameObject,
	type SlicerItem,
	type Workbook,
} from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { h, svgNode } from './dom';

/** What the user toggled on a slicer; kept only in the open editor, as Excel does. */
export interface SlicerUiState {
	multi: boolean;
	anchor?: string;
}

const ICON_MULTI =
	'<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path fill="currentColor" d="M2 3h3v3H2zM7 4h7v1H7zM2 9h3v3H2zM7 10h7v1H7z"/></svg>';
const ICON_CLEAR =
	'<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path fill="currentColor" d="M2 2h12l-4.5 5.5V13l-3-1.5V7.5z"/><path stroke="#c42b1c" stroke-width="1.6" d="M9.5 9.5l5 5m0-5l-5 5"/></svg>';

export interface SlicerPaint {
	ctx: EditorContext;
	workbook: Workbook | undefined;
	state: SlicerUiState;
}

/** The slicer cache and items behind a frame, or undefined when it cannot be resolved. */
export function slicerOf(workbook: Workbook | undefined, frame: SlicerFrameObject) {
	const slicer = workbook?.slicers?.slicers.find((entry) => entry.name === frame.name);
	const cache = workbook?.slicers?.caches.find((entry) => entry.name === slicer?.cache);
	return workbook && slicer && cache
		? { slicer, cache, items: slicerItems(workbook, cache) }
		: undefined;
}

function iconButton(doc: Document, action: string, label: string, markup: string): HTMLElement {
	const button = h(doc, 'button', 'xg-sl-tool', {
		type: 'button',
		'data-slicer-action': action,
		'aria-label': label,
		title: label,
	});
	button.append(svgNode(doc, markup));
	return button;
}

/** Paints one slicer into its drawing node (header, tool buttons and item buttons). */
export function paintSlicer(
	node: HTMLElement,
	frame: SlicerFrameObject,
	width: number,
	height: number,
	paint: SlicerPaint,
): void {
	const { ctx, workbook, state } = paint;
	const doc = node.ownerDocument;
	node.classList.add('xg-slicer');
	const found = slicerOf(workbook, frame);
	if (!workbook || !found) {
		node.classList.add('xg-obj-unsupported');
		node.textContent = ctx.t('{name} (not shown)', { name: 'slicer' });
		return;
	}
	const { slicer, items } = found;
	const colors = slicerStyleColors(slicer.style, workbook.theme);
	const geometry = slicerGeometry(slicer, items, width, height);
	const style = node.style;
	style.setProperty('--sl-bg', colors.background);
	style.setProperty('--sl-border', colors.border);
	style.setProperty('--sl-head-bg', colors.headerBackground);
	style.setProperty('--sl-head-fg', colors.headerText);
	style.setProperty('--sl-head-line', colors.headerBorder);
	style.setProperty('--sl-sel-bg', colors.selectedFill);
	style.setProperty('--sl-sel-fg', colors.selectedText);
	style.setProperty('--sl-sel-line', colors.selectedBorder);
	style.setProperty('--sl-un-bg', colors.unselectedFill);
	style.setProperty('--sl-un-fg', colors.unselectedText);
	style.setProperty('--sl-un-line', colors.unselectedBorder);
	style.setProperty('--sl-nd-fg', colors.noDataText);
	style.setProperty('--sl-nd-bg', colors.noDataSelectedFill);
	node.setAttribute('role', 'group');
	const caption = slicer.caption ?? slicer.name;
	node.setAttribute('aria-label', caption);
	if (geometry.showHeader) {
		const head = h(doc, 'div', 'xg-sl-head');
		head.style.height = `${geometry.header.h}px`;
		const title = h(doc, 'span', 'xg-sl-caption');
		title.textContent = caption;
		const tools = h(doc, 'div', 'xg-sl-tools');
		const multi = iconButton(doc, 'multi', ctx.t('Multi-Select'), ICON_MULTI);
		multi.setAttribute('aria-pressed', String(state.multi));
		multi.classList.toggle('xg-sl-on', state.multi);
		const filtered = items.some((item) => !item.selected);
		const clear = iconButton(doc, 'clear', ctx.t('Clear Filter'), ICON_CLEAR);
		clear.toggleAttribute('disabled', !filtered);
		tools.append(multi, clear);
		head.append(title, tools);
		node.append(head);
	}
	const body = h(doc, 'div', 'xg-sl-body');
	body.style.left = `${geometry.body.x}px`;
	body.style.top = `${geometry.body.y}px`;
	body.style.width = `${geometry.body.w}px`;
	body.style.height = `${geometry.body.h}px`;
	const inner = h(doc, 'div', 'xg-sl-grid');
	inner.style.height = `${geometry.contentHeight}px`;
	for (const button of geometry.buttons) {
		const el = h(doc, 'button', `xg-sl-item xg-sl-${button.state}`, {
			type: 'button',
			'data-slicer-item': button.key,
			'aria-pressed': String(button.state.startsWith('selected')),
		});
		el.style.left = `${button.x - geometry.body.x}px`;
		el.style.top = `${button.y - geometry.body.y}px`;
		el.style.width = `${button.w}px`;
		el.style.height = `${button.h}px`;
		el.textContent = button.label;
		el.title = button.label;
		inner.append(el);
	}
	body.append(inner);
	node.append(body);
}

/** The keys a click on `key` leaves selected (`null` means every item). */
export function slicerClickResult(
	items: readonly SlicerItem[],
	key: string,
	event: Pick<PointerEvent, 'ctrlKey' | 'metaKey' | 'shiftKey'>,
	state: SlicerUiState,
): string[] | null {
	const mode: SlicerClick = event.shiftKey
		? 'range'
		: event.ctrlKey || event.metaKey || state.multi
			? 'toggle'
			: 'single';
	return selectionAfterClick(items, key, mode, state.anchor);
}
