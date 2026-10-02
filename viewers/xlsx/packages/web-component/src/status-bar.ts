/**
 * Excel-style status bar. Left: the cell mode (Ready / Edit), a read-only badge and the
 * compatibility notes button. Right: selection statistics from the core's `selectionStats`
 * (right-click to choose which, like Excel), the view buttons and the zoom controls. Layout
 * mirrors docx-viewer's status-bar.ts.
 */
import { selectionStats } from '@christophervr/xlsx-core';
import type { EditorContext } from './context';
import { formatNumber, normalizeEditorLocale } from './localization';
import { el } from './ribbon/controls';
import { ribbonIcon } from './ribbon/icons';
import { openMenu } from './ribbon/popover';

export const ZOOM_MIN = 10;
export const ZOOM_MAX = 400;

export type StatKey = 'average' | 'count' | 'numericCount' | 'min' | 'max' | 'sum';
const STAT_LABELS: Record<StatKey, string> = {
	average: 'Average',
	count: 'Count',
	numericCount: 'Numerical Count',
	min: 'Minimum',
	max: 'Maximum',
	sum: 'Sum',
};
const STAT_ORDER: StatKey[] = ['average', 'count', 'numericCount', 'min', 'max', 'sum'];

const MODE_LABELS = { ready: 'Ready', enter: 'Enter', edit: 'Edit', point: 'Point' } as const;
type CellMode = keyof typeof MODE_LABELS;

/** Excel's mode indicator: the grid's own `mode()` when it has one, else Ready / Edit. */
export function cellMode(ctx: EditorContext): CellMode {
	const grid = ctx.grid() as (ReturnType<EditorContext['grid']> & { mode?(): string }) | undefined;
	const mode = grid?.mode?.();
	if (mode && mode in MODE_LABELS) return mode as CellMode;
	return grid?.isEditing() ? 'edit' : 'ready';
}

export interface StatusBarHandlers {
	showNotes(): void;
	setZoom(percent: number): void;
}

export interface StatusBar {
	readonly element: HTMLElement;
	refresh(): void;
	relocalize(): void;
}

/** The statistics text for the current selection, or '' when Excel shows none (one cell). */
export function statisticsText(ctx: EditorContext, shown: ReadonlySet<StatKey>): string {
	const workbook = ctx.workbook();
	const selection = ctx.selection.get();
	// A host may hand over any object as `workbook` before it loads; only a real sheet has stats.
	if (!workbook?.sheets?.[selection.sheet]?.rows) return '';
	const stats = selectionStats(workbook, selection.sheet, selection.ranges);
	if (stats.count < 2) return '';
	const locale = normalizeEditorLocale(ctx.locale());
	const parts: string[] = [];
	for (const key of STAT_ORDER) {
		if (!shown.has(key)) continue;
		const value =
			key === 'count'
				? stats.count
				: key === 'numericCount'
					? stats.numericCount
					: key === 'sum'
						? stats.numericCount
							? stats.sum
							: undefined
						: stats[key];
		if (value === undefined) continue;
		parts.push(`${ctx.t(STAT_LABELS[key])}: ${formatNumber(locale, value)}`);
	}
	return parts.join('    ');
}

export function createStatusBar(ctx: EditorContext, handlers: StatusBarHandlers): StatusBar {
	const doc = ctx.host.ownerDocument;
	const element = el(doc, 'footer', 'xve-status');
	element.setAttribute('part', 'status-bar');
	element.setAttribute('role', 'status');
	const left = el(doc, 'div', 'xve-status-left');
	const mode = el(doc, 'span', 'xve-status-mode');
	const readOnly = el(doc, 'span', 'xve-status-badge');
	const notes = el(doc, 'button', 'xve-status-notes');
	notes.type = 'button';
	const notesText = el(doc, 'span');
	notes.append(ribbonIcon(doc, 'warning', 14), notesText);
	notes.addEventListener('click', () => handlers.showNotes());
	left.append(mode, readOnly, notes);

	const right = el(doc, 'div', 'xve-status-right');
	const stats = el(doc, 'span', 'xve-status-stats');
	const shown = new Set<StatKey>(['average', 'count', 'sum']);
	stats.addEventListener('contextmenu', (event) => {
		event.preventDefault();
		openMenu(
			stats,
			STAT_ORDER.map((key) => ({
				label: ctx.t(STAT_LABELS[key]),
				checked: shown.has(key),
				run: () => {
					if (shown.has(key)) shown.delete(key);
					else shown.add(key);
					refresh();
				},
			})),
			ctx.t('Customize Status Bar'),
		);
	});
	const button = (icon: string) => {
		const node = el(doc, 'button', 'xve-icon-button');
		node.type = 'button';
		node.append(ribbonIcon(doc, icon, 16));
		return node;
	};
	const normal = button('normalView');
	normal.setAttribute('aria-pressed', 'true');
	const pageLayout = button('pageLayoutView');
	const pageBreak = button('pageBreakView');
	pageLayout.disabled = true;
	pageBreak.disabled = true;
	const zoomOut = button('minus');
	const zoomIn = button('plus');
	const slider = el(doc, 'input', 'xve-zoom-slider');
	slider.type = 'range';
	slider.min = String(ZOOM_MIN);
	slider.max = String(ZOOM_MAX);
	slider.step = '10';
	const percent = el(doc, 'button', 'xve-zoom-percent');
	percent.type = 'button';
	const zoom = () => ctx.grid()?.zoom() ?? 100;
	const clamp = (value: number) => Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(value)));
	zoomOut.addEventListener('click', () => handlers.setZoom(clamp(zoom() - 10)));
	zoomIn.addEventListener('click', () => handlers.setZoom(clamp(zoom() + 10)));
	slider.addEventListener('input', () => handlers.setZoom(clamp(Number(slider.value))));
	percent.addEventListener('click', () => {
		if (ctx.commands.get('view.zoom')) void ctx.commands.run('view.zoom');
		else handlers.setZoom(100);
	});
	right.append(stats, normal, pageLayout, pageBreak, zoomOut, slider, zoomIn, percent);
	element.append(left, right);

	const label = (node: HTMLElement, text: string) => {
		node.setAttribute('aria-label', ctx.t(text));
		node.title = ctx.t(text);
	};
	const refresh = () => {
		mode.textContent = ctx.t(MODE_LABELS[cellMode(ctx)]);
		readOnly.hidden = !ctx.readOnly();
		const count = ctx.workbook()?.warnings.length ?? 0;
		notes.hidden = count === 0;
		notesText.textContent = ctx.t(
			count === 1 ? '{count} compatibility note' : '{count} compatibility notes',
			{ count },
		);
		stats.textContent = statisticsText(ctx, shown);
		const value = zoom();
		slider.value = String(value);
		percent.textContent = `${value}%`;
	};
	const relocalize = () => {
		readOnly.textContent = ctx.t('Read-only');
		label(normal, 'Normal');
		label(pageLayout, 'Page Layout (not available)');
		label(pageBreak, 'Page Break Preview (not available)');
		label(zoomOut, 'Zoom out');
		label(zoomIn, 'Zoom in');
		label(slider, 'Zoom level');
		label(percent, 'Zoom');
		stats.title = ctx.t('Right-click to choose which statistics to show');
		refresh();
	};
	relocalize();
	return { element, refresh, relocalize };
}
