/**
 * Excel-style status bar. Left: the cell mode (Ready / Edit), a read-only badge and the
 * compatibility notes button. Right: selection statistics from the core's `selectionStats`
 * (right-click to choose which, like Excel), the view buttons and the zoom controls. The shared
 * `office-ui-status-bar` and `office-ui-zoom-slider` draw it from translated state; this module
 * holds the model and routes activations to the editor.
 */
import { selectionStats } from 'ooxml-core/xlsx';
import { defineStatusBar, defineZoomSlider } from '../controls';
import type { OfficeStatusBarState } from '../controls';
import { registerIcon } from '../icons';
import type { EditorContext } from './context';
import { formatNumber, normalizeEditorLocale } from './localization';
import { el } from './ribbon/controls';
import { openMenu } from './ribbon/popover';

/** Excel's three workbook views on the shared 20px icon grid. */
const VIEW_ICONS = {
	normalView: 'M3 3h14v14H3z M3 8h14 M8 3v14',
	pageLayoutView: 'M5 2h10v16H5z M7.5 6h5 M7.5 9.5h5 M7.5 13h3',
	pageBreakView: 'M5 2h10v16H5z M2 10h3 M7.5 10h2 M11 10h2 M15 10h3',
};

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
	defineStatusBar();
	defineZoomSlider();
	for (const [name, d] of Object.entries(VIEW_ICONS)) registerIcon(name, d);
	const doc = ctx.host.ownerDocument;
	const element = doc.createElement('office-ui-status-bar') as HTMLElement & {
		state: OfficeStatusBarState;
	};
	element.className = 'xve-status';
	element.setAttribute('part', 'status-bar');
	element.setAttribute('role', 'status');

	// Statistics sit in the trailing slot, before the views, as in Excel; right-click picks them.
	const stats = el(doc, 'span', 'xve-status-stats');
	stats.slot = 'end';
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
	const slider = doc.createElement('office-ui-zoom-slider') as HTMLElement & { value: number };
	slider.slot = 'end';
	slider.setAttribute('min', String(ZOOM_MIN));
	slider.setAttribute('max', String(ZOOM_MAX));
	element.append(stats, slider);

	const zoom = () => ctx.grid()?.zoom() ?? 100;
	slider.addEventListener('input', () => handlers.setZoom(slider.value));
	element.addEventListener('office-status-activate', (event) => {
		if ((event as CustomEvent<{ id: string }>).detail.id === 'notes') handlers.showNotes();
	});

	const render = () => {
		const count = ctx.workbook()?.warnings.length ?? 0;
		const notesText = ctx.t(
			count === 1 ? '{count} compatibility note' : '{count} compatibility notes',
			{ count },
		);
		element.state = {
			items: [
				{ id: 'mode', text: ctx.t(MODE_LABELS[cellMode(ctx)]) },
				...(ctx.readOnly() ? [{ id: 'readonly', text: ctx.t('Read-only') }] : []),
			],
			toggles: [
				{ id: 'notes', icon: 'warning', label: notesText, text: notesText, hidden: count === 0 },
			],
			views: [
				{ id: 'normal', icon: 'normalView', label: ctx.t('Normal'), pressed: true },
				{
					id: 'pageLayout',
					icon: 'pageLayoutView',
					label: ctx.t('Page Layout (not available)'),
					disabled: true,
				},
				{
					id: 'pageBreak',
					icon: 'pageBreakView',
					label: ctx.t('Page Break Preview (not available)'),
					disabled: true,
				},
			],
		};
	};
	const refresh = () => {
		render();
		stats.textContent = statisticsText(ctx, shown);
		slider.value = zoom();
	};
	const relocalize = () => {
		slider.setAttribute('out-label', ctx.t('Zoom out'));
		slider.setAttribute('in-label', ctx.t('Zoom in'));
		slider.setAttribute('slider-label', ctx.t('Zoom level'));
		stats.title = ctx.t('Right-click to choose which statistics to show');
		refresh();
	};
	relocalize();
	return { element, refresh, relocalize };
}
