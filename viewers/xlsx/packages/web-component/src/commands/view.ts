// View tab: Normal view, Show (formula bar, gridlines, headings), Zoom and Freeze Panes. Also the
// Help tab commands (keyboard shortcuts, feature status).
import { createGridMetrics } from '@christophervr/xlsx-core';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { icon } from './icons.js';
import { target, viewing } from './util.js';

export const currentZoom = (ctx: EditorContext): number =>
	ctx.grid()?.zoom() ?? target(ctx)?.ws.view.zoom ?? 100;

export function setZoom(ctx: EditorContext, percent: number): void {
	const value = Math.max(10, Math.min(400, Math.round(percent)));
	const grid = ctx.grid();
	if (grid) grid.setZoom(value);
	else {
		const t = target(ctx);
		if (t) t.session.setSheetView(t.sheet, { zoom: value });
	}
	ctx.requestRender();
}

/** The zoom at which the selection fills the visible grid (10 to 400 percent). */
export function zoomToFit(ctx: EditorContext): number | undefined {
	const t = target(ctx);
	if (!t) return undefined;
	const pane = ctx.root.querySelector?.<HTMLElement>('[part="grid"]');
	const width = pane?.clientWidth || 800;
	const height = pane?.clientHeight || 500;
	const metrics = createGridMetrics(t.ws, { zoom: 100 });
	const r = t.range;
	const w = metrics.colLeft(r.end.col) + metrics.colWidth(r.end.col) - metrics.colLeft(r.start.col);
	const h = metrics.rowTop(r.end.row) + metrics.rowHeight(r.end.row) - metrics.rowTop(r.start.row);
	// Leave room for the row and column headers.
	const fit = Math.min((width - 40) / Math.max(1, w), (height - 24) / Math.max(1, h));
	return Math.max(10, Math.min(400, Math.floor(fit * 100)));
}

const frozen = (ctx: EditorContext): boolean => !!target(ctx)?.ws.view.freeze;

function freeze(ctx: EditorContext, rows: number, cols: number): void {
	const t = target(ctx);
	if (!t) return;
	t.session.setFreeze(t.sheet, rows || cols ? { rows, cols } : undefined);
}

interface FormulaBarHost {
	showFormulaBar: boolean;
}
const formulaBarHost = (ctx: EditorContext): FormulaBarHost | undefined =>
	'showFormulaBar' in ctx.host ? (ctx.host as unknown as FormulaBarHost) : undefined;

export function viewCommands(): Command[] {
	return [
		viewing({
			id: 'view.normal',
			label: 'Normal',
			icon: icon('normalView'),
			checked: () => true,
			run: () => undefined,
		}),
		{
			id: 'view.formula-bar',
			label: 'Formula Bar',
			icon: icon('formulaBar'),
			enabled: (ctx) => !!formulaBarHost(ctx),
			checked: (ctx) => formulaBarHost(ctx)?.showFormulaBar !== false,
			run: (ctx) => {
				const host = formulaBarHost(ctx);
				if (host) host.showFormulaBar = !host.showFormulaBar;
			},
		},
		viewing({
			id: 'view.zoom',
			label: 'Zoom',
			icon: icon('zoom'),
			run: (ctx) => void ctx.dialogs.open('zoom'),
		}),
		viewing({
			id: 'view.zoom-100',
			label: '100%',
			icon: icon('zoom100'),
			checked: (ctx) => currentZoom(ctx) === 100,
			run: (ctx) => setZoom(ctx, 100),
		}),
		viewing({
			id: 'view.zoom-selection',
			label: 'Zoom to Selection',
			icon: icon('zoomSelection'),
			run: (ctx) => {
				const zoom = zoomToFit(ctx);
				const t = target(ctx);
				if (zoom === undefined || !t) return;
				setZoom(ctx, zoom);
				ctx.grid()?.scrollTo(t.range.start);
			},
		}),
		viewing({
			id: 'view.freeze-panes',
			label: 'Freeze Panes',
			icon: icon('freeze'),
			checked: frozen,
			run: (ctx) => {
				const t = target(ctx);
				if (!t) return;
				if (frozen(ctx)) return freeze(ctx, 0, 0);
				// Everything above and left of the active cell (Excel measures from the top-left visible
				// cell; the model's freeze counts from row and column 0).
				freeze(ctx, t.active.row, t.active.col);
			},
		}),
		viewing({
			id: 'view.freeze-top-row',
			label: 'Freeze Top Row',
			icon: icon('freeze'),
			run: (ctx) => freeze(ctx, 1, 0),
		}),
		viewing({
			id: 'view.freeze-first-column',
			label: 'Freeze First Column',
			icon: icon('freeze'),
			run: (ctx) => freeze(ctx, 0, 1),
		}),
		viewing({
			id: 'view.unfreeze',
			label: 'Unfreeze Panes',
			icon: icon('freeze'),
			enabled: frozen,
			run: (ctx) => freeze(ctx, 0, 0),
		}),
		{
			id: 'help.shortcuts',
			label: 'Keyboard Shortcuts',
			icon: icon('keyboard'),
			shortcut: 'Ctrl+/',
			run: async (ctx) => {
				// The shell owns the shortcut help panel; use its command when it registered one.
				for (const id of [
					'shell.shortcut-help',
					'keyboard.shortcut-help',
					'help.keyboard-shortcuts',
				])
					if (ctx.commands.get(id)) return void (await ctx.commands.run(id));
				void ctx.dialogs.open('shortcut-help');
			},
		},
		{
			id: 'help.feature-status',
			label: 'About and Feature Status',
			icon: icon('info'),
			run: (ctx) => void ctx.dialogs.open('feature-status'),
		},
	];
}
