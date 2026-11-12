// Opens the cell / row header / column header context menus from right click, Shift+F10 or the
// ContextMenu key, with the state the item builders need.
import { MAX_COL, MAX_ROW } from 'ooxml-core/xlsx';
import type { MenuEntry } from 'ooxml-core/xlsx/ui';
import { openContextMenu } from '../context-menu/index';
import { menuEntries } from '../context-menu/host-api';
import type { GridView } from './grid-view';
import { hostMenu } from '../customization';

export type MenuKind = 'cell' | 'row' | 'col';

/** Which menu the current selection calls for (whole rows / columns get the header menus). */
export function menuKindFor(view: GridView): MenuKind {
	const ranges = view.ctx.selection.get().ranges;
	if (
		ranges.length &&
		ranges.every(
			(r) =>
				r.start.col === 0 && r.end.col === MAX_COL && !(r.start.row === 0 && r.end.row === MAX_ROW),
		)
	)
		return 'row';
	if (
		ranges.length &&
		ranges.every(
			(r) =>
				r.start.row === 0 && r.end.row === MAX_ROW && !(r.start.col === 0 && r.end.col === MAX_COL),
		)
	)
		return 'col';
	return 'cell';
}

function hiddenInSelection(view: GridView, axis: 'row' | 'col'): boolean {
	const ranges = view.ctx.selection.get().ranges;
	const m = view.metrics;
	for (const r of ranges) {
		const from = axis === 'row' ? r.start.row : r.start.col;
		const to = Math.min(axis === 'row' ? r.end.row : r.end.col, from + 2000);
		for (let i = from; i <= to; i++)
			if (axis === 'row' ? m.isRowHidden(i) : m.isColHidden(i)) return true;
	}
	return false;
}

export function openGridMenu(
	view: GridView,
	kind: MenuKind,
	x: number,
	y: number,
	refocus: () => void,
): void {
	const ctx = view.ctx;
	const menu = kind === 'col' ? 'column' : kind;
	// The grid measures hidden rows and columns (including filtered ones) itself.
	const build = (): MenuEntry[] =>
		menuEntries(ctx, {
			kind: menu,
			sheet: view.sheetIndex(),
			...(kind === 'cell' ? {} : { hiddenInSelection: hiddenInSelection(view, kind) }),
		});
	openContextMenu(ctx, build(), x, y, {
		restoreFocus: refocus,
		offer: { input: { kind: menu, sheet: view.sheetIndex() }, build },
		...hostMenu(ctx, menu, view.sheetIndex()),
	});
}

/** Opens the menu for the active cell's position (keyboard). */
export function openMenuAtActive(view: GridView, refocus: () => void): void {
	const box = view.geometry.rangeBox(view.activeRange());
	const rect = view.view.getBoundingClientRect();
	openGridMenu(
		view,
		menuKindFor(view),
		rect.left + box.x + box.w / 2,
		rect.top + box.y + box.h,
		refocus,
	);
}

export function wireMenus(view: GridView, refocus: () => void): () => void {
	const onMenu = (event: MouseEvent) => {
		const target = event.target as Element;
		if (target.closest('.xg-editor-host.xg-editing, .xg-popup')) return;
		event.preventDefault();
		const hit = view.hitClient(event.clientX, event.clientY);
		const kind: MenuKind =
			hit.area === 'row-header' ? 'row' : hit.area === 'col-header' ? 'col' : 'cell';
		if (kind !== 'cell' && menuKindFor(view) !== kind) {
			const sel = view.ctx.selection.get();
			const within = sel.ranges.some((r) =>
				kind === 'row'
					? hit.row >= r.start.row && hit.row <= r.end.row
					: hit.col >= r.start.col && hit.col <= r.end.col,
			);
			if (!within) {
				const range =
					kind === 'row'
						? { start: { row: hit.row, col: 0 }, end: { row: hit.row, col: MAX_COL } }
						: { start: { row: 0, col: hit.col }, end: { row: MAX_ROW, col: hit.col } };
				view.ctx.selection.set({
					sheet: view.sheetIndex(),
					active: range.start,
					anchor: range.start,
					ranges: [range],
				});
			}
		}
		openGridMenu(view, kind, event.clientX, event.clientY, refocus);
	};
	view.view.addEventListener('contextmenu', onMenu);
	return () => view.view.removeEventListener('contextmenu', onMenu);
}
