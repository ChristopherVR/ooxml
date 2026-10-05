// Sort & Filter (shared by Home > Editing and the Data tab): sort A to Z / Z to A, custom sort,
// the auto-filter toggle, clear and reapply.
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { icon } from './icons.js';
import { editing, guessHeader, regionOf, tableAt, target } from './util.js';

function sortBy(ctx: EditorContext, descending: boolean): void {
	const t = target(ctx);
	if (!t) return;
	const col = t.active.col;
	const filter = t.ws.autoFilter?.range;
	const inFilter =
		filter &&
		t.active.row >= filter.start.row &&
		t.active.row <= filter.end.row &&
		col >= filter.start.col &&
		col <= filter.end.col;
	if (tableAt(t.ws, t.active) || inFilter) return t.session.sortByColumn(t.sheet, col, descending);
	const region = regionOf(t);
	t.session.sort(t.sheet, region, [{ col, descending }], guessHeader(t.ws, region));
}

const filtered = (ctx: EditorContext) => target(ctx)?.ws.autoFilter?.columns ?? [];

export function sortCommands(): Command[] {
	return [
		editing({
			id: 'data.sort-asc',
			label: 'Sort A to Z',
			icon: icon('sortAsc'),
			lock: 'sort',
			run: (ctx) => sortBy(ctx, false),
		}),
		editing({
			id: 'data.sort-desc',
			label: 'Sort Z to A',
			icon: icon('sortDesc'),
			lock: 'sort',
			run: (ctx) => sortBy(ctx, true),
		}),
		editing({
			id: 'data.sort-custom',
			label: 'Custom Sort...',
			icon: icon('sortCustom'),
			lock: 'sort',
			run: (ctx) => void ctx.dialogs.open('sort'),
		}),
		editing({
			id: 'data.filter',
			label: 'Filter',
			icon: icon('filter'),
			shortcut: 'Ctrl+Shift+L',
			lock: 'autoFilter',
			checked: (ctx) => !!target(ctx)?.ws.autoFilter,
			run: (ctx) => {
				const t = target(ctx);
				if (!t) return;
				t.session.setAutoFilter(t.sheet, t.ws.autoFilter ? undefined : regionOf(t));
			},
		}),
		editing({
			id: 'data.filter-clear',
			label: 'Clear',
			icon: icon('filterClear'),
			lock: 'autoFilter',
			enabled: (ctx) => filtered(ctx).length > 0,
			run: (ctx) => {
				const t = target(ctx);
				const filter = t?.ws.autoFilter;
				if (!t || !filter) return;
				t.session.batch('Clear filter', () => {
					for (const fc of filter.columns ?? [])
						t.session.filterColumn(t.sheet, filter.range.start.col + fc.offset, undefined);
				});
			},
		}),
		editing({
			id: 'data.filter-reapply',
			label: 'Reapply',
			icon: icon('reapply'),
			lock: 'autoFilter',
			enabled: (ctx) => filtered(ctx).length > 0,
			run: (ctx) => {
				const t = target(ctx);
				const filter = t?.ws.autoFilter;
				if (!t || !filter) return;
				const columns = structuredClone(filter.columns ?? []);
				t.session.batch('Reapply', () => {
					for (const fc of columns)
						t.session.filterColumn(t.sheet, filter.range.start.col + fc.offset, [
							...(fc.values ?? []),
							...(fc.blank ? [''] : []),
						]);
				});
			},
		}),
	];
}
